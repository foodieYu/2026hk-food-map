/**
 * 香港三日食地圖：試算表後台
 * 部署成「網頁應用程式」後，網頁可以：
 *   action=resolve  展開 Google 地圖連結，回傳店名、地址、座標
 *   action=add      新增一列店家
 *   action=update   修改某一列的任何欄位（編號除外）
 *   action=delete   刪除某一列
 */
const SHEET_ID = '1V6pm0Qq2RDSsaduO5p4c1hOy9V2_aJjzhmDVqiwUL1Q';
const EXTRA_COLUMNS = ['標籤'];
const TEXT_COLUMNS = ['開門', '打烊', '週末開門', '週末打烊', '電話', '地鐵出口'];

function doGet(e) {
  const p = (e && e.parameter) || {};
  let out;
  try {
    if (p.action === 'resolve') out = resolvePlace_(p.url || '');
    else if (p.action === 'add') out = addShop_(JSON.parse(p.data || '{}'));
    else if (p.action === 'update') out = updateShop_(JSON.parse(p.data || '{}'));
    else if (p.action === 'delete') out = deleteShop_(JSON.parse(p.data || '{}'));
    else out = { ok: true, msg: 'ready' };
  } catch (err) {
    out = { ok: false, error: String(err && err.message || err) };
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

function resolvePlace_(input) {
  const m = String(input).match(/https?:\/\/\S+/);
  let url = m ? m[0] : '';
  const text = String(input).replace(/https?:\/\/\S+/g, ' ').trim();
  // 展開短網址
  for (let i = 0; i < 6 && url && /goo\.gl|maps\.app/.test(url); i++) {
    const r = UrlFetchApp.fetch(url, { followRedirects: false, muteHttpExceptions: true });
    const h = r.getHeaders();
    const loc = h.Location || h.location;
    if (!loc) break;
    url = loc;
  }
  let dec = url;
  try { dec = decodeURIComponent(url.replace(/\+/g, ' ')); } catch (e) {}
  let name = '', q = '', lat = null, lng = null;
  let mm = dec.match(/\/place\/([^/@?]+)/);
  if (mm) name = mm[1].trim();
  mm = dec.match(/!3d(-?\d+\.\d+)!4d(-?\d+\.\d+)/) || dec.match(/@(-?\d+\.\d+),(-?\d+\.\d+)/) || dec.match(/[?&](?:q|ll|query)=(-?\d+\.\d+),\s*(-?\d+\.\d+)/);
  if (mm) { lat = +mm[1]; lng = +mm[2]; }
  mm = dec.match(/[?&](?:q|query)=([^&]+)/);
  if (mm && !/^-?\d+\.\d+,/.test(mm[1])) q = mm[1].trim();
  if (!name && text) name = text.split(/\n/)[0].trim();

  let address = '';
  const geocoder = Maps.newGeocoder().setRegion('hk').setLanguage('zh-TW');
  if (lat == null && (q || name)) {
    const g = geocoder.geocode(q || (name + ' 香港'));
    if (g.status === 'OK' && g.results.length) {
      lat = g.results[0].geometry.location.lat;
      lng = g.results[0].geometry.location.lng;
      address = g.results[0].formatted_address;
    }
  }
  if (lat != null && !address) {
    const g = geocoder.reverseGeocode(lat, lng);
    if (g.status === 'OK' && g.results.length) address = g.results[0].formatted_address;
  }
  // q 通常是「地址＋店名」，把店名拆出來
  if (!name && q && address) name = q.replace(address, '').trim();
  if (!name && q) name = q;
  return { ok: true, name: name, address: cleanAddress_(address), lat: lat, lng: lng, url: m ? m[0] : '', expanded: url };
}

function cleanAddress_(a) {
  return String(a || '').replace(/^香港\s*/, '').replace(/^(香港島|九龍|新界)\s*/, '');
}

function sheet_() {
  const sh = SpreadsheetApp.openById(SHEET_ID).getSheets()[0];
  let headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);
  EXTRA_COLUMNS.forEach(function (c) {
    if (headers.indexOf(c) < 0) {
      sh.getRange(1, headers.length + 1).setValue(c);
      headers.push(c);
    }
  });
  return { sh: sh, headers: headers };
}

function addShop_(d) {
  if (!d || !d['店名']) return { ok: false, error: '缺少店名' };
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const s = sheet_();
    d['編號'] = d['編號'] || ('u' + new Date().getTime().toString(36));
    const row = s.headers.map(function (h) {
      const v = d[h];
      if (v === undefined || v === null) return '';
      if (TEXT_COLUMNS.indexOf(h) >= 0 && v !== '') return "'" + v;
      return v;
    });
    s.sh.appendRow(row);
    return { ok: true, id: d['編號'], row: s.sh.getLastRow() };
  } finally {
    lock.releaseLock();
  }
}

function checkRow_(s, d) {
  const idCol = s.headers.indexOf('編號');
  const row = Number(d.row);
  if (!row || row < 2 || row > s.sh.getLastRow()) throw new Error('找不到這一列，請按更新後再試');
  const cur = String(s.sh.getRange(row, idCol + 1).getValue());
  if (d.id && !/^row\d+$/.test(d.id) && cur !== d.id) throw new Error('資料已變動，請按更新後再試');
  return row;
}

function updateShop_(d) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const s = sheet_();
    const row = checkRow_(s, d);
    Object.keys(d.fields || {}).forEach(function (k) {
      if (k === '編號') return;
      const c = s.headers.indexOf(k);
      if (c < 0) return;
      let v = d.fields[k];
      if (v === null || v === undefined) v = '';
      if (TEXT_COLUMNS.indexOf(k) >= 0 && v !== '') v = "'" + v;
      s.sh.getRange(row, c + 1).setValue(v);
    });
    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}

function deleteShop_(d) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const s = sheet_();
    const row = checkRow_(s, d);
    s.sh.deleteRow(row);
    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}
