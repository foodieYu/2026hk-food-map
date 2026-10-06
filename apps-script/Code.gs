/**
 * 旅遊地圖：試算表後台
 * 每個旅程是一份獨立的 Google 試算表（第一個分頁放地點，「設定」分頁放標題、同行的人等），
 * 另有一份「旅程總表」記錄所有旅程，給主畫面用。
 *
 * 部署成「網頁應用程式」（執行身分：我，存取：所有人）後，網頁可以：
 *   旅程頁（帶 trip=試算表ID，沒帶時是香港那份）
 *     action=resolve  展開 Google 地圖連結，回傳店名、地址、座標、區域
 *     action=add / update / delete  新增、修改、刪除地點
 *     action=config   修改旅程設定（例如標題）
 *   主畫面（要密碼）
 *     action=pinStatus  是否已設定密碼
 *     action=setPin     第一次設定密碼
 *     action=list       列出所有旅程
 *     action=create     建立新旅程
 */
const LEGACY_ID = '1V6pm0Qq2RDSsaduO5p4c1hOy9V2_aJjzhmDVqiwUL1Q'; // 香港三日食地圖
const SHOP_HEADERS = ['編號', '店名', '類型', '區域', '地址', '緯度', '經度', 'Google 地圖連結', '開門', '打烊', '公休日',
  '週末開門', '週末打烊', '營業時間說明', '必吃', '備註', '地鐵出口', '電話', '來源', '狀態', '哪一天', '新增者', '標籤', '想吃的人', '吃過的人'];
const EXTRA_COLUMNS = ['標籤', '想吃的人', '吃過的人'];
// 這些欄位要存成文字（避免電話變成數字）；時間欄位不加，舊的香港試算表才會存成真正的時間
const TEXT_COLUMNS = ['電話', '地鐵出口'];
const CONFIG_SHEET = '設定';
const CONFIG_KEYS = ['標題', '國家', '城市', '地區代碼', '時區', '開始日期', '結束日期', '天數', '同行的人', '每日備註', '港鐵', '中心緯度', '中心經度'];
const EDITABLE_CONFIG = ['標題', '開始日期', '結束日期', '天數', '同行的人', '每日備註'];
const REGISTRY_HEADERS = ['試算表ID', '標題', '國家', '城市', '開始日期', '結束日期', '同行的人', '建立時間'];
const FOLDER_NAME = '旅遊地圖';
const LEGACY_CONFIG = {
  '標題': '香港三日食地圖', '國家': '香港', '城市': '香港', '地區代碼': 'hk', '時區': 'Asia/Hong_Kong',
  '開始日期': '', '結束日期': '', '天數': '3', '同行的人': '瑜、翎、芷、佩', '每日備註': '2=迪士尼', '港鐵': '是',
  '中心緯度': '22.30', '中心經度': '114.172'
};

function doGet(e) {
  const p = (e && e.parameter) || {};
  let out;
  try {
    const data = p.data ? JSON.parse(p.data) : {};
    switch (p.action) {
      case 'resolve': out = resolvePlace_(p.url || '', p.region || '', p.city || ''); break;
      case 'add': out = addShop_(p.trip, data); break;
      case 'update': out = updateShop_(p.trip, data); break;
      case 'delete': out = deleteShop_(p.trip, data); break;
      case 'config': out = setConfig_(p.trip, data); break;
      case 'pinStatus': out = { ok: true, hasPin: !!props_().getProperty('PIN_HASH') }; break;
      case 'setPin': out = setPin_(p.pin); break;
      case 'list': checkPin_(p.pin); out = listTrips_(); break;
      case 'create': checkPin_(p.pin); out = createTrip_(data); break;
      default: out = { ok: true, msg: 'ready' };
    }
  } catch (err) {
    out = { ok: false, error: String(err && err.message || err) };
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

/** 在 Apps Script 編輯器手動執行一次，用來授權並建立旅程總表 */
function setup() {
  const reg = registry_();
  Logger.log('旅程總表：' + reg.getParent().getUrl());
}

// ---------------- 密碼 ----------------
function props_() { return PropertiesService.getScriptProperties(); }
function hash_(s) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, 'trip-map:' + String(s))
    .map(function (b) { return ('0' + (b & 255).toString(16)).slice(-2); }).join('');
}
function setPin_(pin) {
  if (props_().getProperty('PIN_HASH')) throw new Error('已經設定過密碼');
  if (!pin || String(pin).length < 4) throw new Error('密碼至少 4 個字');
  props_().setProperty('PIN_HASH', hash_(pin));
  return { ok: true };
}
function checkPin_(pin) {
  const h = props_().getProperty('PIN_HASH');
  if (!h) throw new Error('NO_PIN');
  if (!pin || hash_(pin) !== h) throw new Error('密碼錯誤');
}

// ---------------- 旅程總表 ----------------
function folder_() {
  const id = props_().getProperty('FOLDER_ID');
  if (id) { try { return DriveApp.getFolderById(id); } catch (e) {} }
  const f = DriveApp.createFolder(FOLDER_NAME);
  props_().setProperty('FOLDER_ID', f.getId());
  return f;
}
function registry_() {
  const id = props_().getProperty('REGISTRY_ID');
  if (id) { try { return SpreadsheetApp.openById(id).getSheets()[0]; } catch (e) {} }
  const ss = SpreadsheetApp.create('旅遊地圖・旅程總表');
  DriveApp.getFileById(ss.getId()).moveTo(folder_());
  const sh = ss.getSheets()[0];
  sh.setName('旅程');
  sh.getRange(1, 1, 1, REGISTRY_HEADERS.length).setValues([REGISTRY_HEADERS]).setFontWeight('bold');
  sh.getRange('A:G').setNumberFormat('@');
  sh.setFrozenRows(1);
  props_().setProperty('REGISTRY_ID', ss.getId());
  // 把原本的香港旅程登記進去
  const hk = SpreadsheetApp.openById(LEGACY_ID);
  const cfg = ensureConfig_(hk, LEGACY_CONFIG);
  sh.appendRow([LEGACY_ID, cfg['標題'], cfg['國家'], cfg['城市'], cfg['開始日期'], cfg['結束日期'], cfg['同行的人'], new Date()]);
  return sh;
}
function registeredIds_() {
  const sh = registry_();
  const n = sh.getLastRow();
  if (n < 2) return [];
  return sh.getRange(2, 1, n - 1, 1).getValues().map(function (r) { return String(r[0]); });
}
function listTrips_() {
  const sh = registry_();
  const n = sh.getLastRow();
  const rows = n < 2 ? [] : sh.getRange(2, 1, n - 1, REGISTRY_HEADERS.length).getValues();
  const trips = rows.filter(function (r) { return r[0]; }).map(function (r) {
    const t = { id: String(r[0]), title: r[1], country: r[2], city: r[3], start: fmtDate_(r[4]), end: fmtDate_(r[5]), people: r[6], created: r[7] ? new Date(r[7]).getTime() : 0, places: 0, done: 0, days: '' };
    try {
      const ss = SpreadsheetApp.openById(t.id);
      const cfg = readConfig_(ss, t.id === LEGACY_ID ? LEGACY_CONFIG : {});
      t.title = cfg['標題'] || t.title; t.people = cfg['同行的人'] || t.people;
      t.start = cfg['開始日期'] || t.start; t.end = cfg['結束日期'] || t.end; t.days = cfg['天數'] || '';
      const sh2 = ss.getSheets()[0];
      const last = sh2.getLastRow();
      if (last > 1) {
        const head = sh2.getRange(1, 1, 1, sh2.getLastColumn()).getValues()[0].map(String);
        const vals = sh2.getRange(2, 1, last - 1, head.length).getValues();
        const iType = head.indexOf('類型'), iAte = head.indexOf('吃過的人'), iSt = head.indexOf('狀態');
        vals.forEach(function (v) {
          if (!v[1] || v[iType] === '住宿') return;
          t.places++;
          if ((iAte >= 0 && v[iAte]) || (iSt >= 0 && /吃過/.test(v[iSt]))) t.done++;
        });
      }
    } catch (e) { t.error = '打不開這份試算表'; }
    return t;
  });
  return { ok: true, trips: trips };
}
function fmtDate_(v) {
  if (!v) return '';
  if (Object.prototype.toString.call(v) === '[object Date]') return Utilities.formatDate(v, 'Asia/Taipei', 'yyyy-MM-dd');
  return String(v);
}

// ---------------- 建立新旅程 ----------------
function createTrip_(d) {
  const title = String(d.title || '').trim();
  if (!title) throw new Error('請填標題');
  const people = (d.people || []).map(String).map(function (s) { return s.trim(); }).filter(String);
  const country = String(d.country || '').trim();
  const city = String(d.city || '').trim();
  const region = String(d.region || '').trim();
  let lat = '', lng = '';
  try {
    const g = Maps.newGeocoder().setLanguage('zh-TW').setRegion(region || 'tw').geocode([city, country].filter(String).join(' '));
    if (g.status === 'OK' && g.results.length) { lat = g.results[0].geometry.location.lat; lng = g.results[0].geometry.location.lng; }
  } catch (e) {}

  const ss = SpreadsheetApp.create(title);
  const file = DriveApp.getFileById(ss.getId());
  file.moveTo(folder_());
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  const sh = ss.getSheets()[0];
  sh.setName('地點');
  sh.getRange(1, 1, 1, SHOP_HEADERS.length).setValues([SHOP_HEADERS]).setFontWeight('bold');
  sh.getRange(1, 1, sh.getMaxRows(), SHOP_HEADERS.length).setNumberFormat('@'); // 全部存文字，網頁讀取最穩定
  sh.setFrozenRows(1);

  const cfg = {
    '標題': title, '國家': country, '城市': city, '地區代碼': region, '時區': d.tz || 'Asia/Taipei',
    '開始日期': d.start || '', '結束日期': d.end || '', '天數': d.days ? String(d.days) : '',
    '同行的人': people.join('、'), '每日備註': '', '港鐵': /香港/.test(country) ? '是' : '否',
    '中心緯度': String(lat), '中心經度': String(lng)
  };
  ensureConfig_(ss, cfg);

  if (d.hotel) {
    try {
      const h = resolvePlace_(d.hotel, region, city);
      if (h.lat != null) {
        const row = {}; row['編號'] = 'hotel'; row['店名'] = h.name || '住宿'; row['類型'] = '住宿'; row['區域'] = h.area || '';
        row['地址'] = h.address || ''; row['緯度'] = h.lat; row['經度'] = h.lng; row['Google 地圖連結'] = h.url || d.hotel;
        sh.appendRow(SHOP_HEADERS.map(function (k) { return row[k] === undefined ? '' : String(row[k]); }));
      }
    } catch (e) {}
  }

  registry_().appendRow([ss.getId(), title, country, city, cfg['開始日期'], cfg['結束日期'], cfg['同行的人'], new Date()]);
  return { ok: true, id: ss.getId() };
}

// ---------------- 旅程設定 ----------------
function configSheet_(ss) { return ss.getSheetByName(CONFIG_SHEET); }
function ensureConfig_(ss, defaults) {
  let sh = configSheet_(ss);
  if (!sh) {
    sh = ss.insertSheet(CONFIG_SHEET);
    sh.getRange('A:B').setNumberFormat('@');
    const rows = [['項目', '內容']].concat(CONFIG_KEYS.map(function (k) { return [k, defaults[k] === undefined ? '' : String(defaults[k])]; }));
    sh.getRange(1, 1, rows.length, 2).setValues(rows);
    sh.getRange(1, 1, 1, 2).setFontWeight('bold');
    sh.setColumnWidth(1, 110); sh.setColumnWidth(2, 320);
  }
  return readConfig_(ss, defaults);
}
function readConfig_(ss, defaults) {
  const out = Object.assign({}, defaults || {});
  const sh = configSheet_(ss);
  if (!sh || sh.getLastRow() < 2) return out;
  sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues().forEach(function (r) { if (r[0]) out[String(r[0])] = fmtDate_(r[1]); });
  return out;
}
function setConfig_(trip, fields) {
  const ss = tripSS_(trip);
  ensureConfig_(ss, ss.getId() === LEGACY_ID ? LEGACY_CONFIG : {});
  const sh = configSheet_(ss);
  const keys = sh.getRange(1, 1, sh.getLastRow(), 1).getValues().map(function (r) { return String(r[0]); });
  Object.keys(fields || {}).forEach(function (k) {
    if (EDITABLE_CONFIG.indexOf(k) < 0) return;
    const v = String(fields[k] == null ? '' : fields[k]).slice(0, 200);
    const i = keys.indexOf(k);
    if (i >= 0) sh.getRange(i + 1, 2).setValue(v);
    else { sh.appendRow([k, v]); keys.push(k); }
  });
  // 總表的標題也一起更新
  if (fields && fields['標題']) {
    try {
      const reg = registry_();
      const ids = registeredIds_();
      const at = ids.indexOf(ss.getId());
      if (at >= 0) reg.getRange(at + 2, 2).setValue(String(fields['標題']));
      if (ss.getId() !== LEGACY_ID) ss.rename(String(fields['標題']));
    } catch (e) {}
  }
  return { ok: true };
}

// ---------------- 地點（店家） ----------------
function tripSS_(trip) {
  const id = trip || LEGACY_ID;
  if (id !== LEGACY_ID && registeredIds_().indexOf(id) < 0) throw new Error('找不到這個旅程');
  return SpreadsheetApp.openById(id);
}
function sheet_(trip) {
  const sh = tripSS_(trip).getSheets()[0];
  let headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);
  EXTRA_COLUMNS.forEach(function (c) {
    if (headers.indexOf(c) < 0) {
      sh.getRange(1, headers.length + 1).setValue(c);
      headers.push(c);
    }
  });
  return { sh: sh, headers: headers };
}

function addShop_(trip, d) {
  if (!d || !d['店名']) return { ok: false, error: '缺少店名' };
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const s = sheet_(trip);
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

function updateShop_(trip, d) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const s = sheet_(trip);
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

function deleteShop_(trip, d) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const s = sheet_(trip);
    const row = checkRow_(s, d);
    s.sh.deleteRow(row);
    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}

// ---------------- Google 地圖連結 ----------------
function resolvePlace_(input, region, city) {
  region = region || 'hk';
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

  let address = '', area = '';
  const geocoder = Maps.newGeocoder().setRegion(region).setLanguage('zh-TW');
  if (lat == null && (q || name)) {
    const g = geocoder.geocode(q || (name + ' ' + (city || '')));
    if (g.status === 'OK' && g.results.length) {
      lat = g.results[0].geometry.location.lat;
      lng = g.results[0].geometry.location.lng;
      address = g.results[0].formatted_address;
      area = areaOf_(g.results[0]);
    }
  }
  if (lat != null && (!address || !area)) {
    const g = geocoder.reverseGeocode(lat, lng);
    if (g.status === 'OK' && g.results.length) {
      if (!address) address = g.results[0].formatted_address;
      for (let i = 0; i < g.results.length && !area; i++) area = areaOf_(g.results[i]);
    }
  }
  // q 通常是「地址＋店名」，把店名拆出來
  if (!name && q && address) name = q.replace(address, '').trim();
  if (!name && q) name = q;
  return { ok: true, name: name, address: cleanAddress_(address), area: area, lat: lat, lng: lng, url: m ? m[0] : '', expanded: url };
}

function areaOf_(res) {
  const want = ['neighborhood', 'sublocality_level_1', 'sublocality', 'locality'];
  const comps = res.address_components || [];
  for (let w = 0; w < want.length; w++) {
    for (let i = 0; i < comps.length; i++) {
      if ((comps[i].types || []).indexOf(want[w]) >= 0) return comps[i].long_name;
    }
  }
  return '';
}

function cleanAddress_(a) {
  return String(a || '')
    .replace(/^(香港|日本|韓國|南韓|台灣|臺灣|泰國|新加坡|馬來西亞|越南|澳門|中國|菲律賓|印尼)\s*/, '')
    .replace(/^〒?\d{3}-?\d{4}\s*/, '')
    .replace(/^(香港島|九龍|新界)\s*/, '');
}
