/**
 * REPORT NCC - web app.
 * Data được dựng SẴN cho từng NCC (hàm refreshAll), mỗi NCC một file riêng trong thư mục Drive "Report NCC - data".
 * Khách mở link chỉ đọc đúng file của NCC mình -> trang mở nhanh, không thể lấy data NCC khác.
 * Tự làm mới lúc 0h hằng ngày (bật ở menu Report NCC), hoặc bấm "Làm mới data ngay" sau khi cập nhật Sheet.
 * Mỗi NCC có một mã link riêng, quản lý ở sheet "5. LINK NCC".
 */
var CACHE_FOLDER = 'Report NCC - data';
var TRACUU_URL = 'https://khovanhcm.github.io/report-ncc/tracuu/';
var SHEET_LINK = '5. LINK NCC';
var ALL_NAME = 'Tất cả NCC';
var DONE = '5 - Hoàn tất giao hàng';
var DONE2 = 'FBC - Đã Trả Hàng';  // rule 9.2: cũng tính là hoàn tất

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Report NCC')
    .addItem('Tạo / cập nhật link NCC', 'setupLinks')
    .addItem('Nhập data từ file NCC + làm mới link', 'dailyJob')
    .addItem('Làm mới data ngay', 'refreshAll')
    .addItem('Bật lịch tự động 8:00 hằng ngày', 'installDailyTrigger')
    .addToUi();
}

function doGet(e) {
  var token = String((e && e.parameter && e.parameter.k) || '').trim();
  // Lưu giao diện mặc định cho khách: chỉ nhận từ mã link "Tất cả NCC" (nội bộ)
  if (e && e.parameter && e.parameter.theme != null) {
    var res = saveTheme_(token, e.parameter.theme, String(e.parameter.scope || ''));
    return ContentService.createTextOutput(JSON.stringify(res)).setMimeType(ContentService.MimeType.JSON);
  }
  // Tra hành trình vận đơn trực tiếp từ logisticpartner.thegioididong.com (chỉ bill thuộc NCC của link)
  if (e && e.parameter && e.parameter.track != null) {
    var tr = trackBills_(token, String(e.parameter.track || ''));
    return ContentService.createTextOutput(JSON.stringify(tr)).setMimeType(ContentService.MimeType.JSON);
  }
  var json = withTheme_(readCache_(token));
  // Trang khách (GitHub Pages) gọi kèm api=1: chỉ trả data JSON của đúng mã link
  if (e && e.parameter && e.parameter.api) {
    return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
  }
  var t = HtmlService.createTemplateFromFile('Index');
  // chống chèn mã: không để chuỗi "</script" lọt vào trang
  t.data = json.replace(/</g, '\\u003c').replace(/\u2028/g, '\\u2028').replace(/\u2029/g, '\\u2029');
  return t.evaluate().setTitle('Report NCC')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/** Đọc file data đã dựng sẵn của đúng mã link. Mã sai -> báo lỗi, không trả data nào. */
function readCache_(token) {
  var props = PropertiesService.getScriptProperties();
  var id = token ? props.getProperty('f_' + token) : null;
  // Nhanh nhất: bộ nhớ đệm của Google (giữ 6 tiếng)
  if (id) {
    var fast = cacheGet_(token);
    if (fast) return fast;
  }
  var ncc = token ? tokenMap_()[token] : null;
  if (!ncc) return JSON.stringify({ error: 'Link không hợp lệ hoặc đã bị thu hồi. Vui lòng liên hệ Kho vận HCM.' });
  if (id) {
    try { var s = DriveApp.getFileById(id).getBlob().getDataAsString('UTF-8'); cachePut_(token, s); return s; } catch (err) {}
  }
  // Chưa có file (NCC mới / mã mới): dựng ngay cho NCC này rồi lưu lại
  var d = getData(token);
  saveCache_(token, d);
  return JSON.stringify(d);
}

function cacheFolder_() {
  var it = DriveApp.getFoldersByName(CACHE_FOLDER);
  return it.hasNext() ? it.next() : DriveApp.createFolder(CACHE_FOLDER);
}

/** Bộ nhớ đệm: chia nhỏ mỗi phần 90.000 ký tự (giới hạn 100KB/khóa). */
function cachePut_(token, s) {
  try {
    var c = CacheService.getScriptCache(), parts = {}, n = Math.ceil(s.length / 90000);
    for (var i = 0; i < n; i++) parts['d_' + token + '_' + i] = s.substr(i * 90000, 90000);
    parts['d_' + token + '_n'] = String(n);
    c.putAll(parts, 21600);
  } catch (err) {}
}

function cacheGet_(token) {
  try {
    var c = CacheService.getScriptCache(), n = +c.get('d_' + token + '_n');
    if (!n) return null;
    var keys = [];
    for (var i = 0; i < n; i++) keys.push('d_' + token + '_' + i);
    var got = c.getAll(keys), s = '';
    for (var j = 0; j < n; j++) { if (got[keys[j]] == null) return null; s += got[keys[j]]; }
    return s;
  } catch (err) { return null; }
}

function saveCache_(token, d) {
  var props = PropertiesService.getScriptProperties(), key = 'f_' + token, id = props.getProperty(key);
  var s = JSON.stringify(d);
  cachePut_(token, s);
  if (id) {
    try { DriveApp.getFileById(id).setContent(s); return; } catch (err) {}
  }
  var f = cacheFolder_().createFile(token + '.json', s, 'application/json');
  props.setProperty(key, f.getId());
}

/** Đọc Sheet 1 lần, dựng data riêng cho từng NCC và lưu mỗi NCC một file. */
function refreshAll() {
  var src = readSource_(), map = tokenMap_(), n = 0;
  Object.keys(map).forEach(function (token) {
    var ncc = map[token];
    saveCache_(token, buildData_(src.rows2, src.s3map, src.khoName, ncc === ALL_NAME ? null : ncc, src.snap, src.noRet, src.khoProv));
    n++;
  });
  try { SpreadsheetApp.getActive().toast('Đã làm mới data cho ' + n + ' link lúc ' + src.snap, 'Report NCC', 8); } catch (err) {}
}

/** Tạo lịch tự chạy dailyJob lúc 8:00 mỗi ngày (giờ Việt Nam) + nạp bộ nhớ đệm 4 tiếng/lần. Chạy 1 lần là đủ. */
function installDailyTrigger() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    var f = t.getHandlerFunction();
    if (f === 'refreshAll' || f === 'dailyJob') ScriptApp.deleteTrigger(t);
  });
  // 8:00 hằng ngày: nhập data từ file NCC rồi làm mới link (Google chạy trong khung 8:00-8:15)
  ScriptApp.newTrigger('dailyJob').timeBased().atHour(8).nearMinute(0).everyDays(1).inTimezone('Asia/Ho_Chi_Minh').create();
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'warmCache') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('warmCache').timeBased().everyHours(4).create();
  refreshAll();
}

/** Nạp lại bộ nhớ đệm từ file đã dựng (không đọc lại Sheet) để link luôn mở nhanh. */
function warmCache() {
  var props = PropertiesService.getScriptProperties();
  Object.keys(tokenMap_()).forEach(function (token) {
    var id = props.getProperty('f_' + token);
    if (!id) return;
    try { cachePut_(token, DriveApp.getFileById(id).getBlob().getDataAsString('UTF-8')); } catch (err) {}
  });
}

function sheetByPrefix_(prefix) {
  var sheets = SpreadsheetApp.getActive().getSheets();
  for (var i = 0; i < sheets.length; i++) {
    if (sheets[i].getName().trim().indexOf(prefix) === 0) return sheets[i];
  }
  throw new Error('Không tìm thấy sheet bắt đầu bằng "' + prefix + '"');
}

function col_(sheet, col) {
  var n = sheet.getLastRow() - 1;
  if (n < 1) return [];
  return sheet.getRange(2, col, n, 1).getDisplayValues().map(function (r) { return r[0]; });
}

/** Tạo sheet "5. LINK NCC": mỗi NCC một mã link. Mã đã có được giữ nguyên, NCC mới được thêm mã mới. */
function setupLinks() {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(SHEET_LINK) || ss.insertSheet(SHEET_LINK);
  var old = {};
  if (sh.getLastRow() >= 4) {
    sh.getRange(4, 1, sh.getLastRow() - 3, 2).getDisplayValues().forEach(function (r) {
      if (r[0] && r[1]) old[nkey_(r[0])] = r[1].trim();
    });
  }
  var url = String(sh.getRange('B1').getDisplayValue() || '').trim();
  var page = String(sh.getRange('B2').getDisplayValue() || '').trim();
  var names = [ALL_NAME], seen = {};
  col_(sheetByPrefix_('2.'), 1).forEach(function (v) {
    v = v.trim();
    if (v && !seen[v]) { seen[v] = 1; names.push(v); }
  });
  sh.clear();
  sh.getRange('A1:B2').setValues([['URL web app (dán link kết thúc bằng /exec vào ô B1)', url],
    ['URL trang khách GitHub Pages (để trống nếu chưa có)', page]]);
  sh.getRange('A3:D3').setValues([['Tên NCC', 'Mã link (đổi mã = thu hồi link cũ)', 'Link gửi khách', 'Link tra cứu đơn hàng']]).setFontWeight('bold');
  var rows = names.map(function (n) {
    return [n, old[nkey_(n)] || Utilities.getUuid().replace(/-/g, '').slice(0, 16)];
  });
  sh.getRange(4, 1, rows.length, 2).setValues(rows);
  for (var i = 0; i < rows.length; i++) {
    sh.getRange(4 + i, 3).setFormula('=IF($B$2<>"",$B$2,$B$1)&"?k="&B' + (4 + i));
    sh.getRange(4 + i, 4).setFormula('="' + TRACUU_URL + '?k="&B' + (4 + i));  // trang tra cứu đơn hàng, tự sinh theo mã
  }
  sh.autoResizeColumns(1, 4);
}

function tokenMap_() {
  var sh = SpreadsheetApp.getActive().getSheetByName(SHEET_LINK);
  var map = {};
  if (!sh || sh.getLastRow() < 4) return map;
  sh.getRange(4, 1, sh.getLastRow() - 3, 2).getDisplayValues().forEach(function (r) {
    if (r[0] && r[1]) map[r[1].trim()] = r[0].trim();
  });
  return map;
}

/** Dựng data trực tiếp từ Sheet cho 1 mã link (dùng khi chưa có file dựng sẵn). */
function getData(token) {
  var ncc = tokenMap_()[String(token || '').trim()];
  if (!ncc) return { error: 'Link không hợp lệ hoặc đã bị thu hồi. Vui lòng liên hệ Kho vận HCM.' };
  var src = readSource_();
  return buildData_(src.rows2, src.s3map, src.khoName, ncc === ALL_NAME ? null : ncc, src.snap, src.noRet, src.khoProv);
}

function readSource_() {
  var s2 = sheetByPrefix_('2.');
  var n2 = s2.getLastRow() - 1;
  var rows2 = n2 > 0 ? s2.getRange(2, 1, n2, 13).getDisplayValues() : [];  // A..M (M = DO, rule A.12)

  // Sheet 3: cột E mã chứng từ đối tác, cột J nơi nhận, cột AH trạng thái giao hàng
  var s3 = sheetByPrefix_('3.');
  var e = col_(s3, 5), noi = col_(s3, 10), tn = col_(s3, 11), st = col_(s3, 34), dt = col_(s3, 35);  // K = tỉnh nhận
  var s3map = {};
  for (var i = 0; i < e.length; i++) { var k = e[i].trim(); if (k) s3map[k] = [st[i].trim(), noi[i].trim(), dt[i], String(tn[i] || '').trim()]; }

  // Sheet 4: cột B mã kho, cột C tên kho
  var s4 = sheetByPrefix_('4.');
  var code = col_(s4, 2), name = col_(s4, 3), tinh = col_(s4, 5);
  var khoName = {}, khoProv = {};
  for (var j = 0; j < code.length; j++) {
    var c = String(parseInt(code[j], 10));
    if (c !== 'NaN' && !(c in khoName)) { khoName[c] = name[j].trim(); khoProv[c] = String(tinh[j] || '').trim(); }
  }

  var snap = Utilities.formatDate(new Date(), 'Asia/Ho_Chi_Minh', "HH:mm 'ngày' dd/MM/yyyy");
  return { rows2: rows2, s3map: s3map, khoName: khoName, khoProv: khoProv, snap: snap, noRet: noRet_() };
}

/** Rule B.10 ở Sheet 1: danh sách NCC không có rule trả chứng từ (các dòng giữa "10.1" và "10.2"). */
function noRet_() {
  try {
    var vals = sheetByPrefix_('1.').getDataRange().getDisplayValues();
    for (var i = 0; i < vals.length; i++) for (var j = 0; j < vals[i].length; j++) {
      var t = vals[i][j];
      if (t.indexOf('10.1') >= 0 && t.indexOf('10.2') >= 0) {
        return t.split('10.2')[0].split('10.1')[1].split('\n').slice(1)
          .map(function (s) { return s.replace(/^[\s\-–•]+/, '').trim(); }).filter(String);
      }
    }
  } catch (err) {}
  return [];
}

/** So tên NCC không phân biệt hoa thường và dấu (Saring = SARING, HOÀ PHÁT = HÒA PHÁT). */
function nkey_(s) {
  return String(s || '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/đ/g, 'd').replace(/Đ/g, 'D').toUpperCase().replace(/\s+/g, ' ').trim();
}

function nd_(s) {
  var m = /^\s*(\d{1,2})\/(\d{1,2})\/(\d{4})/.exec(s || '');
  return m ? ('0' + m[1]).slice(-2) + '/' + ('0' + m[2]).slice(-2) + '/' + m[3] : '';
}

function prov_(a) {
  var m = /(Tỉnh|Thành phố)\s+([^,]+?)\s*(,\s*Vi[eêệ]t Nam)?\s*$/i.exec(a || '');
  return m ? ((m[1].toLowerCase() === 'tỉnh' ? 'Tỉnh ' : 'Thành phố ') + m[2].trim()) : '';
}

/**
 * Gom data cho trang web.
 * Dòng: [booking, lấy hàng, PO, hạn PO, bill, địa chỉ, cột I, KTT (cột L hoặc K), tỉnh, NCC, kho nhận]
 * BILLS: bill -> [trạng thái Sheet 3, ngày hoàn tất (cột J Sheet 2), data trả chứng từ (cột I)]
 */
function buildData_(rows2, s3map, khoName, onlyNcc, snap, noRet, khoProv) {
  khoProv = khoProv || {};
  var dict = [], ix = {}, nccs = [], R = [], B = {}, only = onlyNcc ? nkey_(onlyNcc) : '';
  function D(v) { if (!(v in ix)) { ix[v] = dict.length; dict.push(v); } return ix[v]; }
  rows2.forEach(function (r) {
    var nm = String(r[0] || '').trim();
    if (!nm || (only && nkey_(nm) !== only)) return;
    var b = nd_(r[2]); if (!b) return;
    var p = nd_(r[3]), po = String(r[4] || '').trim(), han = nd_(r[5]) || String(r[5] || '').trim();
    var bill = String(r[6] || '').trim(), addr = String(r[7] || '').trim(), I = String(r[8] || '').trim();
    // KTT phát cuối: ưu tiên cột L, không có thì lấy cột K; bỏ mã kho ở đầu cho đồng nhất
    var ktt = String(r[11] || r[10] || '').trim().replace(/^\d+\s*-\s*/, ''), pv = prov_(addr);
    // Địa chỉ là tên kho có mã ở đầu (vd KAROFI "13359 TNB_CTH - Kho TT…"): lấy tỉnh theo mã kho ở Sheet 4 cột E
    if (!pv) { var mc = /^(\d{3,6})\b/.exec(addr) || /^(\d{5})PO/.exec(po); if (mc) pv = khoProv[String(parseInt(mc[1], 10))] || ''; }
    var x = bill ? s3map[bill] : null;
    var kn = '';
    if (/^\d{5}PO\d{10}$/.test(po)) kn = khoName[String(parseInt(po.slice(0, 5), 10))] || '';
    if (!kn && x && x[1].indexOf(' - ') > 0) kn = x[1].split(' - ').slice(1).join(' - ').trim();
    if (bill) {
      // Ngày hoàn tất giao: cột J Sheet 2; cột J trống thì lấy cột AI Sheet 3 khi trạng thái là hoàn tất (rule A.9)
      if (!B[bill]) B[bill] = [x ? x[0] : '', nd_(r[9]) || (x && (x[0] === DONE || x[0] === DONE2) ? nd_(x[2]) : ''), '', x && x[3] ? x[3] : ''];  // [3] = tỉnh nhận (Sheet 3 cột K)
      if (I && !B[bill][2]) B[bill][2] = I;
    }
    if (nccs.indexOf(nm) < 0) nccs.push(nm);
    R.push([b, p, po, han, bill, D(addr), I, ktt ? D(ktt) : -1, pv ? D(pv) : -1, nccs.indexOf(nm), kn ? D(kn) : -1, String(r[12] || '').trim()]);
  });
  var nrk = (noRet || []).map(nkey_);
  var nr = nccs.filter(function (n) { return nrk.indexOf(nkey_(n)) >= 0; });
  return { NCC: onlyNcc ? (nccs[0] || onlyNcc) : ALL_NAME, NCCS: nccs, NORET: nr, SNAP: snap, DICT: dict, BILLS: B, ROWS: R };
}

/* =====================================================================
 * NHẬP DATA TỪ FILE NGUỒN NCC VÀO SHEET "2.DATA REPORT"
 * Khai báo ở sheet "6.NGUỒN NCC", mỗi NCC một khối:
 *   I. <TÊN NCC>
 *   1 LINK: | <link file nguồn>
 *   2 Sheet : | "<tên tab>"
 *   4.x <tên thông tin> | CỘT <chữ cột> | (ghi chú, giá trị mặc định đặt trong "...")
 *   (tùy chọn) Lấy data từ ngày | dd/mm/yyyy   -> mặc định 01/10/2026
 * Mỗi lần chạy: thay toàn bộ dòng của NCC đó trong Sheet 2 bằng data mới; cột J (công thức) giữ nguyên.
 * NCC nào lỗi thì giữ nguyên data cũ và ghi lỗi cạnh tên NCC ở sheet 6.
 * ===================================================================== */
var SRC_SHEET = '6.';
var FROM_DEFAULT = '01/10/2026';

/** Chạy mỗi sáng 8:00: nhập data từ file NCC rồi làm mới các link báo cáo. */
function dailyJob() {
  importAll();
  // Làm mới link ở một lượt chạy riêng sau 1 phút, tránh vượt giới hạn 6 phút/lượt của Google
  ScriptApp.newTrigger('refreshAfterImport').timeBased().after(60 * 1000).create();
  try { SpreadsheetApp.getActive().toast('Đã nhập xong data. Link báo cáo sẽ tự làm mới sau khoảng 1–3 phút.', 'Report NCC', 10); } catch (err) {}
}

/** Chạy 1 lần sau dailyJob: xóa lịch tạm của chính nó rồi làm mới link. */
function refreshAfterImport() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'refreshAfterImport') ScriptApp.deleteTrigger(t);
  });
  refreshAll();
}

function importAll() {
  var cfgs = readSourceConfig_();
  var s2 = sheetByPrefix_('2.');
  var res = [];
  cfgs.forEach(function (cfg) {
    var msg;
    try {
      var src = SpreadsheetApp.openByUrl(cfg.url);
      var sh = cfg.tab ? src.getSheetByName(cfg.tab) : src.getSheets()[0];
      if (!sh && cfg.tab) sh = src.getSheets().filter(function (s) { return nkey_(s.getName()) === nkey_(cfg.tab); })[0];
      if (!sh) throw new Error('Không thấy tab "' + cfg.tab + '". Các tab có trong file: ' + src.getSheets().map(function (s) { return s.getName(); }).join(', '));
      var vals = sh.getDataRange().getDisplayValues().slice(1);
      var out = mapRows_(vals, cfg);
      if (!out.rows.length) throw new Error('Không có dòng nào hợp lệ từ ngày ' + cfg.from);
      replaceNccRows_(s2, cfg.ncc, out.rows);
      msg = 'OK ' + Utilities.formatDate(new Date(), 'Asia/Ho_Chi_Minh', 'HH:mm dd/MM/yyyy') + ': ' + out.rows.length + ' dòng'
        + (out.skipped ? ' · bỏ ' + out.skipped + ' dòng sai ngày booking' : '');
    } catch (err) {
      msg = 'LỖI ' + Utilities.formatDate(new Date(), 'Asia/Ho_Chi_Minh', 'HH:mm dd/MM/yyyy') + ': ' + err.message + ' (giữ data cũ)';
    }
    cfg.statusCell.setValue(msg);
    res.push(cfg.ncc + ': ' + msg);
  });
  try { var nk = fillKtt_(); res.push('Cột K (rule 11): ' + nk.ok + ' dòng có kho, ' + nk.miss + ' dòng chưa xác định được tỉnh'); }
  catch (err) { res.push('Cột K (rule 11) lỗi: ' + err.message); }
  try { SpreadsheetApp.getActive().toast(res.join('\n'), 'Nhập data NCC', 10); } catch (err) {}
  return res;
}

/** Đọc các khối khai báo ở sheet 6. */
function readSourceConfig_() {
  var sh = sheetByPrefix_(SRC_SHEET);
  var v = sh.getDataRange().getDisplayValues(), cfgs = [], cur = null;
  for (var i = 0; i < v.length; i++) {
    var a = v[i][0].trim(), b = (v[i][1] || '').trim(), c = (v[i][2] || '').trim(), d = (v[i][3] || '').trim();
    var head = /^[IVXLC]+\.\s*(.+)$/.exec(a);
    if (head) {
      cur = { ncc: head[1].trim(), url: '', tab: '', from: FROM_DEFAULT, map: {}, def: {}, statusCell: sh.getRange(i + 1, 3) };
      cfgs.push(cur); continue;
    }
    if (!cur) continue;
    var ka = nkey_(a), kb = nkey_(b);
    if (ka.indexOf('LINK') >= 0) { cur.url = b; continue; }
    if (/^\d+\s*SHEET/.test(ka)) { cur.tab = b.replace(/^["“”\s]+|["“”\s]+$/g, ''); continue; }
    if (ka.indexOf('TU NGAY') >= 0 || kb.indexOf('TU NGAY') >= 0) { var f = normDate_(ka.indexOf('TU NGAY') >= 0 ? b : c); if (f) cur.from = f; continue; }
    if (!/^4\.\d+/.test(b)) continue;
    var field = fieldOf_(kb);
    if (!field) continue;
    if (field === 'A') { if (c) cur.ncc = c; continue; }
    var m = /COT\s+([A-Z]{1,2})\b/.exec(nkey_(c));
    if (m) cur.map[field] = colIndex_(m[1]);
    var yr = /\/\s*(20\d\d)\b/.exec(d);
    if (yr && (field === 'C' || field === 'D')) { cur.year = +yr[1]; continue; }
    var q = /["“]([^"”]+)["”]/.exec(d);
    if (q) cur.def[field] = q[1].trim();
  }
  return cfgs.filter(function (x) { return x.url; });
}

/** Nhận biết thông tin theo tên dòng khai báo -> cột ở Sheet 2. */
function fieldOf_(k) {
  if (/(^|\s)DO$/.test(k)) return 'M';  // rule A.12: số DO (cột M Sheet 2)
  if (k.indexOf('TEN NHA CUNG CAP') >= 0) return 'A';
  if (k.indexOf('THANG') >= 0) return 'B';
  if (k.indexOf('LAY HANG') >= 0) return 'D';
  if (k.indexOf('BOOKING') >= 0) return 'C';
  if (k.indexOf('HAN PO') >= 0) return 'F';
  if (k.indexOf('MA PO') >= 0) return 'E';
  if (k.indexOf('TRA') >= 0 && (k.indexOf('BILL') >= 0 || k.indexOf('CHUNG TU') >= 0)) return 'I';
  if (k.indexOf('BILL') >= 0) return 'G';
  if (k.indexOf('DIA CHI') >= 0) return 'H';
  if (k.indexOf('KHO TRUNG TAM') >= 0 || k.indexOf('KTT') >= 0) return 'L';
  return '';
}

function colIndex_(s) {
  var n = 0;
  for (var i = 0; i < s.length; i++) n = n * 26 + (s.charCodeAt(i) - 64);
  return n - 1;
}

/** Chuẩn hóa ngày về dd/mm/yyyy; năm phải từ 2020 đến 2099, sai thì trả ''. */
function normDate_(s) {
  var m = /^\s*(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{4})/.exec(s || '');
  if (!m) return '';
  var d = +m[1], mo = +m[2], y = +m[3];
  if (mo > 12 && d <= 12) { var t = d; d = mo; mo = t; }  // dạng tháng/ngày/năm (vd 10/21/2026)
  if (y < 2020 || y > 2099 || mo < 1 || mo > 12 || d < 1 || d > 31) return '';
  return ('0' + d).slice(-2) + '/' + ('0' + mo).slice(-2) + '/' + y;
}

function dkey_(s) { return s ? s.slice(6) + s.slice(3, 5) + s.slice(0, 2) : ''; }

/** Chuyển dòng file nguồn thành dòng Sheet 2: [A..I, K, L] (J là công thức nên bỏ qua). */
function mapRows_(vals, cfg) {
  var rows = [], skipped = 0, fromK = dkey_(cfg.from);
  function g(r, f) {
    var i = cfg.map[f];
    var x = i == null ? '' : String(r[i] == null ? '' : r[i]).trim();
    if (/^#(N\/A|REF!|VALUE!|DIV\/0!|NAME\?|NUM!|NULL!|ERROR!)/.test(x)) x = '';  // ô lỗi công thức coi như trống
    return x || cfg.def[f] || '';
  }
  var year = cfg.year || +cfg.from.slice(6);
  vals.forEach(function (r) {
    if (!r.some(function (x) { return String(x).trim(); })) return;
    // Cột tháng: là số 1-12 thì dùng; nếu là ngày (NCC chỉ có cột ngày booking) thì để lấy từ ngày booking
    var mRaw = g(r, 'B'), mNum = /^\d{1,2}$/.test(mRaw) && +mRaw >= 1 && +mRaw <= 12 ? +mRaw : 0;
    var bRaw = g(r, 'C'), book = normDate_(bRaw);
    // NCC chỉ ghi số ngày (vd 16) + cột tháng riêng -> ghép thành ngày/tháng/năm
    if (!book && /^\d{1,2}$/.test(bRaw) && mNum) book = normDate_(bRaw + '/' + mNum + '/' + year);
    if (!book) { skipped++; return; }
    // Ngày ghi kiểu tháng/ngày (vd tháng 5 mà ghi 05/11/2026 = 11/05) -> đối chiếu cột tháng để đảo lại
    if (mNum && +book.slice(3, 5) !== mNum && +book.slice(0, 2) === mNum) book = book.slice(3, 5) + '/' + book.slice(0, 2) + '/' + book.slice(6);
    if (dkey_(book) < fromK) return;
    var month = String(mNum || +book.slice(3, 5));
    var pRaw = g(r, 'D'), pick = normDate_(pRaw);
    if (!pick && /^\d{1,2}$/.test(pRaw)) {
      var bm = +book.slice(3, 5), by = +book.slice(6);
      if (+pRaw < +book.slice(0, 2)) { bm++; if (bm > 12) { bm = 1; by++; } }  // ngày lấy hàng nhỏ hơn ngày booking -> sang tháng sau
      pick = normDate_(pRaw + '/' + bm + '/' + by);
    }
    pick = pick || pRaw;
    var han = normDate_(g(r, 'F')) || g(r, 'F');
    var bill = g(r, 'G'), po = g(r, 'E') || bill;  // không có PO thì lấy mã bill làm mã PO (như cách nhập tay trước đây)
    rows.push([cfg.ncc, month, book, pick, po, han, bill, g(r, 'H'), g(r, 'I'), '', g(r, 'L'), g(r, 'M')]);
  });
  return { rows: rows, skipped: skipped };
}

/** Thay các dòng của 1 NCC trong Sheet 2 (giữ nguyên dòng NCC khác và cột J). */
function replaceNccRows_(s2, ncc, newRows) {
  var last = s2.getLastRow(), key = nkey_(ncc), keep = [];
  if (last >= 2) {
    var ai = s2.getRange(2, 1, last - 1, 9).getDisplayValues();
    var kl = s2.getRange(2, 11, last - 1, 3).getDisplayValues();  // K, L, M
    for (var i = 0; i < ai.length; i++) {
      if (!ai[i][0].trim() || nkey_(ai[i][0]) === key) continue;
      keep.push(ai[i].concat(kl[i]));
    }
  }
  var all = keep.concat(newRows);
  if (last >= 2) { s2.getRange(2, 1, last - 1, 9).clearContent(); s2.getRange(2, 11, last - 1, 3).clearContent(); }
  if (!all.length) return;
  var left = all.map(function (r) { return r.slice(0, 9); });
  var right = all.map(function (r) { var x = r.slice(9, 12); while (x.length < 3) x.push(''); return x; });
  s2.getRange(2, 1, all.length, 9).setNumberFormat('@').setValues(left);
  s2.getRange(2, 11, all.length, 3).setNumberFormat('@').setValues(right);
}

/**
 * Rule A.11: lấy Tỉnh/Thành phố trong địa chỉ (cột H Sheet 2), so với cột E Sheet 4,
 * trả về cột L Sheet 4 (KTT phụ trách) vào cột K Sheet 2.
 */
function fillKtt_() {
  var v4 = sheetByPrefix_('4.').getDataRange().getDisplayValues().slice(1);
  var map = {}, names = [], byKho = {}, byCode = {};
  function pk(s) { return nkey_(s).replace(/^(TINH|THANH PHO|TP\.?)\s+/, '').trim(); }
  function kk(s) { return nkey_(String(s || '').replace(/^\d+\s*-?\s*/, '')); }
  v4.forEach(function (r) {
    var k = pk(r[4] || ''), l = String(r[11] || '').trim();
    if (!l) return;
    if (k && !(k in map)) { map[k] = l; names.push(k); }
    var c = String(parseInt(r[1], 10)), n = kk(r[2]);
    if (c !== 'NaN' && !(c in byCode)) byCode[c] = l;
    if (n && !(n in byKho)) byKho[n] = l;
    var t = kk(l); if (t && !(t in byKho)) byKho[t] = l;          // tên KTT cũng là 1 kho
    var tc = String(parseInt(l, 10)); if (tc !== 'NaN' && !(tc in byCode)) byCode[tc] = l;
  });
  names.sort(function (a, b) { return b.length - a.length; });
  var s2 = sheetByPrefix_('2.'), n = s2.getLastRow() - 1;
  if (n < 1) return { ok: 0, miss: 0 };
  var addr = s2.getRange(2, 8, n, 1).getDisplayValues(), out = [], ok = 0, miss = 0;
  for (var i = 0; i < n; i++) {
    var a = String(addr[i][0] || '').trim(), val = '';
    if (a) {
      // 1) Địa chỉ có "Tỉnh/Thành phố ..." -> tra cột E Sheet 4
      var p = prov_(a), k = p ? pk(p) : '';
      val = (k && map[k]) || '';
      // 2) Cột H là tên kho / mã kho (NCC không ghi địa chỉ) -> tra tên kho cột C, mã kho cột B Sheet 4
      if (!val) val = byKho[kk(a)] || '';
      if (!val) { var m = /^(\d{3,6})\b/.exec(a); if (m) val = byCode[String(+m[1])] || ''; }
      // 3) Không có chữ "Tỉnh/Thành phố": tìm tên tỉnh trong 2 phần cuối địa chỉ
      if (!val) {
        var tail = nkey_(a.split(',').slice(-2).join(','));
        k = names.filter(function (x) { return tail.indexOf(x) >= 0; })[0] || '';
        val = map[k] || '';
      }
      if (val) ok++; else miss++;
    }
    out.push([val]);
  }
  s2.getRange(2, 11, n, 1).setNumberFormat('@').setValues(out);
  return { ok: ok, miss: miss };
}

/* ---------- Giao diện mặc định cho khách (màu, font) ---------- */
function saveTheme_(token, raw, scope) {
  if (tokenMap_()[token] !== ALL_NAME) return { ok: false, error: 'Chỉ link Tất cả NCC mới được đổi giao diện cho khách.' };
  var t;
  try { t = JSON.parse(raw || '{}'); } catch (err) { return { ok: false, error: 'Dữ liệu giao diện không hợp lệ.' }; }
  var out = {}, hex = /^#[0-9a-fA-F]{6}$/;
  ['bg', 'ac', 'tile', 'card', 'c1', 'c2', 'band', 'ok', 'bad', 'wait'].forEach(function (k) { if (hex.test(t[k] || '')) out[k] = t[k]; });
  if (+t.size >= 12 && +t.size <= 20) out.size = +t.size;
  if (typeof t.font === 'string' && t.font.length < 120 && !/[<>{};]/.test(t.font)) out.font = t.font;
  // scope=karofi: giao diện riêng của trang báo cáo KAROFI, không ảnh hưởng link các NCC khác
  var key = scope === 'karofi' ? 'THEME_KAROFI' : 'THEME';
  if (t.reset) PropertiesService.getScriptProperties().deleteProperty(key);
  else PropertiesService.getScriptProperties().setProperty(key, JSON.stringify(out));
  return { ok: true, theme: t.reset ? {} : out };
}

/** Gắn giao diện mặc định vào data trả cho trang (không cần dựng lại data). */
function withTheme_(json) {
  if (!json || json.charAt(json.length - 1) !== '}' || json.indexOf('{"error"') === 0) return json;
  var P = PropertiesService.getScriptProperties(), t = P.getProperty('THEME'), tk = P.getProperty('THEME_KAROFI'), add = '';
  if (t) add += ',"THEME":' + t;
  if (tk) add += ',"THEME_K":' + tk;
  return add ? json.slice(0, -1) + add + '}' : json;
}

/* ---------- Tra cứu vận đơn trực tiếp (logisticpartner.thegioididong.com) ---------- */
var TRACK_HOST = 'https://logisticpartner.thegioididong.com';

/** Nhận danh sách mã bill (cách nhau dấu phẩy, tối đa 5), chỉ tra bill có trong data của link này. */
function trackBills_(token, list) {
  var json = readCache_(token), d;
  try { d = JSON.parse(json); } catch (err) { return { ok: false, error: 'Không đọc được dữ liệu báo cáo.' }; }
  if (d.error) return { ok: false, error: d.error };
  var own = {};
  Object.keys(d.BILLS || {}).forEach(function (b) { own[b.toUpperCase()] = 1; });
  (d.ROWS || []).forEach(function (r) { if (r[4]) own[String(r[4]).toUpperCase()] = 1; });
  var codes = list.split(',').map(function (s) { return s.trim().toUpperCase(); })
    .filter(function (s) { return /^[A-Z0-9]{6,30}$/.test(s); }).slice(0, 5);
  var out = {}, cache = CacheService.getScriptCache();
  codes.forEach(function (b) {
    if (!own[b]) { out[b] = { ok: false, error: 'Mã bill không thuộc dữ liệu của link này.' }; return; }
    var c = cache.get('trk_' + b);
    if (c) { out[b] = JSON.parse(c); return; }
    var r = fetchTrack_(b);
    out[b] = r;
    if (r.ok) { try { cache.put('trk_' + b, JSON.stringify(r), 600); } catch (err) {} }  // giữ 10 phút
  });
  return { ok: true, bills: out };
}

/** Lấy mã chống giả mạo (token + cookie) của trang tra cứu, giữ 20 phút. */
function trackSession_(force) {
  var cache = CacheService.getScriptCache();
  if (!force) { var s = cache.get('trk_session'); if (s) return JSON.parse(s); }
  var res = UrlFetchApp.fetch(TRACK_HOST + '/TraCuu?Bill', { muteHttpExceptions: true, followRedirects: true });
  var html = res.getContentText();
  var m = /name="RequestVerificationToken"[^>]*value="([^"]+)"/.exec(html) || /value="([^"]+)"[^>]*name="RequestVerificationToken"/.exec(html);
  if (!m) throw new Error('Không lấy được mã phiên của trang tra cứu.');
  var hd = res.getAllHeaders(), sc = hd['Set-Cookie'] || hd['set-cookie'] || [];
  if (!Array.isArray(sc)) sc = [sc];
  var cookie = sc.map(function (x) { return String(x).split(';')[0]; }).join('; ');
  var sess = { token: m[1], cookie: cookie };
  cache.put('trk_session', JSON.stringify(sess), 1200);
  return sess;
}

function fetchTrack_(bill) {
  for (var attempt = 0; attempt < 2; attempt++) {
    try {
      var s = trackSession_(attempt > 0);
      var res = UrlFetchApp.fetch(TRACK_HOST + '/Home/GetTrackingBill', {
        method: 'post', contentType: 'application/json; charset=utf-8', muteHttpExceptions: true,
        headers: { 'X-Requested-With': 'XMLHttpRequest', 'RequestVerificationToken': encodeURIComponent(s.token), 'Cookie': s.cookie },
        payload: JSON.stringify({ transportvoucherid: bill })
      });
      if (res.getResponseCode() === 400 && attempt === 0) continue;  // phiên hết hạn -> lấy phiên mới rồi thử lại
      if (res.getResponseCode() !== 200) return { ok: false, error: 'Trang tra cứu báo lỗi ' + res.getResponseCode() + '.' };
      var j = JSON.parse(res.getContentText());
      if (j.isError) return { ok: false, error: j.message || 'Trang tra cứu báo lỗi.' };
      var rows = j.data || [];
      if (!rows.length) return { ok: true, at: nowTxt_(), events: [], note: '' };
      // Chỉ giữ thông tin cần cho khách; bỏ tên, số điện thoại người thao tác
      return {
        ok: true, at: nowTxt_(),
        note: rows[0].transportvouchernote || '',
        from: rows[0].sender_address || '', to: rows[0].receiver_address || '',
        events: rows.map(function (x) {
          return { t: x.tracking_time || '', s: x.statusname || '', id: x.statusid,
            img: x.image_paths ? String(x.image_paths).split(';').filter(function (u) { return /^https:\/\/[a-z0-9.-]*tgdd\.vn\//i.test(u); }) : [] };
        })
      };
    } catch (err) {
      if (attempt === 1) return { ok: false, error: 'Không kết nối được trang tra cứu: ' + err.message };
    }
  }
  return { ok: false, error: 'Không tra được.' };
}

function nowTxt_() { return Utilities.formatDate(new Date(), 'Asia/Ho_Chi_Minh', 'HH:mm dd/MM/yyyy'); }

/** Chạy thử trong trình soạn Apps Script: chọn hàm testTrack -> Chạy, xem Nhật ký thực thi. */
function testTrack() {
  Logger.log(JSON.stringify(fetchTrack_('01907TD26110002106')).slice(0, 1500));
}
