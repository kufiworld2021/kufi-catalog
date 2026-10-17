/**
 * عالم الكوفي — مخزن بلاغات العملاء (Google Apps Script)
 * يستقبل البلاغ من صفحة GitHub، يحفظ الصور في Drive، ويكتب صف في الشيت.
 * ويرد على صفحة التتبع بحالة البلاغ.
 *
 * أول مرة: شغّل الدالة setup مرة وحدة من الأعلى (Run) ووافق على الصلاحيات.
 */

var SHEET = 'البلاغات';
var FOLDER = 'بلاغات العملاء — صور';
var NOTIFY = true;          // إيميل لك مع كل بلاغ جديد
var SHARE_PHOTOS = true;    // مجلد الصور يفتح لمن عنده الرابط (عشان الموظف يشوفه من التطبيق)

// الحالات — غيّر الحالة من القائمة في الشيت والعميل يشوفها فوراً في صفحة التتبع
var STATUS = [
  ['جديد', 'new'],
  ['قيد المراجعة', 'review'],
  ['بانتظار معلومات من العميل', 'info'],
  ['تم القرار', 'decided'],
  ['تم الشحن / التحويل', 'shipped'],
  ['مغلق', 'closed'],
  ['مرفوض', 'rejected']
];

var COLS = ['رقم البلاغ', 'وقت البلاغ', 'رقم الطلب', 'الاسم', 'الجوال', 'نوع الضرر', 'تاريخ الاستلام',
  'الوصف', 'عدد الصور', 'مجلد الصور', 'الحالة', 'رسالة للعميل (تظهر له)', 'آخر تحديث', 'تنبيه', 'نص التطبيق (انسخه)'];
var C_FOLDER = 10, C_ID = 1, C_AT = 2, C_ORDER = 3, C_NAME = 4, C_PHONE = 5, C_ST = 11, C_NOTE = 12, C_UPD = 13, C_FLAG = 14, C_APP = 15;

function setup() {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(SHEET) || ss.insertSheet(SHEET, 0);
  sh.getRange(1, 1, 1, COLS.length).setValues([COLS]).setFontWeight('bold')
    .setBackground('#53230E').setFontColor('#F6E7D2').setHorizontalAlignment('center');
  sh.setFrozenRows(1);
  sh.setRightToLeft(true);
  var rule = SpreadsheetApp.newDataValidation().requireValueInList(STATUS.map(function (s) { return s[0]; }), true).build();
  sh.getRange(2, C_ST, 2000, 1).setDataValidation(rule);
  sh.setColumnWidth(C_ID, 100); sh.setColumnWidth(8, 260); sh.setColumnWidth(C_NOTE, 260); sh.setColumnWidth(C_APP, 140);
  // تلوين الحالات
  var rng = sh.getRange(2, C_ST, 2000, 1), cf = [];
  [['جديد', '#53230E', '#F6E7D2'], ['بانتظار معلومات من العميل', '#B42318', '#FFFFFF'], ['مرفوض', '#000000', '#FFFFFF'], ['مغلق', '#E7D6C0', '#53230E']]
    .forEach(function (x) { cf.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo(x[0]).setBackground(x[1]).setFontColor(x[2]).setRanges([rng]).build()); });
  sh.setConditionalFormatRules(cf);
  getFolder_();
  // تحديث "آخر تحديث" لما تغيّر الحالة أو الرسالة
  ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === 'onEditStamp') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('onEditStamp').forSpreadsheet(ss).onEdit().create();
}

function onEditStamp(e) {
  var r = e.range, sh = r.getSheet();
  if (sh.getName() !== SHEET || r.getRow() < 2) return;
  if (r.getColumn() <= C_NOTE && r.getLastColumn() >= C_ST) sh.getRange(r.getRow(), C_UPD, r.getNumRows(), 1).setValue(new Date());
}

function getFolder_() {
  var p = PropertiesService.getScriptProperties(), id = p.getProperty('folder');
  if (id) { try { return DriveApp.getFolderById(id); } catch (x) { } }
  var f = DriveApp.createFolder(FOLDER);
  p.setProperty('folder', f.getId());
  return f;
}

function out_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function clean_(s, n) { return String(s == null ? '' : s).replace(/[<>]/g, '').trim().slice(0, n || 200); }
function digits_(s) { return String(s || '').replace(/\D/g, ''); }

/* ===== استقبال بلاغ جديد ===== */
function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    var d = JSON.parse(e.postData.contents);
    var order = digits_(d.order).slice(0, 20), phone = digits_(d.phone).slice(0, 15);
    var photos = (d.photos || []).slice(0, 8);
    if (order.length < 4 || phone.length < 9 || !d.deliv || !(d.types || []).length || photos.length < 2) return out_({ ok: false, err: 'missing' });

    lock.waitLock(20000);
    var sh = SpreadsheetApp.getActive().getSheetByName(SHEET);
    var props = PropertiesService.getScriptProperties();
    var seq = Number(props.getProperty('seq') || 1000) + 1;
    props.setProperty('seq', String(seq));
    var id = 'CWC-' + seq;
    var name = clean_(d.name, 60), desc = clean_(d.desc, 600), deliv = clean_(d.deliv, 10);
    var typesAr = (d.typesAr || []).map(function (x) { return clean_(x, 40); }).join('، ');

    // تنبيه: نفس الطلب له بلاغ سابق؟
    var flag = '';
    var last = sh.getLastRow();
    if (last > 1) {
      var orders = sh.getRange(2, C_ORDER, last - 1, 1).getValues();
      var dup = orders.filter(function (r) { return String(r[0]) === order; }).length;
      if (dup) flag = 'للطلب ' + dup + ' بلاغ سابق';
    }
    var row = sh.appendRow([id, new Date(), order, name, phone, typesAr, deliv, desc, photos.length, '', 'جديد', '', new Date(), flag, '']).getLastRow();
    sh.getRange(row, C_ORDER).setNumberFormat('@').setValue(order);
    sh.getRange(row, C_PHONE).setNumberFormat('@').setValue(phone);
    lock.releaseLock();

    // الصور
    var folder = getFolder_().createFolder(id + ' — طلب ' + order);
    photos.forEach(function (b64, i) {
      folder.createFile(Utilities.newBlob(Utilities.base64Decode(b64), 'image/jpeg', id + '-' + ('0' + (i + 1)).slice(-2) + '.jpg'));
    });
    if (SHARE_PHOTOS) { try { folder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (x) { } }
    var url = folder.getUrl();
    sh.getRange(row, C_FOLDER).setValue(url);

    // نص جاهز للصق في تطبيق مركز البلاغات
    var app = 'KUFI' + JSON.stringify({ id: id, order: order, name: name, phone: phone, types: d.types, deliv: deliv, desc: desc, n: photos.length, drive: url, flag: flag });
    sh.getRange(row, C_APP).setValue(app);

    if (NOTIFY) {
      try {
        MailApp.sendEmail(Session.getEffectiveUser().getEmail(), '🔔 بلاغ عميل جديد ' + id + ' — طلب ' + order,
          'رقم البلاغ: ' + id + '\nالطلب: ' + order + '\nالعميل: ' + (name || '—') + ' · ' + phone + '\nالضرر: ' + typesAr +
          '\nتاريخ الاستلام: ' + deliv + '\nالصور (' + photos.length + '): ' + url + (desc ? '\nالوصف: ' + desc : '') + (flag ? '\n⚠️ ' + flag : '') +
          '\n\nالشيت: ' + SpreadsheetApp.getActive().getUrl());
      } catch (x) { }
    }
    return out_({ ok: true, id: id });
  } catch (err) {
    try { lock.releaseLock(); } catch (x) { }
    return out_({ ok: false, err: String(err).slice(0, 120) });
  }
}

/* ===== التتبع ===== */
function doGet(e) {
  var q = e.parameter || {};
  if (q.a !== 'track') return out_({ ok: true, ping: 'kufi-claims' });
  var id = String(q.id || '').toUpperCase().trim(), p4 = digits_(q.p);
  if (!/^CWC-\d+$/.test(id) || p4.length !== 4) return out_({ ok: false });
  var sh = SpreadsheetApp.getActive().getSheetByName(SHEET), last = sh.getLastRow();
  if (last < 2) return out_({ ok: false });
  var rows = sh.getRange(2, 1, last - 1, C_UPD).getValues();
  for (var i = rows.length - 1; i >= 0; i--) {
    var r = rows[i];
    if (String(r[C_ID - 1]).toUpperCase() !== id) continue;
    if (digits_(r[C_PHONE - 1]).slice(-4) !== p4) return out_({ ok: false });
    var st = String(r[C_ST - 1] || 'جديد'), code = 'new';
    STATUS.forEach(function (s) { if (s[0] === st) code = s[1]; });
    return out_({ ok: true, id: id, order: String(r[C_ORDER - 1]), code: code, note: String(r[C_NOTE - 1] || ''),
      at: r[C_AT - 1] ? new Date(r[C_AT - 1]).toISOString() : '', upd: r[C_UPD - 1] ? new Date(r[C_UPD - 1]).toISOString() : '' });
  }
  return out_({ ok: false });
}
