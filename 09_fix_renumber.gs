/**
 * แก้เลขที่ใบซ้ำ เดือน ก.ย. 2569 (ใช้ครั้งเดียว)
 * Requests + MaintenanceRecords แก้มือแล้ว: ลบใบ 019/030/031 เดิม · 032 → 030 · 033 → 019
 * สคริปต์นี้ตามแก้ส่วนที่เหลือให้ตรงกัน: StatusLogs · ไฟล์ PDF · Supabase tasks
 *
 * วิธีใช้: เลือก renumberCheck → Run → ดู Execution log (ยังไม่แก้อะไร)
 *          ถ้าตัวเลขถูกต้อง เลือก renumberRun → Run
 */
var RN_MAP = { 'HRC-202609-032': 'HRC-202609-030', 'HRC-202609-033': 'HRC-202609-019' };
var RN_DELETED = ['HRC-202609-019', 'HRC-202609-030', 'HRC-202609-031'];
var RN_DEL_TAG = '-ลบแล้ว';

function renumberCheck() { renumber_(true); }
function renumberRun() { renumber_(false); }

function renumber_(dry) {
  var log = function (s) { console.log((dry ? '[ตรวจ] ' : '[แก้] ') + s); };
  var lock = LockService.getScriptLock(); lock.waitLock(30000);
  try {
    // 1) ตรวจว่า Requests แก้ครบแล้วจริง
    var req = getRows(SHEETS.REQ);
    var cnt = function (t) { return req.filter(function (r) { return String(r.ticket_no) === t; }).length; };
    var bad = false;
    Object.keys(RN_MAP).forEach(function (o) {
      var n = RN_MAP[o];
      log('Requests: ' + o + ' เหลือ ' + cnt(o) + ' แถว · ' + n + ' มี ' + cnt(n) + ' แถว');
      if (cnt(o) !== 0 || cnt(n) !== 1) bad = true;
    });
    if (cnt('HRC-202609-031') !== 0) { log('Requests: ยังมี HRC-202609-031 อยู่ ' + cnt('HRC-202609-031') + ' แถว'); bad = true; }
    if (bad) { log('❌ Requests ยังไม่ตรงกับที่คาดไว้ หยุดก่อน ไม่แก้อะไร'); return; }

    var mr = getRows(SHEETS.HIST);
    Object.keys(RN_MAP).forEach(function (o) {
      var n = RN_MAP[o];
      var a = mr.filter(function (r) { return String(r['เลขที่ใบแจ้งซ่อม']) === o; }).length;
      var b = mr.filter(function (r) { return String(r['เลขที่ใบแจ้งซ่อม']) === n; }).length;
      log('MaintenanceRecords: ' + o + ' เหลือ ' + a + ' · ' + n + ' มี ' + b + (a ? '  ⚠️ ยังมีเลขเก่าค้าง' : ''));
    });

    // 2) StatusLogs — ประวัติของใบที่ลบติดป้าย -ลบแล้ว ก่อน แล้วค่อยย้ายประวัติ 032/033 ไปเลขใหม่
    var s = sh(SHEETS.LOG), vals = s.getDataRange().getValues(), col = vals[0].indexOf('ticket_no');
    var moved = {}, tagged = {};
    var oldLeft = 0;
    for (var i = 1; i < vals.length; i++) if (RN_MAP[String(vals[i][col])]) oldLeft++;
    if (!oldLeft) log('StatusLogs: ไม่พบประวัติของ 032/033 แล้ว (อาจแก้ไปแล้ว) ข้ามขั้นนี้');
    else {
      for (var j = 1; j < vals.length; j++) {
        var t = String(vals[j][col]), nv = null;
        if (RN_DELETED.indexOf(t) >= 0) { nv = t + RN_DEL_TAG; tagged[t] = (tagged[t] || 0) + 1; }
        else if (RN_MAP[t]) { nv = RN_MAP[t]; moved[t] = (moved[t] || 0) + 1; }
        if (nv && !dry) s.getRange(j + 1, col + 1).setValue(nv);
      }
      Object.keys(tagged).forEach(function (t) { log('StatusLogs: ประวัติใบที่ลบ ' + t + ' ' + tagged[t] + ' แถว → ' + t + RN_DEL_TAG); });
      Object.keys(moved).forEach(function (t) { log('StatusLogs: ' + t + ' ' + moved[t] + ' แถว → ' + RN_MAP[t]); });
    }

    // 3) ไฟล์ PDF เก่า — เปลี่ยนชื่อเก็บไว้ (ไม่ลบ) กัน genPdf ทิ้งไฟล์ของใบที่ลบ
    var folder = getPdfFolder();
    var renameFile = function (ticket, suffix) {
      var it = folder.getFilesByName('ใบแจ้งซ่อม_' + ticket + '.pdf'), n = 0;
      while (it.hasNext()) { var f = it.next(); n++; if (!dry) f.setName('ใบแจ้งซ่อม_' + ticket + suffix + '.pdf'); }
      log('PDF: ใบแจ้งซ่อม_' + ticket + '.pdf ' + n + ' ไฟล์ → เปลี่ยนชื่อเป็น ' + ticket + suffix);
    };
    RN_DELETED.forEach(function (t) { renameFile(t, RN_DEL_TAG); });
    Object.keys(RN_MAP).forEach(function (o) { renameFile(o, '_เลขเก่า(เป็น ' + RN_MAP[o].slice(-3) + ')'); });

    // 4) Supabase tasks (external_ref = ssb:<เลขที่>)
    var c = sbProps();
    if (!c.url || !c.key) log('Supabase: ไม่ได้ตั้งค่า ข้าม');
    else {
      var sbPatch = function (from, to) {
        var q = c.url + '/rest/v1/tasks?external_ref=eq.' + encodeURIComponent('ssb:' + from);
        var got = JSON.parse(UrlFetchApp.fetch(q + '&select=id', { headers: sbHeaders(c.key), muteHttpExceptions: true }).getContentText() || '[]');
        var n = Array.isArray(got) ? got.length : 0;
        if (n && !dry) {
          var h = sbHeaders(c.key); h['Content-Type'] = 'application/json';
          UrlFetchApp.fetch(q, { method: 'patch', headers: h, payload: JSON.stringify({ external_ref: 'ssb:' + to }), muteHttpExceptions: true });
        }
        log('Supabase: ssb:' + from + ' ' + n + ' งาน → ssb:' + to);
      };
      RN_DELETED.forEach(function (t) { sbPatch(t, t + RN_DEL_TAG); });
      Object.keys(RN_MAP).forEach(function (o) { sbPatch(o, RN_MAP[o]); });
    }

    // 5) สร้าง PDF ใหม่ด้วยเลขใหม่ + บันทึกประวัติการเปลี่ยนเลข
    Object.keys(RN_MAP).forEach(function (o) {
      var n = RN_MAP[o];
      var r = getRows(SHEETS.REQ).filter(function (x) { return String(x.ticket_no) === n; })[0];
      if (dry) { log('PDF: จะสร้างใหม่ ใบแจ้งซ่อม_' + n + '.pdf (สถานะ ' + r.status + ')'); return; }
      if (String(r.status) === 'approved') log('PDF: ' + n + ' → ' + genPdf(n));
      else log('PDF: ' + n + ' สถานะ ' + r.status + ' ยังไม่อนุมัติ ไม่สร้าง PDF');
      logStatus(n, r.status, r.status, 'script', 'เปลี่ยนเลขที่จาก ' + o + ' เป็น ' + n + ' (แก้ใบซ้ำ)');
    });
    log(dry ? '✅ ตรวจเสร็จ ถ้าตัวเลขถูกต้อง ให้ Run renumberRun' : '✅ แก้เสร็จแล้ว');
  } finally { lock.releaseLock(); }
}
