// ===== ปริ้นใบงานหลายใบในครั้งเดียว =====
// หน้าเว็บขอไฟล์ PDF ทีละชุด (pdf_batch) → รวมเป็นไฟล์เดียวในเบราว์เซอร์ (pdf-lib)
// → ส่งไฟล์รวมกลับมาเก็บใน Drive (pdf_batch_save) แล้วเปิดลิงก์ปริ้น (ใช้ได้ทั้งในไลน์และคอม)

API.pdf_batch = function(p){
  denyIf(!isPrivLine(p.actor), 'เฉพาะแอดมิน/ผู้อนุมัติ');
  var want = [].concat(p.ticket_nos || []).slice(0, 5);   // ทีละไม่เกิน 5 ใบ กันหมดเวลา
  var rows = getRows(SHEETS.REQ);
  return want.map(function(t){
    var r = rows.filter(function(x){ return String(x.ticket_no) === String(t); })[0];
    if(!r) return { ticket_no: t, error: 'ไม่พบใบ' };
    var url = String(r.pdf_url || '');
    if(!/^https?:/.test(url)){
      if(String(r.status) !== 'approved') return { ticket_no: t, error: 'ยังไม่อนุมัติ จึงยังไม่มีใบงาน' };
      url = genPdf(t);
    }
    var m = url.match(/\/d\/([\w-]{20,})/) || url.match(/[?&]id=([\w-]{20,})/);
    if(!m) return { ticket_no: t, error: 'ลิงก์ PDF ไม่ถูกต้อง' };
    try {
      return { ticket_no: t, b64: Utilities.base64Encode(DriveApp.getFileById(m[1]).getBlob().getBytes()) };
    } catch(e){ return { ticket_no: t, error: 'เปิดไฟล์ PDF ไม่ได้' }; }
  });
};

API.pdf_batch_save = function(p){
  denyIf(!isPrivLine(p.actor), 'เฉพาะแอดมิน/ผู้อนุมัติ');
  denyIf(!p.b64, 'ไม่มีไฟล์');
  var folder = pdfBatchFolder_();
  // ไฟล์รวมใช้แค่ปริ้น เก็บไว้ 3 วันแล้วย้ายลงถังขยะ
  var cut = Date.now() - 3 * 864e5, it = folder.getFiles();
  while(it.hasNext()){ var f = it.next(); if(f.getDateCreated().getTime() < cut) f.setTrashed(true); }
  var name = 'ใบงานรวม_' + Utilities.formatDate(new Date(), TZ, 'yyyyMMdd-HHmmss') + '_' + (Number(p.count) || '') + 'ใบ.pdf';
  var file = folder.createFile(Utilities.newBlob(Utilities.base64Decode(p.b64), 'application/pdf', name));
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return { url: 'https://drive.google.com/file/d/' + file.getId() + '/view', name: name };
};

function pdfBatchFolder_(){
  var parent = getPdfFolder(), it = parent.getFoldersByName('ใบงานรวม (ปริ้น)');
  return it.hasNext() ? it.next() : parent.createFolder('ใบงานรวม (ปริ้น)');
}
