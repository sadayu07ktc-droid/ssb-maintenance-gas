/**
 * สร้าง PDF ใบแจ้งซ่อม FM-EN-04 จากแท็บ "FormTemplate" (ตรงฟอร์มเป๊ะ 1 หน้า)
 * วิธี: copy แท็บ FormTemplate -> แทนค่า {{placeholder}} -> export เป็น PDF
 *
 * ต้องมีแท็บ "FormTemplate" (import จาก FM-EN-04.xls + วาง placeholder)
 * ครั้งแรก: รัน testPdf() 1 ครั้ง เพื่ออนุญาตสิทธิ์ แล้วดูลิงก์ใน log
 */
function getPdfFolder(){
  var it = DriveApp.getFoldersByName('ssb-maintenance-pdf');
  return it.hasNext() ? it.next() : DriveApp.createFolder('ssb-maintenance-pdf');
}
function chk(v, on){ return String(v) === on ? '●' : '○'; }  // วงกลมทึบ=เลือก / โปร่ง=ไม่เลือก

// โฟลเดอร์เก็บรูปลายเซ็น (ตั้งชื่อไฟล์ = ชื่อพนักงาน หรือ รหัสพนักงาน)
var SIGN_FOLDER_NAMES = ['ลายเซ็นต์ผู้อนุมัติ','ลายเซ็นผู้อนุมัติ','signatures'];
function getSignFolder(){
  // 1) ระบุ ID ตรงๆ ใน Script Property 'SIGN_FOLDER_ID' (ชัวร์สุด ใช้ได้กับ Shared drive ด้วย)
  var id = PropertiesService.getScriptProperties().getProperty('SIGN_FOLDER_ID');
  if(id){ try{ return DriveApp.getFolderById(id); }catch(e){} }
  // 2) ชื่อตรงเป๊ะ
  for(var i=0;i<SIGN_FOLDER_NAMES.length;i++){
    var it = DriveApp.getFoldersByName(SIGN_FOLDER_NAMES[i]);
    if(it.hasNext()) return it.next();
  }
  // 3) ชื่อใกล้เคียง (เผื่อสะกด/เว้นวรรคต่างกัน)
  var s = DriveApp.searchFolders('title contains "ลายเซ" and trashed = false');
  if(s.hasNext()) return s.next();
  // 4) ไม่มีเลย -> สร้างเอง (สคริปต์เป็นเจ้าของ จึงเข้าถึงได้แน่นอน)
  return DriveApp.createFolder('ssb-signatures');
}
/** วางรูปลายเซ็นให้พอดีช่อง (รองรับเซลล์ที่ merge) */
function placeSign(tmp, blob, a1){
  var rng = tmp.getRange(a1);
  var m = rng.getMergedRanges()[0] || rng;
  var w = 0, h = 0, c, r;
  for(c = m.getColumn(); c < m.getColumn() + m.getNumColumns(); c++) w += tmp.getColumnWidth(c);
  for(r = m.getRow();    r < m.getRow()    + m.getNumRows();    r++) h += tmp.getRowHeight(r);
  var img = tmp.insertImage(blob, m.getColumn(), m.getRow());
  var iw = img.getWidth(), ih = img.getHeight();
  var maxW = Math.max(w - 12, 40), maxH = Math.max(h - 6, 26);
  var s = Math.min(maxW / iw, maxH / ih);
  if(s > 3) s = 3;   // รูปเล็กมากก็ไม่ขยายจนแตก
  var nw = Math.round(iw * s), nh = Math.round(ih * s);
  img.setWidth(nw).setHeight(nh);
  img.setAnchorCellXOffset(Math.round((w - nw) / 2));
  img.setAnchorCellYOffset(Math.round(Math.max((h - nh) / 2, 0)));
  return img;
}
/** หารูปลายเซ็นของ line_user_id นี้ · จับคู่ด้วยชื่อไฟล์ = full_name หรือ emp_code */
function signBlobByLine(lid){
  if(!lid) return null;
  var e = getRows(SHEETS.EMP).filter(function(r){ return String(r.line_user_id) === String(lid); })[0];
  if(!e) return null;
  // 1) อัปโหลดผ่านแอปไว้แล้ว -> ใช้ไฟล์นั้นเลย (ชัวร์สุด ไม่ต้องเดาชื่อไฟล์)
  if(e.signature_file_id){
    try{ return DriveApp.getFileById(String(e.signature_file_id)).getBlob(); }catch(err){}
  }
  // 2) ไม่มี -> เดาจากชื่อไฟล์ในโฟลเดอร์
  var folder = getSignFolder(); if(!folder) return null;
  var keys = [e.full_name, e.emp_code, e.id].filter(function(k){ return k && String(k).trim(); })
    .map(function(k){ return String(k).trim().toLowerCase(); });
  var it = folder.getFiles();
  while(it.hasNext()){
    var f = it.next();
    if(String(f.getMimeType()).indexOf('image/') !== 0) continue;
    var base = f.getName().replace(/\.[^.]+$/, '').trim().toLowerCase();
    for(var i=0;i<keys.length;i++){
      if(base === keys[i] || base.indexOf(keys[i]) >= 0 || keys[i].indexOf(base) >= 0) return f.getBlob();
    }
  }
  return null;
}
// แปลง line_user_id -> ชื่อพนักงาน (จากแท็บ Employees)
function empNameByLine(lid){
  if(!lid) return '';
  var e = getRows(SHEETS.EMP).filter(function(r){ return String(r.line_user_id) === String(lid); })[0];
  return e ? (e.full_name || String(lid)) : String(lid);
}
function money(v){ v = String(v==null?'':v).replace(/,/g,'');
  if(v==='' || isNaN(v)) return '';
  var n = Number(v);
  return n.toLocaleString('en-US', { minimumFractionDigits: (n % 1 ? 2 : 0), maximumFractionDigits: 2 });
}
function ddmmyyyy(d){ var p = String(d||'').slice(0,10).split('-'); return p.length===3 ? (p[2]+'-'+p[1]+'-'+p[0]) : String(d||''); }

function kindLine(v){ return chk(v,'repair')+' ซ่อม   '+chk(v,'replace_part')+' เปลี่ยนอะไหล่   '+chk(v,'inspect')+' ตรวจเช็ค   '+chk(v,'tire')+' เปลี่ยนยาง   '+chk(v,'install')+' ติดตั้ง   '+chk(v,'other')+' อื่นๆ'; }
function catLine(v){ return chk(v,'machine')+' เครื่องจักร   '+chk(v,'electrical')+' ระบบไฟฟ้า   '+chk(v,'building')+' อาคาร-สถานที่   '+chk(v,'vehicle')+' รถ   '+chk(v,'other')+' อื่นๆ'; }

/** สร้างชุดค่าที่จะพิมพ์ลงฟอร์ม (แยกออกมาเพื่อให้ตรวจสอบได้โดยไม่ต้องสร้างไฟล์) */
function pdfMap(ticketNo){
  var r = getRows(SHEETS.REQ).filter(function(x){ return x.ticket_no === ticketNo; })[0];
  if(!r) throw 'ไม่พบใบ ' + ticketNo;
  var veh = r.vehicle_key ? getRows(SHEETS.VEH).filter(function(v){ return v.vehicle_key === r.vehicle_key; })[0] : null;
  var isVeh = !!r.vehicle_key;

  // อาการ (รวมข้อมูลรถ + เงิน เพราะฟอร์มเดิมไม่มีช่องไมล์/เงิน)
  var sym = r.symptom || '';
  if(isVeh){
    var extra = [];
    if(r.mileage) extra.push('เลขไมล์ ' + money(r.mileage));
    if(r.service_interval_km) extra.push('ระยะเช็ค ' + money(r.service_interval_km) + ' กม.');
    if(r.vendor) extra.push('ศูนย์/อู่ ' + r.vendor);
    if(extra.length) sym = (sym ? sym + '  ' : '') + '[' + extra.join(' · ') + ']';
  }
  // จำนวนเงิน -> ต่อท้ายช่อง "การแก้ไข" (ไม่ใช่ช่องอาการ)
  var fix = r.fix_detail || '';
  if(r.amount) fix += (fix ? '   ' : '') + 'จำนวนเงิน ' + money(r.amount) + ' บาท';

  var apprvName = r.approver_id ? empNameByLine(r.approver_id) : '';
  // ไม่ต่อวันที่ตรงชื่อ — มีช่อง "วันที่" แยกอยู่แล้ว
  var apprv = r.approved_at ? apprvName : (apprvName + '   (รออนุมัติ)');
  var repTime = ''; var _mt = String(r.reported_at||'').match(/\d{1,2}:\d{2}/); if(_mt) repTime = _mt[0];

  var map = {
    'ผู้แจ้ง':      r.requester_name || '',
    'หน่วยงาน':     r.department || '',
    'เลขที่':       r.ticket_no,
    'วันที่':       ddmmyyyy(r.reported_at),
    'เวลา':         repTime ? (repTime + ' น.') : '',
    'ซ่อมใน':       chk(r.repair_by,'internal'),
    'ซ่อมนอก':      chk(r.repair_by,'external'),
    'อู่':          r.vendor ? ('(' + r.vendor + ')') : '',
    'วันที่อนุมัติ': ddmmyyyy(r.approved_at),
    'รหัส':         isVeh ? (veh ? veh.plate_current : r.vehicle_key)
                          : (String(r.asset_category) === 'building' && r.machine_code ? ('ห้อง ' + r.machine_code) : (r.machine_code || '')),
    'ชื่อเครื่อง':   isVeh ? (veh ? veh['ยี่ห้อ_รุ่น'] : '') : (r.machine_name || ''),
    'ประเภทคำขอ':   kindLine(r.request_kind),
    'ประเภทงาน':    catLine(r.asset_category),
    'อาการ':        sym,
    'ผู้แจ้งเซ็น':   r.requester_name || '',
    'ผู้อนุมัติ':    apprv,
    'สาเหตุ':       r.cause || '',
    'การแก้ไข':     fix
  };
  return map;
}

function genPdf(ticketNo){
  var r = getRows(SHEETS.REQ).filter(function(x){ return x.ticket_no === ticketNo; })[0];
  if(!r) throw 'ไม่พบใบ ' + ticketNo;
  var map = pdfMap(ticketNo);
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var tpl = ss.getSheetByName('FormTemplate');
  if(!tpl) throw 'ไม่พบแท็บ FormTemplate (import ฟอร์ม + วาง placeholder ก่อน)';
  var apprvName = r.approver_id ? empNameByLine(r.approver_id) : '';

  // copy แท็บ -> แทนค่า
  var tmpName = '__pdf_' + ticketNo;
  var ex = ss.getSheetByName(tmpName); if(ex) ss.deleteSheet(ex);
  var tmp = tpl.copyTo(ss).setName(tmpName);
  SpreadsheetApp.flush();
  // ช่องวันที่ต้องเป็น "ข้อความ" ก่อนแทนค่า ไม่งั้น Sheets แปลง 21-07-2026 เป็น Date แล้วโชว์ Tue Jul 21 2026
  ['S14','G4','O33','O37'].forEach(function(a1){ try{ tmp.getRange(a1).setNumberFormat('@'); }catch(e){} });
  Object.keys(map).forEach(function(k){ tmp.createTextFinder('{{' + k + '}}').replaceAllWith(String(map[k])); });
  SpreadsheetApp.flush();

  // ล้าง rich-text ที่ import จาก Excel + ตั้งขนาดฟอนต์ช่องค่าให้เท่ากัน (กันฟอนต์เพี้ยน)
  var VAL_FONT = 11;
  ['G3','M3','R3','G4','N4','H5','N5','E7','K7','E8','E9','E10','B14','M14','S14','B16','B22'].forEach(function(a1){
    try{
      var c = tmp.getRange(a1), v = c.getValue();
      // ถ้า Sheets แปลงเป็นวันที่ไปแล้ว ให้เขียนกลับเป็น dd-mm-yyyy (ไม่ใช่ Tue Jul 21 2026 …)
      v = (v instanceof Date) ? Utilities.formatDate(v, TZ, 'dd-MM-yyyy') : String(v == null ? '' : v);
      c.setNumberFormat('@'); c.setValue(v);
      c.setFontSize(VAL_FONT).setFontFamily('Sarabun');
    }catch(e){}
  });
  SpreadsheetApp.flush();

  // ลายเซ็นผู้อนุมัติ -> ช่อง "ผู้อนุมัติ" ท่อนล่าง (F37) + วันที่ (O37) · ใส่เมื่ออนุมัติแล้วเท่านั้น
  if(r.approved_at && r.approver_id){
    try{
      var sig = signBlobByLine(r.approver_id);
      // แถวลายเซ็นเตี้ยเกินไป (~20px) รูปจะเล็กจนอ่านไม่ออก -> ขยายแถวก่อนวาง
      if(sig && tmp.getRowHeight(37) < 52) tmp.setRowHeight(37, 52);
      if(sig) placeSign(tmp, sig, 'F37');
      else    tmp.getRange('F37').setValue(apprvName);   // ไม่มีรูป -> พิมพ์ชื่อแทน
      tmp.getRange('O37').setNumberFormat('@').setValue(ddmmyyyy(r.approved_at));
      SpreadsheetApp.flush();
    }catch(e){}
  }

  // export แท็บนั้นเป็น PDF (ขนาด/ฟอนต์ = ตามที่จัดในแท็บ FormTemplate เป๊ะ ไม่มีโค้ดทับ)
  var url = 'https://docs.google.com/spreadsheets/d/' + ss.getId() + '/export?format=pdf'
    + '&gid=' + tmp.getSheetId()
    + '&size=A4&portrait=true&fitw=true&gridlines=false&sheetnames=false&printtitle=false&pagenumbers=false&fzr=false'
    + '&top_margin=0.4&bottom_margin=0.4&left_margin=0.4&right_margin=0.4';
  var resp = UrlFetchApp.fetch(url, { headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() } });
  var pdf = resp.getBlob().setName('ใบแจ้งซ่อม_' + ticketNo + '.pdf');

  var folder = getPdfFolder();
  var old = folder.getFilesByName('ใบแจ้งซ่อม_' + ticketNo + '.pdf'); while(old.hasNext()){ old.next().setTrashed(true); }
  var file = folder.createFile(pdf);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  ss.deleteSheet(tmp); // ลบแท็บชั่วคราว

  var link = 'https://drive.google.com/file/d/' + file.getId() + '/view';
  patchByTicket(SHEETS.REQ, ticketNo, { pdf_url: link, updated_at: now() });
  return link;
}

// สำรวจตำแหน่งเซลล์ในฟอร์ม (รันแล้วก๊อป log ส่งให้ผม)
function dumpTemplate(){
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('FormTemplate');
  if(!sh){ Logger.log('ไม่พบแท็บ FormTemplate'); return; }
  var vals = sh.getDataRange().getValues();
  var out = [];
  for(var i=0;i<vals.length;i++){
    for(var j=0;j<vals[i].length;j++){
      var v = String(vals[i][j]).trim();
      if(v!=='') out.push(sh.getRange(i+1,j+1).getA1Notation()+' = '+v);
    }
  }
  var merges = sh.getDataRange().getMergedRanges().map(function(m){ return m.getA1Notation(); });
  Logger.log('==== CELLS ====\n' + out.join('\n') + '\n\n==== MERGES ====\n' + merges.join('   '));
}

// วาง placeholder ลงฟอร์มอัตโนมัติ (รันครั้งเดียว)
function setupFormPlaceholders(){
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('FormTemplate');
  if(!sh) throw 'ไม่พบแท็บ FormTemplate';
  try{ sh.getRange('E10:S13').merge(); }catch(e){}   // รวมกล่องอาการเป็นช่องเดียว
  sh.createTextFinder('¦').replaceAllWith('○');       // เปลี่ยน ¦ ที่เหลือเป็น ○ วงกลม
  var set = {
    'G3': 'ชื่อผู้แจ้ง {{ผู้แจ้ง}}',
    'M3': 'หน่วยงาน {{หน่วยงาน}}',
    'R3': 'เลขที่ {{เลขที่}}',
    'G4': 'วันที่แจ้งซ่อม {{วันที่}}',
    'N4': 'เวลาแจ้งซ่อม {{เวลา}}',
    'H5': '{{ซ่อมใน}} ซ่อมโดยช่างภายใน',
    'N5': '{{ซ่อมนอก}} ติดต่อผู้ซ่อมจากภายนอก {{อู่}}',
    'E7': '{{รหัส}}',
    'K7': '{{ชื่อเครื่อง}}',
    'E8': '{{ประเภทคำขอ}}',
    'E9': '{{ประเภทงาน}}',
    'E10': '{{อาการ}}',
    'B14': '{{ผู้แจ้งเซ็น}}',
    'M14': '{{ผู้อนุมัติ}}',
    'S14': '{{วันที่อนุมัติ}}',
    'B16': '{{สาเหตุ}}',
    'B22': '{{การแก้ไข}}'
  };
  Object.keys(set).forEach(function(a1){ sh.getRange(a1).setValue(set[a1]); });
  sh.getRange('E10').setWrap(true);
  // ฟอนต์/ขนาด: จัดเองในแท็บ FormTemplate ได้เลย โค้ดไม่ไปยุ่งแล้ว
  return 'วาง placeholder ' + Object.keys(set).length + ' ช่อง เสร็จ';
}

function testPdf(){
  var r = getRows(SHEETS.REQ)[0];
  if(!r){ Logger.log('ยังไม่มีใบแจ้งซ่อม — ส่งสักใบจากแอปก่อน'); return; }
  var url = genPdf(r.ticket_no);
  Logger.log('PDF พร้อม: ' + url);
  return url;
}
