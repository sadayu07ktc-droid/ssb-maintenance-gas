/**
 * ssb-maintenance — GAS backend (Phase 2)
 * ผูกกับ Google Sheet "ssb-maintenance-db"
 *
 * วิธีติดตั้ง:
 *  1) ในชีต: Extensions → Apps Script → ลบโค้ดเดิม แล้ววางไฟล์นี้ทั้งหมด → Save
 *  2) รันฟังก์ชัน setup() 1 ครั้ง (กด Run, อนุญาตสิทธิ์) → สร้างหัวคอลัมน์ให้ Requests/Employees/StatusLogs
 *  3) Deploy → New deployment → Web app → Execute as: Me, Who has access: Anyone → คัดลอก URL
 *  4) ทดสอบ: เปิด <URL>?action=ping  และ  <URL>?action=vehicles  ในเบราว์เซอร์ (ควรได้ JSON)
 */

// ===== CONFIG =====
var TZ = 'Asia/Bangkok';
var SHEETS = {
  VEH: 'Vehicles', HIST: 'MaintenanceRecords',
  REQ: 'Requests', EMP: 'Employees', LOG: 'StatusLogs'
};

var HEADERS = {
  Requests: ['ticket_no','product_line','requester_id','requester_name','department','reported_at',
    'request_type','repair_by','asset_category','request_kind','machine_code','machine_name',
    'vehicle_key','symptom','mileage','service_interval_km','vendor','receipt_no','amount',
    'receipt_urls','symptom_urls','cause','cause_type','fix_detail','due_date','prevention',
    'contamination','technician1','technician2','repair_start','repair_finish','received_by',
    'reviewed_by','reviewed_at','accepted_by','accepted_at','approver_id','approved_at',
    'rejected_reason','status','pdf_url','created_at','updated_at','assignee_id','assignee_name'],
  Employees: ['id','line_user_id','emp_code','full_name','department','phone','role','active'],
  StatusLogs: ['id','ticket_no','from_status','to_status','actor','note','created_at']
};

// ===== SHEET HELPERS =====
function getReceiptFolder(){
  var it = DriveApp.getFoldersByName('ssb-maintenance-receipts');
  return it.hasNext() ? it.next() : DriveApp.createFolder('ssb-maintenance-receipts');
}
function ss(){ return SpreadsheetApp.getActiveSpreadsheet(); }
function sh(name){ return ss().getSheetByName(name); }
function now(){ return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss'); }

/**
 * ลงประวัติซ่อมรถ (MaintenanceRecords) โดยอ้างอิงด้วย "เลขที่ใบแจ้งซ่อม" กัน record ซ้ำ
 * create=true  -> เขียนแถวใหม่ถ้ายังไม่มี (ใช้ตอน "อนุมัติ" เท่านั้น)
 * create=false -> อัปเดตแถวเดิมอย่างเดียว (แก้ไขข้อมูล / แนบบิลทีหลัง) ใบที่ยังไม่อนุมัติจะไม่ถูกบันทึก
 */
function syncHistory(ticketNo, create){
  var r = getRows(SHEETS.REQ).filter(function(x){ return x.ticket_no === ticketNo; })[0];
  if(!r || !r.vehicle_key) return;   // ชีตนี้เก็บประวัติของรถเท่านั้น
  var veh = getRows(SHEETS.VEH).filter(function(v){ return v.vehicle_key === r.vehicle_key; })[0] || {};
  var s = sh(SHEETS.HIST);
  var head = s.getRange(1, 1, 1, s.getLastColumn()).getValues()[0];
  var tcol = head.indexOf('เลขที่ใบแจ้งซ่อม');
  if(tcol < 0) return;
  var vals = s.getDataRange().getValues();
  var row = -1;
  for(var i = 1; i < vals.length; i++){ if(String(vals[i][tcol]) === String(ticketNo)){ row = i + 1; break; } }
  if(row < 0 && !create) return;

  var data = {
    'vehicle_key':      r.vehicle_key,
    'ทะเบียนรถ':        veh.plate_current || '',
    'plate_type':       veh.plate_type || '',
    'วันที่ซ่อม':        String(r.approved_at || now()).slice(0, 10),
    'เลขที่ใบแจ้งซ่อม':  ticketNo,
    'ระยะเช็ค_กม':      r.service_interval_km || '',
    'เลขไมล์':          r.mileage || '',
    'รายการซ่อม':       r.fix_detail || r.symptom || '',
    'จำนวนเงิน':        r.amount || '',
    'ไฟล์แนบ':          r.receipt_urls || ''
  };
  if(row < 0){ row = s.getLastRow() + 1; data.id = uuid(); }
  Object.keys(data).forEach(function(k){
    var c = head.indexOf(k);
    if(c >= 0) s.getRange(row, c + 1).setValue(data[k]);
  });
}
function uuid(){ return Utilities.getUuid(); }

/**
 * ประวัติซ่อมของรถ 1 คัน — จับคู่ด้วย vehicle_key "หรือ" ทะเบียนใดก็ได้ของคันนั้น
 * (รถเปลี่ยนทะเบียน เช่น ก 4080-1 -> ขง 9219 แถวเก่า/ใหม่จะติดคนละคีย์ ประวัติเลยขาดเป็นท่อน)
 */
function normPlate(s){ return String(s == null ? '' : s).replace(/\s|-/g, '').toLowerCase(); }
function histOf(vk){
  if(!vk) return [];
  var v = getRows(SHEETS.VEH).filter(function(x){ return x.vehicle_key === vk; })[0] || {};
  var plates = {};
  [v.plate_current, v.plate_red, v.plate_all].forEach(function(p){
    String(p || '').split(/\s*[,/]\s*/).forEach(function(one){ if(normPlate(one)) plates[normPlate(one)] = 1; });
  });
  return getRows(SHEETS.HIST).filter(function(r){
    if(r.vehicle_key === vk) return true;
    if(r.vehicle_key) return false;               // ติดคีย์คันอื่นแล้ว ไม่ใช่ของเรา
    return !!plates[normPlate(r['ทะเบียนรถ'])];   // ยังไม่มีคีย์ -> เทียบด้วยทะเบียน
  });
}
function getRows(name){
  var s = sh(name); if(!s) return [];
  var vals = s.getDataRange().getValues();
  if(vals.length < 2) return [];
  var head = vals[0];
  var out = [];
  for(var i=1;i<vals.length;i++){
    if(vals[i].join('') === '') continue;
    var o = {};
    for(var c=0;c<head.length;c++){
      var val = vals[i][c];
      if(val instanceof Date){ var _t = Utilities.formatDate(val, TZ, 'HH:mm'); val = Utilities.formatDate(val, TZ, 'yyyy-MM-dd') + (_t==='00:00' ? '' : ' '+_t); }
      o[head[c]] = val;
    }
    o._row = i + 1;
    out.push(o);
  }
  return out;
}
function appendObj(name, obj){
  var s = sh(name);
  var head = s.getRange(1,1,1,s.getLastColumn()).getValues()[0];
  var row = head.map(function(h){ return obj[h] !== undefined ? obj[h] : ''; });
  s.appendRow(row);
}
function patchByTicket(name, ticket, patch){
  var s = sh(name);
  var vals = s.getDataRange().getValues();
  var head = vals[0];
  var tcol = head.indexOf('ticket_no');
  for(var i=1;i<vals.length;i++){
    if(String(vals[i][tcol]) === String(ticket)){
      for(var c=0;c<head.length;c++){
        if(patch[head[c]] !== undefined) s.getRange(i+1, c+1).setValue(patch[head[c]]);
      }
      return true;
    }
  }
  return false;
}

// เพิ่มคอลัมน์ assignee ในแท็บ Requests ถ้ายังไม่มี
function ensureReqCols(){
  var s = sh(SHEETS.REQ);
  var head = s.getRange(1,1,1,Math.max(1,s.getLastColumn())).getValues()[0];
  ['assignee_id','assignee_name'].forEach(function(c){
    if(head.indexOf(c) < 0){ s.getRange(1, s.getLastColumn()+1).setValue(c); head.push(c); }
  });
}

// ===== SETUP (รัน 1 ครั้ง) =====
function setup(){
  Object.keys(HEADERS).forEach(function(name){
    var s = sh(name);
    if(!s) s = ss().insertSheet(name);
    if(s.getLastRow() === 0){ s.getRange(1,1,1,HEADERS[name].length).setValues([HEADERS[name]]); }
  });
  return 'setup done';
}

// ===== TICKET NO =====
/**
 * เลขที่ใบ = ต่อจาก "เลขสูงสุดที่เคยใช้" ของเดือนนั้น (ไม่ใช่นับจำนวนแถว)
 * ดูทั้ง Requests และ StatusLogs เพราะถ้าลบแถวใน Requests ทิ้ง การนับแถวจะถอยกลับ
 * แล้วออกเลขซ้ำ -> ประวัติ/PDF ของคนละใบจะปนกัน
 */
function ticketNo(){
  var ym = Utilities.formatDate(new Date(), TZ, 'yyyyMM');
  var prefix = 'HRC-' + ym + '-';
  var max = 0;
  function scan(rows){
    rows.forEach(function(r){
      var t = String(r.ticket_no || '');
      if(t.indexOf(prefix) !== 0) return;
      var n = parseInt(t.slice(prefix.length), 10);
      if(!isNaN(n) && n > max) max = n;
    });
  }
  scan(getRows(SHEETS.REQ));
  scan(getRows(SHEETS.LOG));
  return prefix + ('000' + (max + 1)).slice(-3);
}
function logStatus(ticket, from, to, actor, note){
  appendObj(SHEETS.LOG, { id: uuid(), ticket_no: ticket, from_status: from||'', to_status: to||'',
    actor: actor||'', note: note||'', created_at: now() });
}

// ===== API ROUTER =====
function json(o){ return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function doGet(e){ e = e || {}; return route((e.parameter||{}).action, e.parameter||{}); }
function doPost(e){
  var b = {};
  try { b = JSON.parse(e.postData.contents); } catch(_) { b = (e && e.parameter) || {}; }
  // LINE webhook ส่ง { destination, events:[...] } มา — ไม่ใช่ API ของแอปเรา
  if(b && b.events){ try{ handleLineEvents(b.events); }catch(err){} return json({ ok:true }); }
  return route(b.action, b);
}
function route(action, p){
  try {
    var fn = API[action];
    if(!fn) return json({ ok:false, error:'unknown action: ' + action });
    return json({ ok:true, data: fn(p||{}) });
  } catch(err){ return json({ ok:false, error: String(err) }); }
}

var API = {
  ping: function(){ return { pong: now() }; },
  me: function(p){
    var u = getRows(SHEETS.EMP).filter(function(r){ return r.line_user_id === p.line_user_id; })[0];
    return u || { role: 'requester', unknown: true };
  },
  vehicles: function(){ return getRows(SHEETS.VEH).map(strip); },
  vehicle_history: function(p){
    return histOf(p.vehicle_key)
      .sort(function(a,b){ return String(b['วันที่ซ่อม']).localeCompare(String(a['วันที่ซ่อม'])); })
      .map(strip);
  },
  vehicle_summary: function(p){
    var rows = histOf(p.vehicle_key);
    var total = rows.reduce(function(s,r){ return s + (Number(r['จำนวนเงิน'])||0); }, 0);
    var last = rows.sort(function(a,b){ return String(b['วันที่ซ่อม']).localeCompare(String(a['วันที่ซ่อม'])); })[0] || null;
    return { times: rows.length, total: total, last: last ? strip(last) : null };
  },
  submit_request: function(p){
    var t = ticketNo();
    var rec = { ticket_no: t, reported_at: now(), created_at: now(), updated_at: now() };
    HEADERS.Requests.forEach(function(h){ if(p[h] !== undefined) rec[h] = p[h]; });
    rec.ticket_no = t;
    // แอดมินแจ้งเอง + ไม่ใช่งานอาคาร -> ส่งให้ผู้อนุมัติเลย (ไม่ต้องผ่านแอดมินตรวจซ้ำ)
    var emp = getRows(SHEETS.EMP).filter(function(r){ return String(r.line_user_id) === String(p.requester_id); })[0];
    var byAdmin = emp && /admin/i.test(String(emp.role||''));
    rec.status = (byAdmin && String(p.asset_category) !== 'building') ? 'pending_approval' : 'submitted';
    appendObj(SHEETS.REQ, rec);
    logStatus(t, '', rec.status, p.requester_id || '', byAdmin ? 'แอดมินแจ้งเอง → ส่งอนุมัติเลย' : '');
    if(rec.status === 'pending_approval') notifyApprovalCard(t, false);
    return { ticket_no: t, status: rec.status };
  },
  pending_approvals: function(){
    return getRows(SHEETS.REQ).filter(function(r){ return r.status === 'pending_approval'; }).map(strip);
  },
  requests: function(p){
    var rows = getRows(SHEETS.REQ);
    if(p.status){ var set = String(p.status).split(','); rows = rows.filter(function(r){ return set.indexOf(String(r.status)) >= 0; }); }
    if(p.requester){ rows = rows.filter(function(r){ return r.requester_id === p.requester; }); }
    if(p.ticket_no){ rows = rows.filter(function(r){ return String(r.ticket_no) === String(p.ticket_no); }); }
    return rows.sort(function(a,b){ return String(b.created_at).localeCompare(String(a.created_at)); }).map(strip);
  },
  todo_users: function(){
    var c = sbProps(); if(!c.url || !c.key) return [];
    var r = UrlFetchApp.fetch(c.url + '/rest/v1/users?select=id,display_name,line_user_id&order=display_name', { headers: sbHeaders(c.key), muteHttpExceptions:true });
    var arr = JSON.parse(r.getContentText() || '[]'); return Array.isArray(arr) ? arr : [];
  },
  assign: function(p){
    ensureReqCols();
    var cur = getRows(SHEETS.REQ).filter(function(r){ return r.ticket_no === p.ticket_no; })[0];
    var base = { assignee_id: p.assignee_id||'', assignee_name: p.assignee_name||'', technician1: p.assignee_name||'', updated_at: now() };
    // อาคาร-สถานที่: มอบหมาย -> ส่งให้ผู้รับผิดชอบทำงานเลย (อนุมัติทีหลังตอนงานเสร็จ)
    if(cur && String(cur.asset_category) === 'building'){
      base.status = 'in_progress';
      patchByTicket(SHEETS.REQ, p.ticket_no, base);
      logStatus(p.ticket_no, cur.status, 'in_progress', p.actor||'admin', 'มอบหมาย ' + (p.assignee_name||'') + ' → ส่งให้ผู้รับผิดชอบ');
      var task = createTodoTask({
        ticket_no: p.ticket_no,
        title: 'ซ่อม ' + (cur.machine_name || cur.machine_code || '') + ' · ' + p.ticket_no,
        description: cur.symptom || '',
        priority: 'medium',
        assignee_id: p.assignee_id || '',
        creator_line_id: cur.requester_id || '',
        due_date: cur.due_date || ''
      });
      return { ok:true, todo: task, flow:'building' };
    }
    base.status = 'pending_approval';
    patchByTicket(SHEETS.REQ, p.ticket_no, base);
    logStatus(p.ticket_no, cur?cur.status:'', 'pending_approval', p.actor||'admin', 'มอบหมาย ' + (p.assignee_name||''));
    notifyApprovalCard(p.ticket_no, false);
    return { ok:true };
  },
  // แอดมินตรวจงานเสร็จแล้ว -> ส่งให้ผู้อนุมัติ (ขั้นตอนสุดท้ายของ flow อาคาร-สถานที่)
  send_approval: function(p){
    var cur = getRows(SHEETS.REQ).filter(function(r){ return r.ticket_no === p.ticket_no; })[0];
    var prev = cur ? String(cur.status) : '';
    var isResend = (prev === 'rejected');
    var patch = { status:'pending_approval', reviewed_by: p.actor||'', reviewed_at: now(), updated_at: now() };
    if(isResend) patch.rejected_reason = '';   // ล้างเหตุผลเดิม จะได้ไม่ค้างบนใบที่แก้แล้ว
    patchByTicket(SHEETS.REQ, p.ticket_no, patch);
    logStatus(p.ticket_no, prev, 'pending_approval', p.actor||'admin',
      isResend ? 'แก้ไขตามที่ตีกลับแล้ว → ส่งอนุมัติใหม่'
               : (prev === 'done' ? 'แอดมินตรวจงานแล้ว → ส่งอนุมัติ' : 'แอดมินตรวจข้อมูลแล้ว → ส่งอนุมัติ'));
    notifyApprovalCard(p.ticket_no, isResend);
    if(isResend && cur) linePush(cur.requester_id, '🔁 ใบแจ้งซ่อม ' + p.ticket_no + ' แก้ไขแล้ว ส่งให้ผู้อนุมัติอีกครั้ง');
    return { ok:true, resend: isResend };
  },
  logs: function(p){
    return getRows(SHEETS.LOG).filter(function(r){ return r.ticket_no === p.ticket_no; })
      .sort(function(a,b){ return String(a.created_at).localeCompare(String(b.created_at)); }).map(strip);
  },
  set_status: function(p){
    var cur = getRows(SHEETS.REQ).filter(function(r){ return r.ticket_no === p.ticket_no; })[0];
    patchByTicket(SHEETS.REQ, p.ticket_no, { status: p.status, updated_at: now() });
    logStatus(p.ticket_no, cur ? cur.status : '', p.status, p.actor || '', p.note || '');
    return { ok:true };
  },
  approve: function(p){
    var cur = getRows(SHEETS.REQ).filter(function(r){ return r.ticket_no === p.ticket_no; })[0];
    if(!cur) throw 'ticket not found';
    patchByTicket(SHEETS.REQ, p.ticket_no, { status:'approved', approved_at: now(), approver_id: p.approver_line_id||'', updated_at: now() });
    logStatus(p.ticket_no, cur.status, 'approved', p.approver_line_id||'', '');
    // ส่งงานเข้า to-do เฉพาะ "ซ่อมใน" (ช่างภายใน) — ซ่อมนอกแค่เดินเรื่องเบิก
    var task = { skipped: 'ไม่สร้าง task (ซ่อมนอก หรือ อาคาร-สถานที่ที่สร้างตอนมอบหมายแล้ว)' };
    if(cur.repair_by === 'internal' && String(cur.asset_category) !== 'building'){
      task = createTodoTask({
        ticket_no: p.ticket_no,
        title: 'ซ่อม ' + (cur.machine_name || cur.vehicle_key || '') + ' · ' + p.ticket_no,
        description: cur.symptom || cur.fix_detail || '',
        priority: 'medium',
        assignee_id: cur.assignee_id || '',
        assignee_line_id: p.assignee_line_id || '',
        creator_line_id: cur.requester_id || '',
        due_date: cur.due_date || ''
      });
    }
    try{ syncHistory(p.ticket_no, true); }catch(e){}   // ลงประวัติซ่อม (เฉพาะตอนอนุมัติเท่านั้น)
    var pdfUrl='';
    try{ pdfUrl = genPdf(p.ticket_no); }catch(e){ pdfUrl = 'PDF error: ' + e; }
    linePush(cur.requester_id, '✅ ใบแจ้งซ่อม ' + p.ticket_no + ' ได้รับการอนุมัติแล้ว');
    return { ok:true, todo: task, pdf_url: pdfUrl };
  },
  gen_pdf: function(p){ return { pdf_url: genPdf(p.ticket_no) }; },
  // ดึงลายเซ็นของตัวเองมาแสดง (ส่งเป็น base64 — ไม่ต้องเปิดไฟล์เป็น public)
  signature: function(p){
    var e = getRows(SHEETS.EMP).filter(function(r){ return String(r.line_user_id) === String(p.line_user_id); })[0];
    if(!e || !e.signature_file_id) return { has:false };
    try{
      var f = DriveApp.getFileById(String(e.signature_file_id)), b = f.getBlob();
      return { has:true, mime:b.getContentType(), data:Utilities.base64Encode(b.getBytes()), updated:String(f.getLastUpdated()) };
    }catch(err){ return { has:false, error:String(err) }; }
  },
  // อัปโหลดลายเซ็นของตัวเอง -> เก็บใน Drive + ผูก file id ไว้ที่แถวพนักงาน
  upload_signature: function(p){
    var s = sh(SHEETS.EMP);
    var vals = s.getDataRange().getValues(), head = vals[0];
    var lcol = head.indexOf('line_user_id');
    if(head.indexOf('signature_file_id') < 0){
      s.getRange(1, head.length + 1).setValue('signature_file_id');
      head.push('signature_file_id');
    }
    var scol = head.indexOf('signature_file_id'), row = -1, name = '';
    for(var i=1;i<vals.length;i++){
      if(String(vals[i][lcol]) === String(p.line_user_id)){ row = i + 1; name = String(vals[i][head.indexOf('full_name')] || ''); break; }
    }
    if(row < 0) throw 'ไม่พบพนักงานคนนี้';
    var folder = getSignFolder();
    var blob = Utilities.newBlob(Utilities.base64Decode(p.data), 'image/png',
                 'sign_' + (name || p.line_user_id).replace(/[\\/:*?"<>|]/g,'') + '.png');
    // ลบไฟล์เก่าก่อน จะได้ไม่มีลายเซ็นค้างหลายอัน
    var old = String(vals[row-1][scol] || '');
    if(old){ try{ DriveApp.getFileById(old).setTrashed(true); }catch(e){} }
    var file = folder.createFile(blob);
    s.getRange(row, scol + 1).setValue(file.getId());
    return { ok:true, folder: folder.getName(), file: file.getName() };
  },
  // ทดสอบส่งการ์ด Flex (ไม่แก้ข้อมูลใดๆ) — คืนสถานะจาก LINE API มาดูว่าพลาดตรงไหน
  test_flex: function(p){
    var r = getRows(SHEETS.REQ).filter(function(x){ return x.ticket_no === p.ticket_no; })[0];
    if(!r) return { error:'ไม่พบใบ ' + p.ticket_no };
    var tk = PropertiesService.getScriptProperties().getProperty('LINE_PUSH_TOKEN');
    if(!tk) return { error:'ยังไม่ได้ตั้ง LINE_PUSH_TOKEN' };
    var to = p.to || approverIds()[0];
    if(!to) return { error:'ไม่มีผู้อนุมัติที่ active ในชีต Employees', approvers: approverIds().length };
    var bubble;
    try{ bubble = approvalBubble(r, false); }
    catch(err){ return { error:'สร้างการ์ดไม่สำเร็จ: ' + err }; }
    var res = UrlFetchApp.fetch('https://api.line.me/v2/bot/message/push', {
      method:'post', contentType:'application/json',
      headers:{ Authorization:'Bearer ' + tk },
      payload: JSON.stringify({ to:String(to), messages:[{ type:'flex', altText:'ทดสอบการ์ด ' + p.ticket_no, contents:bubble }] }),
      muteHttpExceptions:true
    });
    return { http: res.getResponseCode(), body: String(res.getContentText()).slice(0, 500), sent_to: String(to).slice(0, 8) + '…' };
  },
  // ดูค่าที่จะถูกพิมพ์ลงฟอร์ม (ไม่สร้างไฟล์ ไม่แก้ชีต) — ไว้ตรวจว่าโค้ดล่าสุดขึ้นแล้วหรือยัง
  pdf_map: function(p){ return pdfMap(p.ticket_no); },
  // อ่านผังเซลล์ของ FormTemplate (ใช้ตอนวางตำแหน่งลายเซ็น) — อ่านอย่างเดียว ไม่แก้ชีต
  dump_template: function(){
    var sh2 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('FormTemplate');
    if(!sh2) return { error:'ไม่พบแท็บ FormTemplate' };
    var vals = sh2.getDataRange().getValues(), out = [];
    for(var i=0;i<vals.length;i++) for(var j=0;j<vals[i].length;j++){
      var v = String(vals[i][j]).trim();
      if(v !== '') out.push(sh2.getRange(i+1,j+1).getA1Notation() + ' = ' + v);
    }
    return { cells: out, merges: sh2.getDataRange().getMergedRanges().map(function(m){ return m.getA1Notation(); }) };
  },
  // รายชื่อไฟล์ในโฟลเดอร์ลายเซ็น (ใช้ตรวจว่าตั้งชื่อไฟล์ตรงกับพนักงานไหม)
  list_signatures: function(){
    var id = PropertiesService.getScriptProperties().getProperty('SIGN_FOLDER_ID');
    var diag = { prop_SIGN_FOLDER_ID: id || '(ยังไม่ได้ตั้ง)',
                 all_props: PropertiesService.getScriptProperties().getKeys() };
    if(id){
      try{ var t = DriveApp.getFolderById(id); diag.by_id = 'OK: ' + t.getName(); }
      catch(e){ diag.by_id = 'ERROR: ' + e; }
    }
    var f = getSignFolder();
    if(!f) return { error:'ไม่พบโฟลเดอร์ลายเซ็น', diag: diag };
    var it = f.getFiles(), names = [];
    while(it.hasNext() && names.length < 60) names.push(it.next().getName());
    return { folder: f.getName(), files: names };
  },
  // เขียน/อัปเดตประวัติย้อนหลังให้ใบที่อนุมัติไปแล้ว (รันครั้งเดียวจากแอดมิน)
  backfill_history: function(){
    var n = 0;
    getRows(SHEETS.REQ).forEach(function(r){
      if(r.vehicle_key && r.approved_at){ try{ syncHistory(r.ticket_no, true); n++; }catch(e){} }
    });
    return { ok:true, synced:n };
  },
  reject: function(p){
    var cur = getRows(SHEETS.REQ).filter(function(r){ return r.ticket_no === p.ticket_no; })[0];
    // เก็บว่าใครเป็นคนตีกลับด้วย ไม่งั้นหน้า "ประวัติของผู้อนุมัติ" จะหาใบนี้ไม่เจอ
    patchByTicket(SHEETS.REQ, p.ticket_no, { status:'rejected', rejected_reason: p.reason||'',
      approver_id: p.approver_line_id||'', updated_at: now() });
    logStatus(p.ticket_no, cur?cur.status:'', 'rejected', p.approver_line_id||'', p.reason||'');
    if(cur) linePush(cur.requester_id, '❌ ใบแจ้งซ่อม ' + p.ticket_no + ' ไม่ได้รับการอนุมัติ' + (p.reason?(' · '+p.reason):''));
    return { ok:true };
  },
  // แก้ไข/เติมรายละเอียดใบแจ้งซ่อม (แอดมิน) — เติมข้อมูลก่อนออก PDF
  edit_request: function(p){
    var allow=['symptom','cause','cause_type','fix_detail','prevention','amount','mileage','service_interval_km','vendor','machine_code','machine_name','department','requester_name','due_date','receipt_no','contamination'];
    var patch={ updated_at: now() };
    allow.forEach(function(k){ if(p[k]!==undefined) patch[k]=p[k]; });
    patchByTicket(SHEETS.REQ, p.ticket_no, patch);
    logStatus(p.ticket_no, '', '', p.actor||'admin', 'แก้ไขข้อมูลใบแจ้งซ่อม');
    try{ syncHistory(p.ticket_no, false); }catch(e){}  // อัปเดตแถวประวัติเดิม (ถ้าอนุมัติแล้ว)
    return { ok:true };
  },
  // แนบบิล/ใบเสร็จ (รูป base64) -> เก็บใน Drive -> ต่อ url เข้า receipt_urls
  upload_receipt: function(p){
    var folder = getReceiptFolder();
    var ext = String(p.mime||'') === 'application/pdf' ? '.pdf' : '.jpg';
    var blob = Utilities.newBlob(Utilities.base64Decode(p.data), p.mime||'image/jpeg',
                 p.filename || ('receipt_'+p.ticket_no+'_'+uuid().slice(0,8)+ext));
    var file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    var url = 'https://drive.google.com/file/d/' + file.getId() + '/view';
    var cur = getRows(SHEETS.REQ).filter(function(x){ return x.ticket_no === p.ticket_no; })[0];
    var urls = (cur && cur.receipt_urls ? String(cur.receipt_urls) + ' , ' : '') + url;
    patchByTicket(SHEETS.REQ, p.ticket_no, { receipt_urls: urls, updated_at: now() });
    try{ syncHistory(p.ticket_no, false); }catch(e){}  // บิลที่แนบทีหลัง -> เข้าช่อง "ไฟล์แนบ" ของประวัติ
    return { ok:true, url:url };
  },
  // แนบไฟล์ (รูป) เข้ารายการประวัติซ่อม (MaintenanceRecords) อ้างอิงด้วย id
  hist_attach: function(p){
    var s = sh(SHEETS.HIST);
    var head = s.getRange(1,1,1,s.getLastColumn()).getValues()[0];
    if(head.indexOf('ไฟล์แนบ') < 0){ s.getRange(1, s.getLastColumn()+1).setValue('ไฟล์แนบ'); head.push('ไฟล์แนบ'); }
    var folder = getReceiptFolder();
    var hext = String(p.mime||'') === 'application/pdf' ? '.pdf' : '.jpg';
    var blob = Utilities.newBlob(Utilities.base64Decode(p.data), p.mime||'image/jpeg',
                 p.filename || ('hist_'+uuid().slice(0,8)+hext));
    var file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    var url = 'https://drive.google.com/file/d/' + file.getId() + '/view';
    var vals = s.getDataRange().getValues();
    var idcol = head.indexOf('id'), fcol = head.indexOf('ไฟล์แนบ');
    for(var i=1;i<vals.length;i++){
      if(String(vals[i][idcol]) === String(p.id)){
        var all = (vals[i][fcol] ? String(vals[i][fcol]) + ' , ' : '') + url;
        s.getRange(i+1, fcol+1).setValue(all);
        return { ok:true, url:url, all:all };
      }
    }
    return { ok:false, error:'ไม่พบรายการ' };
  },
  // ===== auth / register — นิยามใน 05_gas_auth.gs =====
  request_otp:  function(p){ return authRequestOtp(p); },
  verify_otp:   function(p){ return authVerifyOtp(p); },
  set_pin:      function(p){ return authSetPin(p); },
  login_token:  function(p){ return authLoginToken(p); },
  login_pin:    function(p){ return authLoginPin(p); },
  reg_lookup:   function(p){ return regLookup(p); },
  reg_submit:   function(p){ return regSubmit(p); },
  pending_users:function(p){ return pendingUsers(p); },
  approve_user: function(p){ return approveUser(p); },
  reject_user:  function(p){ return rejectUser(p); }
};
function strip(o){ var c={}; for(var k in o){ if(k!=='_row') c[k]=o[k]; } return c; }

// ===== INTEGRATION: ssb-maintenance <-> Job to-do (Supabase) =====
// ตั้งค่าใน Project Settings → Script Properties:
//   SUPABASE_URL          = https://xxxx.supabase.co
//   SUPABASE_SERVICE_KEY  = <service_role key>   (ห้ามใส่ในโค้ด/ฝั่ง client)
//   TODO_PROJECT_ID       = <uuid ของโปรเจค "งานซ่อม">
function sbProps(){
  var P = PropertiesService.getScriptProperties();
  return { url:P.getProperty('SUPABASE_URL'), key:P.getProperty('SUPABASE_SERVICE_KEY'), proj:P.getProperty('TODO_PROJECT_ID') };
}
function sbHeaders(key){ return { apikey:key, Authorization:'Bearer '+key }; }

// ===== LINE push แจ้งเตือน (ผ่าน OA Messaging API) — ตั้ง Script Property: LINE_PUSH_TOKEN =====
function linePush(to, text){
  var tk = PropertiesService.getScriptProperties().getProperty('LINE_PUSH_TOKEN');
  if(!tk || !to) return;
  try{
    UrlFetchApp.fetch('https://api.line.me/v2/bot/message/push', {
      method:'post', contentType:'application/json',
      headers:{ Authorization:'Bearer '+tk },
      payload: JSON.stringify({ to:String(to), messages:[{ type:'text', text:String(text).slice(0,4900) }] }),
      muteHttpExceptions:true
    });
  }catch(e){}
}
function notifyAdmins(text){
  getRows(SHEETS.EMP).filter(function(r){ return /admin/i.test(String(r.role||'')) && r.line_user_id && String(r.active)==='true'; })
    .forEach(function(a){ linePush(a.line_user_id, text); });
}
function notifyApprovers(text){
  approverIds().forEach(function(id){ linePush(id, text); });
}
function approverIds(){
  return getRows(SHEETS.EMP)
    .filter(function(r){ return /approv|exec|manager|บริหาร/i.test(String(r.role||'')) && r.line_user_id && String(r.active)==='true'; })
    .map(function(r){ return r.line_user_id; });
}

// ===== Flex message: ใบรออนุมัติ (กดอนุมัติ/ไม่อนุมัติ/ดูประวัติ ได้จากแชตเลย) =====
var LIFF_ID_APP = '2010695850-idkjafAC';
function liffUrl(q){ return 'https://liff.line.me/' + LIFF_ID_APP + (q ? ('?' + q) : ''); }
function pushFlex(to, alt, bubble){
  var tk = PropertiesService.getScriptProperties().getProperty('LINE_PUSH_TOKEN');
  if(!tk || !to) return;
  try{
    UrlFetchApp.fetch('https://api.line.me/v2/bot/message/push', {
      method:'post', contentType:'application/json',
      headers:{ Authorization:'Bearer ' + tk },
      payload: JSON.stringify({ to:String(to), messages:[{ type:'flex', altText:String(alt).slice(0,390), contents:bubble }] }),
      muteHttpExceptions:true
    });
  }catch(e){}
}
function fxRow(label, value){
  return { type:'box', layout:'baseline', spacing:'sm', contents:[
    { type:'text', text:String(label), size:'sm', color:'#8a8ca3', flex:3, weight:'bold' },
    { type:'text', text:String(value || '-'), size:'sm', color:'#23264d', flex:7, wrap:true }
  ]};
}
/** สร้างการ์ดใบแจ้งซ่อมสำหรับผู้อนุมัติ — ข้อมูลชุดเดียวกับการ์ดในแอป */
function approvalBubble(r, resend){
  var veh = r.vehicle_key ? getRows(SHEETS.VEH).filter(function(v){ return v.vehicle_key === r.vehicle_key; })[0] : null;
  var asset = r.vehicle_key
    ? [r.vehicle_key, veh && veh['ยี่ห้อ_รุ่น'], veh && veh.plate_current].filter(Boolean).join(' · ')
    : [r.machine_name, r.machine_code && ('ห้อง ' + r.machine_code)].filter(Boolean).join(' · ');
  var re = (String(r.request_type) === 'reimburse');
  var body = [];
  body.push(fxRow('รถ/เครื่อง', asset));
  // ผู้ดูแลรถประจำคัน — บางครั้งแอดมินแจ้งแทน ผู้อนุมัติจะได้รู้ว่ารถของใคร
  if(veh){
    var own = [veh['ผู้รับผิดชอบ'], veh['แผนก'], veh['สถานที่']].filter(Boolean).join(' · ');
    if(own) body.push(fxRow('ผู้ดูแลรถ', own));
  }
  if(r.mileage) body.push(fxRow('เลขไมล์', money(r.mileage) + ' กม.'));
  if(r.service_interval_km) body.push(fxRow('รอบเช็ค', money(r.service_interval_km) + ' กม.'));
  body.push(fxRow('รายการ', (KIND_TH[r.request_kind] || r.request_kind || '') + ' ' + (r.symptom || '')));
  if(r.fix_detail) body.push(fxRow('การแก้ไข', r.fix_detail));
  body.push(fxRow(r.repair_by === 'internal' ? 'ช่าง' : 'ศูนย์/อู่',
    r.repair_by === 'internal' ? (r.assignee_name || 'ยังไม่มอบหมาย') : (r.vendor || 'ไม่ได้ระบุ')));
  body.push(fxRow('จำนวนเงิน', baht(r.amount) + ' บาท'));
  body.push(fxRow('ผู้แจ้ง', r.requester_name || ''));

  // LINE ไม่ยอมรับ color:null -> ต้องไม่ใส่คีย์ color เลยเมื่อไม่ระบุสี
  var btn = function(label, style, color, action){
    var o = { type:'button', style:style, height:'sm', action:action };
    if(color) o.color = color;
    return o;
  };
  // postback = ทำงานจบในแชตเลย ไม่ต้องเปิดแอป
  var pb = function(label, act){
    return { type:'postback', label:label, data:pbData(act, r.ticket_no), displayText:(act==='approve'?'✓ อนุมัติ ':'✕ ไม่อนุมัติ ') + r.ticket_no };
  };
  var uri = function(label, url){ return { type:'uri', label:label, uri:url }; };
  return {
    type:'bubble',
    header:{ type:'box', layout:'vertical', backgroundColor:'#33348f', paddingAll:'14px', contents:[
      { type:'text', text:(resend ? '🔁 แก้ไขแล้ว ส่งอนุมัติใหม่' : '📋 ใบแจ้งซ่อมรออนุมัติ'), size:'xs', color:'#c9c9ee' },
      { type:'text', text:String(r.ticket_no), size:'lg', weight:'bold', color:'#ffffff' },
      { type:'text', text:(re ? '💸 ทำแล้วมาเบิก — จ่ายไปแล้ว รออนุมัติเบิกคืน' : '📝 ขออนุมัติก่อนทำ — อนุมัติแล้วจึงเริ่มซ่อม'),
        size:'xxs', color:'#ffd591', wrap:true, margin:'sm' }
    ]},
    body:{ type:'box', layout:'vertical', spacing:'sm', paddingAll:'14px', contents:body },
    footer:{ type:'box', layout:'vertical', spacing:'sm', paddingAll:'12px', contents:[
      { type:'box', layout:'horizontal', spacing:'sm', contents:[
        btn('✓ อนุมัติ', 'primary', '#22a06b', pb('✓ อนุมัติ','approve')),
        btn('✕ ไม่อนุมัติ', 'primary', '#e24b4a', pb('✕ ไม่อนุมัติ','reject'))
      ]},
      btn('🕘 ดูรายละเอียด / ประวัติ', 'secondary', null, uri('ดูรายละเอียด', liffUrl('t=' + r.ticket_no)))
    ]}
  };
}
var KIND_TH = { repair:'ซ่อม', replace_part:'เปลี่ยนอะไหล่', inspect:'ตรวจเช็คระยะ', tire:'เปลี่ยนยาง', install:'ติดตั้ง', other:'อื่นๆ' };

// ===== LINE webhook: กดปุ่มในแชตแล้วอนุมัติได้เลย =====
/**
 * GAS อ่าน header ไม่ได้ จึงตรวจ X-Line-Signature ไม่ได้
 * -> ใส่ลายเซ็นของเราเองไปกับ postback data แล้วตรวจตอนรับกลับ
 *    (ผู้ที่ไม่เคยได้รับการ์ดจะเดา sig ไม่ได้) + ตรวจสิทธิ์ผู้ใช้จาก userId ซ้ำอีกชั้น
 */
function pbSecret(){
  var sp = PropertiesService.getScriptProperties();
  var k = sp.getProperty('PB_SECRET');
  if(!k){ k = Utilities.getUuid() + Utilities.getUuid(); sp.setProperty('PB_SECRET', k); }
  return k;
}
function pbSig(act, ticket){
  var raw = Utilities.computeHmacSha256Signature(act + '|' + ticket, pbSecret());
  return Utilities.base64EncodeWebSafe(raw).slice(0, 16);
}
function pbData(act, ticket){ return 'act=' + act + '&t=' + ticket + '&s=' + pbSig(act, ticket); }
function parseQS(s){
  var o = {};
  String(s || '').split('&').forEach(function(kv){
    var i = kv.indexOf('='); if(i < 0) return;
    o[decodeURIComponent(kv.slice(0, i))] = decodeURIComponent(kv.slice(i + 1));
  });
  return o;
}
function lineReply(token, text){
  var tk = PropertiesService.getScriptProperties().getProperty('LINE_PUSH_TOKEN');
  if(!tk || !token) return;
  try{
    UrlFetchApp.fetch('https://api.line.me/v2/bot/message/reply', {
      method:'post', contentType:'application/json',
      headers:{ Authorization:'Bearer ' + tk },
      payload: JSON.stringify({ replyToken: token, messages:[{ type:'text', text:String(text).slice(0, 4900) }] }),
      muteHttpExceptions:true
    });
  }catch(e){}
}
function handleLineEvents(events){
  (events || []).forEach(function(ev){
    if(!ev || ev.type !== 'postback') return;
    var d = parseQS(ev.postback && ev.postback.data);
    if(!d.act || !d.t) return;
    var uid = ev.source && ev.source.userId;
    var reply = ev.replyToken;

    if(d.s !== pbSig(d.act, d.t)){ lineReply(reply, '❌ ลิงก์ไม่ถูกต้อง'); return; }

    var emp = getRows(SHEETS.EMP).filter(function(r){ return String(r.line_user_id) === String(uid); })[0];
    var isApprover = emp && String(emp.active) === 'true' && /approv|exec|manager|บริหาร/i.test(String(emp.role || ''));
    if(!isApprover){ lineReply(reply, '⛔ บัญชีนี้ไม่มีสิทธิ์อนุมัติ'); return; }

    var cur = getRows(SHEETS.REQ).filter(function(r){ return r.ticket_no === d.t; })[0];
    if(!cur){ lineReply(reply, '❌ ไม่พบใบ ' + d.t); return; }
    if(String(cur.status) !== 'pending_approval'){
      lineReply(reply, 'ℹ️ ใบ ' + d.t + ' ถูกดำเนินการไปแล้ว (' + (STATUS_TH[cur.status] || cur.status) + ')');
      return;
    }

    if(d.act === 'approve'){
      var res = API.approve({ ticket_no: d.t, approver_line_id: uid });
      lineReply(reply, '✅ อนุมัติ ' + d.t + ' เรียบร้อย\nโดย ' + (emp.full_name || '')
        + (res && /^https?:/.test(String(res.pdf_url || '')) ? ('\n\n📄 ใบงาน: ' + res.pdf_url) : ''));
    } else if(d.act === 'reject'){
      API.reject({ ticket_no: d.t, reason: 'ตีกลับจากแชต', approver_line_id: uid });
      lineReply(reply, '↩️ ตีกลับ ' + d.t + ' แล้ว\nโดย ' + (emp.full_name || '')
        + '\n\nระบุเหตุผลเพิ่มได้ที่: ' + liffUrl('t=' + d.t));
    }
  });
}
var STATUS_TH = { submitted:'รอแอดมินตรวจ', pending_approval:'รออนุมัติ', approved:'อนุมัติแล้ว',
  in_progress:'กำลังทำ', done:'เสร็จ', sent_accounting:'ส่งบัญชี', closed:'ปิดงาน', rejected:'ไม่อนุมัติ' };
/** แจ้งผู้อนุมัติด้วยการ์ด (ถ้าสร้างการ์ดไม่ได้ ตกกลับไปเป็นข้อความธรรมดา) */
function notifyApprovalCard(ticketNo, resend){
  var r = getRows(SHEETS.REQ).filter(function(x){ return x.ticket_no === ticketNo; })[0];
  var ids = approverIds();
  if(!r || !ids.length){ notifyApprovers('📋 มีใบแจ้งซ่อมรออนุมัติ: ' + ticketNo); return; }
  try{
    var b = approvalBubble(r, resend);
    var alt = (resend ? '🔁 แก้ไขแล้ว ส่งอนุมัติใหม่: ' : '📋 ใบแจ้งซ่อมรออนุมัติ: ') + ticketNo;
    ids.forEach(function(id){ pushFlex(id, alt, b); });
  }catch(e){
    notifyApprovers('📋 มีใบแจ้งซ่อมรออนุมัติ: ' + ticketNo);
  }
}

function createTodoTask(req){
  var c = sbProps();
  if(!c.url || !c.key || !c.proj) return { skipped:'no supabase config' };
  var assignee = req.assignee_id || null;
  if(!assignee && req.assignee_line_id){
    var r = UrlFetchApp.fetch(c.url + '/rest/v1/users?select=id&line_user_id=eq.' + encodeURIComponent(req.assignee_line_id),
      { headers: sbHeaders(c.key), muteHttpExceptions:true });
    var arr = JSON.parse(r.getContentText() || '[]');
    if(arr[0]) assignee = arr[0].id;
  }
  // ผู้สร้างงาน = ผู้แจ้ง (map line_user_id -> users.id ของ to-do)
  var creator = req.creator_id || null;
  if(!creator && req.creator_line_id){
    var rc = UrlFetchApp.fetch(c.url + '/rest/v1/users?select=id&line_user_id=eq.' + encodeURIComponent(req.creator_line_id),
      { headers: sbHeaders(c.key), muteHttpExceptions:true });
    var ac = JSON.parse(rc.getContentText() || '[]');
    if(ac[0]) creator = ac[0].id;
  }
  var body = { project_id:c.proj, title:req.title, description:req.description||'',
    priority:req.priority||'medium', assignee_id:assignee, created_by:creator, due_date:req.due_date||null,
    external_ref:'ssb:'+req.ticket_no };
  var res = UrlFetchApp.fetch(c.url + '/rest/v1/tasks',
    { method:'post', contentType:'application/json',
      headers: { apikey:c.key, Authorization:'Bearer '+c.key, Prefer:'return=representation' },
      payload: JSON.stringify(body), muteHttpExceptions:true });
  return JSON.parse(res.getContentText() || '{}');
}

// ตั้ง time-driven trigger เรียกทุก 5-10 นาที: task เสร็จใน to-do -> ปิดใบแจ้งซ่อม
function pollTodoDone(){
  var c = sbProps();
  if(!c.url || !c.key) return;
  var P = PropertiesService.getScriptProperties();
  var last = P.getProperty('TODO_LAST_POLL') || '1970-01-01T00:00:00Z';
  var q = c.url + '/rest/v1/tasks?select=external_ref,updated_at,status&status=eq.done&external_ref=like.ssb:*&updated_at=gt.' + encodeURIComponent(last);
  var r = UrlFetchApp.fetch(q, { headers: sbHeaders(c.key), muteHttpExceptions:true });
  var tasks = JSON.parse(r.getContentText() || '[]');
  if(!tasks.length) return;
  var maxTs = last;
  tasks.forEach(function(t){
    var ticket = String(t.external_ref).replace(/^ssb:/, '');
    patchByTicket(SHEETS.REQ, ticket, { status:'done', repair_finish: now(), updated_at: now() });
    logStatus(ticket, '', 'done', 'todo-sync', 'งานเสร็จจากระบบ to-do');
    notifyAdmins('🔧 งานเสร็จแล้ว รอแอดมินตรวจ: ' + ticket);
    if(t.updated_at > maxTs) maxTs = t.updated_at;
  });
  P.setProperty('TODO_LAST_POLL', maxTs);
}

// ทดสอบเชื่อม to-do (รันหลังใส่ Script Properties แล้ว) — สร้าง task ทดสอบ 1 อัน
function testTodo(){
  var r = getRows(SHEETS.REQ)[0];
  if(!r){ Logger.log('ไม่มีใบแจ้งซ่อม'); return; }
  var res = createTodoTask({ ticket_no:r.ticket_no, title:'[ทดสอบ] ซ่อม '+r.ticket_no, description:'ทดสอบเชื่อม ssb-maintenance', priority:'medium', assignee_line_id:'', due_date:'' });
  Logger.log(JSON.stringify(res));
  return res;
}
