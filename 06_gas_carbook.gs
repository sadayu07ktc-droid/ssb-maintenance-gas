// ============================================================
//  จองรถบริษัท (Car booking)
//  พนักงาน -> หัวหน้าแผนก -> ผู้บริหาร -> HR เลือกรถ/คนขับ -> แจ้งผู้ขอ -> บันทึกไมล์ออก/กลับ
//  หัวหน้าแผนก = ติ๊ก dept_approver ในชีต Employees (จับคู่ด้วย department)
//  ผู้บริหาร = role ผู้อนุมัติเดิม · HR = role admin
// ============================================================
var CAR_SHEET = 'CarBookings';
var CAR_HEAD = ['booking_no','status','requester_id','requester_name','department','phone',
  'destination','origin','go_at','back_at','pax','need_driver','car_type','purpose','passengers','note',
  'head_id','head_name','head_at','exec_id','exec_name','exec_at',
  'reject_step','reject_by','reject_reason','reject_note','rejected_at',
  'vehicle_key','plate','car_model','driver_id','driver_name','driver_phone','meet_point','hr_id','hr_name','hr_at','hr_note',
  'out_at','out_mileage','in_at','in_mileage','in_condition','distance_km',
  'cancelled_at','created_at','updated_at','substitute'];   // substitute = ผู้อนุมัติสั่งให้หาคนไปแทน -> HR ต้องเลือกคนไปแทน
var CAR_STEP_TH = { pending_head:'รอหัวหน้าแผนก', pending_exec:'รอผู้อนุมัติ', pending_hr:'รอ HR จัดรถ',
  assigned:'พร้อมเดินทาง', in_use:'กำลังใช้รถ', returned:'คืนรถแล้ว', rejected:'ไม่อนุมัติ', cancelled:'ยกเลิก' };
var CAR_TYPE_TH = { sedan:'เก๋ง / SUV', pickup:'กระบะ', van:'รถตู้', '':'ไม่ระบุ' };
var CAR_BUSY = ['assigned','in_use'];   // สถานะที่ถือว่ารถถูกใช้ช่วงนั้นแล้ว

// ---------- ชีต ----------
function carSheet_(){
  var s = sh(CAR_SHEET);
  if(!s){
    s = ss().insertSheet(CAR_SHEET);
    s.getRange(1,1,1,CAR_HEAD.length).setValues([CAR_HEAD]).setFontWeight('bold');
    // เก็บทุกช่องเป็นข้อความ ไม่งั้นชีตแปลง "08:30" เป็นเวลา แล้วอ่านกลับได้ 1899-12-30
    s.getRange(1,1,s.getMaxRows(),CAR_HEAD.length).setNumberFormat('@');
    s.setFrozenRows(1);
  } else {
    var head = s.getRange(1,1,1,Math.max(1,s.getLastColumn())).getValues()[0];
    CAR_HEAD.forEach(function(h){ if(head.indexOf(h) < 0){ s.getRange(1, s.getLastColumn()+1).setValue(h); } });
  }
  return s;
}
// คอลัมน์ในชีต Vehicles ที่ HR ใช้ตั้งค่า (เพิ่มเฉพาะหัวคอลัมน์ ไม่เขียนแถวว่าง)
function carVehCols_(){
  ['car_bookable','car_type','car_seats','last_mileage','last_mileage_at'].forEach(function(c){ ensureCol_(SHEETS.VEH, c); });
}
function carRows_(){ carSheet_(); return getRows(CAR_SHEET); }
function carGet_(no){ return carRows_().filter(function(r){ return String(r.booking_no) === String(no); })[0] || null; }
/**
 * เขียนเฉพาะช่องที่เปลี่ยน · ห้าม setNumberFormat ในชีตนี้
 * ชีตถูกแปลงเป็น "ตาราง" (คอลัมน์มีประเภท) แล้ว setNumberFormat จะพังตอน Sheets flush ท้ายสุด
 * ซึ่ง try/catch ดักไม่ได้ -> บันทึกได้บางช่อง + ตอบกลับเป็นหน้า HTML error (CAR-202609-001, 30 ก.ย.)
 */
function carPatch_(no, patch){
  var s = carSheet_();
  var vals = s.getDataRange().getValues(), head = vals[0], k = head.indexOf('booking_no');
  patch.updated_at = now();
  for(var i=1;i<vals.length;i++){
    if(String(vals[i][k]) !== String(no)) continue;
    head.forEach(function(h,c){ if(h && patch[h] !== undefined) s.getRange(i+1, c+1).setValue(patch[h]); });
    return true;
  }
  return false;
}
// กันสองคนกด (หรือแอปส่งซ้ำ) พร้อมกันจนเขียนทับกัน / ได้เลขใบซ้ำ
function carLocked_(fn){
  var lk = LockService.getScriptLock();
  if(!lk.tryLock(20000)) throw 'ระบบกำลังบันทึกรายการอื่น ลองใหม่อีกครั้ง';
  try{ return fn(); } finally { lk.releaseLock(); }
}
function carNo_(){
  var prefix = 'CAR-' + Utilities.formatDate(new Date(), TZ, 'yyyyMM') + '-', max = 0;
  carRows_().forEach(function(r){
    var t = String(r.booking_no||''); if(t.indexOf(prefix) !== 0) return;
    var n = parseInt(t.slice(prefix.length), 10); if(!isNaN(n) && n > max) max = n;
  });
  return prefix + ('000' + (max + 1)).slice(-3);
}

// ---------- คน / สิทธิ์ ----------
function carEmp_(uid){
  if(!/^U[0-9a-f]{32}$/i.test(String(uid||''))) return null;
  return getRows(SHEETS.EMP).filter(function(r){ return String(r.line_user_id) === String(uid) && String(r.active) === 'true'; })[0] || null;
}
function carIsHead_(e){ return !!e && (e.dept_approver === true || String(e.dept_approver).toUpperCase() === 'TRUE'); }
function carDeptHeads_(dept){
  var d = String(dept||'').trim(); if(!d) return [];
  return getRows(SHEETS.EMP).filter(function(r){
    return String(r.active) === 'true' && carIsHead_(r) && String(r.department||'').trim() === d && r.line_user_id;
  });
}
// ผู้ใช้คนนี้ตัดสินใบนี้ได้ในขั้นไหน ('head' | 'exec' | '')
function carStepFor_(b, e){
  if(!b || !e) return '';
  if(b.status === 'pending_head' && carIsHead_(e) && String(e.department||'').trim() === String(b.department||'').trim()) return 'head';
  if(b.status === 'pending_exec' && /approv|exec|manager|บริหาร/i.test(String(e.role||''))) return 'exec';
  return '';
}
function carIsHR_(e){ return !!e && /admin/i.test(String(e.role||'')); }
function carIsExec_(e){ return !!e && /approv|exec|manager|บริหาร/i.test(String(e.role||'')); }

// ---------- รถ ----------
function carFleet_(){
  carVehCols_();
  return getRows(SHEETS.VEH).filter(function(v){ return v.vehicle_key; }).map(function(v){
    return { vehicle_key:v.vehicle_key, plate:String(v.plate_current||''), model:String(v['ยี่ห้อ_รุ่น']||''),
      owner:String(v['ผู้รับผิดชอบ']||''), dept:String(v['แผนก']||''),
      bookable:(v.car_bookable === true || String(v.car_bookable).toUpperCase() === 'TRUE'),
      car_type:String(v.car_type||''), seats:Number(v.car_seats||0)||'', last_mileage:v.last_mileage||'', _row:v._row };
  });
}
// ใบที่จองรถคันนี้ทับช่วงเวลาที่ถาม (ใบที่กำลังใช้อยู่แต่ยังไม่คืน ถือว่ายังไม่ว่าง)
function carClashes_(vk, goAt, backAt, exceptNo){
  return carRows_().filter(function(r){
    if(String(r.vehicle_key) !== String(vk) || String(r.booking_no) === String(exceptNo||'')) return false;
    if(CAR_BUSY.indexOf(String(r.status)) < 0) return false;
    var s = String(r.go_at), e = String(r.back_at || r.go_at);
    if(String(r.status) === 'in_use' && !r.in_at) e = '9999';
    return s < String(backAt || goAt) && String(goAt) < e;
  });
}

// ---------- แจ้งเตือน ----------
function carSteps_(status){
  var labels = ['ขอ','หัวหน้า','ผู้อนุมัติ','HR','เดินทาง'];
  var at = { pending_head:1, pending_exec:2, pending_hr:3, assigned:4, in_use:4, returned:4 }[status];
  if(at === undefined) at = 0;
  var bars = labels.map(function(_, i){
    return { type:'box', layout:'vertical', flex:1, height:'4px', cornerRadius:'2px',
      backgroundColor:(i < at ? '#ffd591' : (i === at ? '#ffffff' : '#ffffff40')), contents:[] };
  });
  var lb = labels.map(function(s, i){
    return { type:'text', text:s, size:'xxs', flex:1, align:'center', color:(i === at ? '#ffffff' : '#ffffff99'), weight:(i === at ? 'bold' : 'regular') };
  });
  return [{ type:'box', layout:'horizontal', spacing:'xs', margin:'md', contents:bars },
          { type:'box', layout:'horizontal', margin:'xs', contents:lb }];
}
function carWhen_(b){
  var g = String(b.go_at||''), k = String(b.back_at||'');
  var d = function(x){ var p = x.slice(0,10).split('-'); return p.length === 3 ? (Number(p[2]) + '/' + Number(p[1])) : x; };
  if(!g) return '-';
  if(k && k.slice(0,10) !== g.slice(0,10)) return d(g) + ' ' + g.slice(11,16) + ' – ' + d(k) + ' ' + k.slice(11,16);
  return d(g) + ' · ' + g.slice(11,16) + (k ? ('–' + k.slice(11,16)) : '');
}
function carHours_(b){
  var p = function(x){ var m = String(x||'').match(/^(\d{4})-(\d{2})-(\d{2}) (\d{2}):(\d{2})/); return m ? new Date(+m[1], m[2]-1, +m[3], +m[4], +m[5]) : null; };
  var a = p(b.go_at), z = p(b.back_at); if(!a || !z || z <= a) return '';
  var h = (z - a) / 3600000; return (Math.round(h * 10) / 10) + '';
}
/** การ์ดใบจอง · mode: head/exec = มีปุ่มอนุมัติในแชต · hr = ปุ่มเปิดจัดรถ · ok = พร้อมเดินทาง · rej = ไม่อนุมัติ · info = แจ้งเฉยๆ */
function carBubble_(b, mode, kicker){
  var color = { head:['#0e7490','#1e3a8a'], exec:['#0e7490','#1e3a8a'], hr:['#334155','#0f172a'],
                ok:['#059669','#065f46'], rej:['#b91c1c','#7f1d1d'], info:['#475569','#1e293b'], sub:['#b45309','#7c2d12'] }[mode] || ['#475569','#1e293b'];
  var note = { head:'รอหัวหน้าแผนก', exec:'รอผู้อนุมัติ', hr:'รอ HR จัดรถ', ok:'พร้อมเดินทาง', rej:'ไม่อนุมัติ', info:'จองรถบริษัท', sub:'หาคนไปแทน' }[mode];
  var title = (mode === 'ok' && b.plate) ? (b.plate + (b.car_model ? (' · ' + b.car_model) : '')) : String(b.booking_no);
  var head = fxHead(kicker || 'ใบจองรถบริษัท', title, carWhen_(b), color[0], color[1], note);
  if(mode !== 'rej') carSteps_(String(b.status)).forEach(function(x){ head.contents.push(x); });

  var body = [];
  if(mode !== 'rej' && mode !== 'ok'){
    var hrs = carHours_(b);
    body.push({ type:'box', layout:'horizontal', spacing:'sm', contents:[
      fxTile(String(b.pax || '-'), 'คน', '#0e7490'),
      fxTile(hrs || '-', 'ชั่วโมง', '#0e7490'),
      fxTile(CAR_TYPE_TH[String(b.car_type||'')] || 'ไม่ระบุ', (String(b.need_driver) === 'true' ? 'ขอคนขับ' : 'ขับเอง'), '#33348f')
    ]});
  }
  var rows = [];
  if(mode === 'ok'){
    rows.push(fxRow('ใบจอง', b.booking_no));
    if(String(b.need_driver) === 'true' || b.driver_name) rows.push(fxRow('คนขับ', (b.driver_name || '-') + (b.driver_phone ? (' · ' + b.driver_phone) : '')));
    if(b.meet_point) rows.push(fxRow('จุดนัด', b.meet_point));
  }
  if(mode === 'rej'){
    rows.push(fxRow('เหตุผล', b.reject_reason || '-'));
    if(b.reject_note) rows.push(fxRow('หมายเหตุ', b.reject_note));
    rows.push(fxRow('โดย', (b.reject_by || '-') + (b.reject_step === 'head' ? ' (หัวหน้าแผนก)' : ' (ผู้อนุมัติ)')));
  }
  rows.push(fxRow('ปลายทาง', b.destination));
  if(mode !== 'ok' && mode !== 'rej') rows.push(fxRow('วัตถุประสงค์', b.purpose));
  if(mode !== 'ok') rows.push(fxRow('ผู้ขอ', (b.requester_name || '-') + (b.department ? (' · ' + b.department) : '')));
  if((mode === 'exec' || mode === 'hr') && b.head_name) rows.push(fxRow('หัวหน้า', '✓ ' + b.head_name));
  if(mode === 'hr' && b.exec_name) rows.push(fxRow('ผู้อนุมัติ', '✓ ' + b.exec_name));
  if((mode === 'hr' || mode === 'sub' || mode === 'exec') && String(b.substitute || '')) rows.push(fxRow('หาคนแทน', String(b.substitute).replace(/^หาคนไปแทน · /, 'สั่งโดย ')));
  if(b.hr_note && mode === 'ok') rows.push(fxRow('หมายเหตุ', b.hr_note));
  body.push({ type:'box', layout:'vertical', spacing:'sm', margin:'lg', contents:rows });
  if(mode === 'ok') body.push({ type:'text', text:'กดบันทึกเลขไมล์ตอนออกรถ และตอนคืนรถ', size:'xxs', color:'#8a8ca3', margin:'lg', wrap:true });

  var btn = function(label, style, color, action){ var o = { type:'button', style:style, height:'sm', action:action }; if(color) o.color = color; return o; };
  var uri = function(label){ return { type:'uri', label:label, uri:liffUrl('car=' + b.booking_no) }; };
  var foot = [];
  if(mode === 'head' || mode === 'exec'){
    foot.push({ type:'box', layout:'horizontal', spacing:'sm', contents:[
      btn('', 'primary', '#22a06b', { type:'postback', label:'✓ อนุมัติ', data:pbData('car_ok', b.booking_no), displayText:'✓ อนุมัติ ' + b.booking_no }),
      btn('', 'primary', '#e24b4a', { type:'postback', label:'✕ ไม่อนุมัติ', data:pbData('car_no', b.booking_no), displayText:'✕ ไม่อนุมัติ ' + b.booking_no })
    ]});
    foot.push(btn('', 'secondary', null, uri('ดูรายละเอียด')));
  } else if(mode === 'hr'){
    foot.push(btn('', 'primary', '#0e7490', uri('เลือกรถ + คนขับ')));
  } else if(mode === 'ok'){
    foot.push(btn('', 'primary', '#0e7490', uri('บันทึกเลขไมล์')));
  } else if(mode === 'rej'){
    foot.push(btn('', 'secondary', null, { type:'uri', label:'จองใหม่', uri:liffUrl('go=carbook') }));
  } else {
    foot.push(btn('', 'secondary', null, uri('ดูใบจอง')));
  }
  return { type:'bubble', header:head,
    body:{ type:'box', layout:'vertical', paddingAll:'14px', contents:body },
    footer:{ type:'box', layout:'vertical', spacing:'sm', paddingAll:'12px', contents:foot } };
}
function carPush_(ids, b, mode, kicker, alt){
  var bubble = null;
  try{ bubble = carBubble_(b, mode, kicker); }catch(e){}
  (ids || []).forEach(function(id){
    if(!id) return;
    if(bubble) pushFlex(id, alt, bubble); else linePush(id, alt + '\n' + liffUrl('car=' + b.booking_no));
  });
}
// ส่งต่อไปขั้นถัดไปตามสถานะปัจจุบัน
function carRouteNotify_(no){
  var b = carGet_(no); if(!b) return;
  if(b.status === 'pending_head'){
    carPush_(carDeptHeads_(b.department).map(function(e){ return e.line_user_id; }), b, 'head', 'ลูกทีมขอใช้รถบริษัท', '🚗 ขอใช้รถ รอหัวหน้าอนุมัติ: ' + b.booking_no);
  } else if(b.status === 'pending_exec'){
    carPush_(approverIds(), b, 'exec', 'ขอใช้รถบริษัท · รอผู้อนุมัติ', '🚗 ขอใช้รถ รอผู้อนุมัติ: ' + b.booking_no);
  } else if(b.status === 'pending_hr'){
    if(String(b.substitute || '')) carPush_(adminIds(), b, 'hr', 'ให้หาคนไปแทน · HR เลือกคนไปแทน + รถ', '👥 หาคนไปแทน + จัดรถ ใบ ' + b.booking_no);
    else carPush_(adminIds(), b, 'hr', 'อนุมัติแล้ว · รอ HR จัดรถและคนขับ', '🚗 จัดรถให้ใบ ' + b.booking_no);
  }
}

// ---------- ขั้นตอนหลัก (ใช้ร่วมกันทั้งเว็บและไลน์) ----------
function carDecide_(no, uid, decision, reason, note){ return carLocked_(function(){ return carDecideNow_(no, uid, decision, reason, note); }); }
function carDecideNow_(no, uid, decision, reason, note){
  var e = carEmp_(uid); denyIf(!e, 'ไม่พบบัญชีพนักงาน');
  var b = carGet_(no); if(!b) throw 'ไม่พบใบจอง ' + no;
  // คนเดิมกดซ้ำ / แอปส่งซ้ำเพราะรอบแรกตอบกลับไม่ครบ -> ตอบว่าทำไปแล้ว ไม่ใช่ "ไม่มีสิทธิ์"
  // (เคยเกิดจริง: อนุมัติสำเร็จแล้วแต่ตอบกลับเป็น HTML แอปเลย retry แล้วขึ้น error ทั้งที่ผ่านไปแล้ว)
  var mine = function(id){ return String(id || '') === String(uid); };
  if(decision === 'approve' && b.status !== 'rejected' && b.status !== 'cancelled'
     && ((mine(b.head_id) && b.status !== 'pending_head') || (mine(b.exec_id) && b.status !== 'pending_exec'))){
    return { ok:true, already:true, status:b.status, by:String(e.full_name || '') };
  }
  if(decision !== 'approve' && reason === CAR_RS.sub && String(b.substitute || '') && (mine(b.head_id) || mine(b.exec_id))){
    return { ok:true, already:true, status:b.status, by:String(e.full_name || ''), substitute:true };
  }
  if(decision !== 'approve' && b.status === 'rejected' && String(b.reject_by) === String(e.full_name || '')){
    return { ok:true, already:true, status:b.status, by:String(e.full_name || '') };
  }
  var step = carStepFor_(b, e);
  if(!step){
    var th = CAR_STEP_TH[b.status] || b.status;
    if(mine(b.head_id) || mine(b.exec_id)) throw 'คุณอนุมัติใบนี้ไปแล้ว · ตอนนี้ ' + th;
    throw (/^pending_/.test(String(b.status)) ? ('ใบนี้อยู่ขั้น "' + th + '" · บัญชีนี้ไม่มีสิทธิ์ในขั้นนี้') : ('ใบนี้ดำเนินการไปแล้ว (' + th + ')'));
  }
  var who = String(e.full_name || '');
  if(decision === 'approve'){
    var patch = {}, both = false;
    if(step === 'head'){
      patch = { status:'pending_exec', head_id:uid, head_name:who, head_at:now() };
      // หัวหน้าแผนกที่เป็นผู้บริหารด้วย (เช่น หัวหน้าฝ่ายบริหาร) กดครั้งเดียวผ่านทั้งสองขั้น ไม่ต้องกดซ้ำ
      if(carIsExec_(e)){ patch.status = 'pending_hr'; patch.exec_id = uid; patch.exec_name = who; patch.exec_at = patch.head_at; both = true; }
    }
    else { patch = { status:'pending_hr', exec_id:uid, exec_name:who, exec_at:now() }; }
    carPatch_(no, patch);
    try{ carRouteNotify_(no); }catch(err){}   // แจ้งเตือนพลาด (เช่น โควตาไลน์เต็ม) ต้องไม่ทำให้การบันทึกล้ม
    return { ok:true, step:(both ? 'head+exec' : step), status:patch.status, by:who };
  }
  denyIf(!reason, 'ต้องเลือกเหตุผลที่ไม่อนุมัติ');
  // "หาคนไปแทน" ไม่ใช่การปฏิเสธ: ให้ส่งคนอื่นไปแทน -> ใบเดินต่อเหมือนอนุมัติ แล้ว HR เลือกคนไปแทน + รถ ในใบเดิม
  if(reason === CAR_RS.sub){
    var sp, tag = 'หาคนไปแทน · ' + who + (step === 'head' ? ' (หัวหน้าแผนก)' : ' (ผู้อนุมัติ)') + (note ? (' · ' + note) : '');
    if(step === 'head'){
      sp = { status:'pending_exec', head_id:uid, head_name:who, head_at:now() };
      if(carIsExec_(e)){ sp.status = 'pending_hr'; sp.exec_id = uid; sp.exec_name = who; sp.exec_at = sp.head_at; }
    }
    else sp = { status:'pending_hr', exec_id:uid, exec_name:who, exec_at:now() };
    sp.substitute = tag;
    carPatch_(no, sp);
    var sb = carGet_(no);
    try{ carRouteNotify_(no); }catch(err){}
    var nx = sp.status === 'pending_hr' ? 'HR จัดคนไปแทน' : 'ผู้อนุมัติ';
    // ผู้ขอคือคนที่กดเอง (เช่น ผู้อนุมัติจองให้ตัวเอง) ได้ข้อความตอบกลับแล้ว ไม่ต้องส่งการ์ดซ้ำ
    if(String(sb.requester_id) !== String(uid)) try{ carPush_([sb.requester_id], sb, 'sub', 'ให้หาคนไปแทน · ส่งต่อ ' + nx, '👥 ใบ ' + no + ' ให้หาคนไปแทน · ส่งต่อ ' + nx); }catch(err){}
    return { ok:true, step:step, status:sp.status, by:who, substitute:true };
  }
  carPatch_(no, { status:'rejected', reject_step:step, reject_by:who, reject_reason:reason, reject_note:note||'', rejected_at:now() });
  var nb = carGet_(no);
  try{ carPush_([nb.requester_id], nb, 'rej', 'ใบจองรถไม่ได้รับอนุมัติ', '✕ ใบจองรถไม่ได้รับอนุมัติ: ' + no); }catch(err){}
  return { ok:true, step:step, status:'rejected', by:who };
}

// ---------- API ----------
Object.assign(API, {
  car_setup: function(){ carSheet_(); carVehCols_(); return { sheet:CAR_SHEET, cols:CAR_HEAD.length }; },
  // ตรวจใบล่าสุด: ส่งไปหาใคร (นับจำนวน ไม่คืน UID) + ให้ LINE ตรวจการ์ดว่าถูกรูปแบบไหม
  car_diag: function(){
    var rows = carRows_(); if(!rows.length) return { bookings:0 };
    var b = rows[rows.length - 1];
    var mode = { pending_head:'head', pending_exec:'exec', pending_hr:'hr', assigned:'ok', rejected:'rej' }[b.status] || 'info';
    // ชื่อผู้รับการ์ดของขั้นปัจจุบัน (ชื่อเท่านั้น ไม่คืน UID)
    var emps = getRows(SHEETS.EMP), nameOf = function(id){ var x = emps.filter(function(r){ return String(r.line_user_id) === String(id); })[0]; return x ? x.full_name : '?'; };
    var ids = b.status === 'pending_head' ? carDeptHeads_(b.department).map(function(h){ return h.line_user_id; })
      : (b.status === 'pending_exec' ? approverIds() : (b.status === 'pending_hr' ? adminIds() : []));
    var to = ids.map(nameOf);
    var tk = PropertiesService.getScriptProperties().getProperty('LINE_PUSH_TOKEN'), v = {};
    // ตรวจการ์ดทุกแบบด้วยข้อมูลใบจริง (ใส่ค่าสมมติในช่องที่ยังว่าง) — ส่งไป validate เท่านั้น ไม่มีใครได้รับ
    var fake = Object.assign({}, b, { plate:b.plate||'กข 1234', car_model:b.car_model||'Toyota', driver_name:b.driver_name||'ทดสอบ',
      meet_point:b.meet_point||'หน้าอาคาร', reject_reason:b.reject_reason||'หาคนไปแทน', reject_by:b.reject_by||'ทดสอบ', reject_step:b.reject_step||'head' });
    ['head','exec','hr','ok','rej','info'].forEach(function(m){
      try{
        var r = UrlFetchApp.fetch('https://api.line.me/v2/bot/message/validate/push', { method:'post', contentType:'application/json',
          headers:{ Authorization:'Bearer ' + tk }, muteHttpExceptions:true,
          payload:JSON.stringify({ messages:[{ type:'flex', altText:'test', contents:carBubble_(fake, m) }] }) });
        v[m] = r.getResponseCode() === 200 ? 'ok' : String(r.getContentText()).slice(0, 400);
      }catch(e){ v[m] = String(e); }
    });
    // โควตาข้อความของ OA เดือนนี้ (เต็มแล้ว LINE จะตอบ 429 และการ์ดหายเงียบ)
    var quota = {};
    ['quota','quota/consumption'].forEach(function(q){
      try{ var r = UrlFetchApp.fetch('https://api.line.me/v2/bot/message/' + q, { headers:{ Authorization:'Bearer ' + tk }, muteHttpExceptions:true });
        quota[q] = r.getResponseCode() + ' ' + String(r.getContentText()).slice(0, 200); }catch(e){ quota[q] = String(e); }
    });
    var sp = PropertiesService.getScriptProperties();
    var s = carSheet_(), head = s.getRange(1,1,1,s.getLastColumn()).getValues()[0];
    return { bookings:rows.length, last:b.booking_no, status:b.status, department:b.department, recipients:to, card_mode:mode, line_validate:v, quota:quota,
      steps:{ head_name:b.head_name, head_at:b.head_at, has_head_id:!!b.head_id, exec_name:b.exec_name, exec_at:b.exec_at, updated_at:b.updated_at, created_at:b.created_at },
      header_ok:CAR_HEAD.every(function(h){ return head.indexOf(h) >= 0; }), header_count:head.length, missing:CAR_HEAD.filter(function(h){ return head.indexOf(h) < 0; }),
      cell:(function(){ try{ var r = b._row, st = s.getRange(r, head.indexOf('status')+1), up = s.getRange(r, head.indexOf('updated_at')+1);
        var dv = st.getDataValidation();
        return { status_raw:String(st.getValue()), status_dropdown:!!dv, updated_raw_type:(up.getValue() instanceof Date ? 'date' : typeof up.getValue()),
                 updated_fmt:up.getNumberFormat(), status_fmt:st.getNumberFormat() }; }catch(err){ return { error:String(err) }; } })(),
      push_fail_last:sp.getProperty('PUSH_FAIL_LAST') || '', push_fail_count:sp.getProperty('PUSH_FAIL_COUNT') || '' };
  },
  // ส่งการ์ดของขั้นปัจจุบันซ้ำ (เช่น หลังโควตา LINE กลับมา) — ผู้ขอหรือ HR กดได้
  car_resend: function(p){
    var e = carEmp_(p.caller); denyIf(!e, 'ไม่พบบัญชีพนักงาน');
    var b = carGet_(p.booking_no); if(!b) throw 'ไม่พบใบจอง';
    denyIf(String(b.requester_id) !== String(p.caller) && !carIsHR_(e), 'เฉพาะผู้ขอหรือ HR');
    denyIf(!/^pending_/.test(String(b.status)), 'ใบนี้ไม่ได้รออนุมัติแล้ว');
    carRouteNotify_(b.booking_no);
    return { ok:true, status:b.status };
  },

  // ข้อมูลตั้งต้นของหน้า: สิทธิ์ของผู้ใช้ + ใบนี้จะส่งไปที่ใคร + ตัวเลขคิวที่รอ
  car_init: function(p){
    var e = carEmp_(p.caller); denyIf(!e, 'ไม่พบบัญชีพนักงาน');
    var heads = carDeptHeads_(e.department).filter(function(h){ return String(h.line_user_id) !== String(p.caller); });
    var rows = carRows_();
    var q = rows.filter(function(r){ return carStepFor_(r, e); }).length;
    var hrq = carIsHR_(e) ? rows.filter(function(r){ return r.status === 'pending_hr'; }).length : 0;
    return { me:{ name:e.full_name, department:e.department, phone:e.phone||'', is_head:carIsHead_(e),
                  is_exec:/approv|exec|manager|บริหาร/i.test(String(e.role||'')), is_hr:carIsHR_(e) },
             route:{ heads:heads.map(function(h){ return h.full_name; }), head_exec:heads.filter(carIsExec_).map(function(h){ return h.full_name; }),
                     self_head:carIsHead_(e) },
             queue:q, hr_queue:hrq };
  },

  car_create: function(p){
    var e = carEmp_(p.caller); denyIf(!e, 'ไม่พบบัญชีพนักงาน');
    var need = function(v, m){ if(!String(v||'').trim()) throw m; };
    need(p.destination, 'กรุณากรอกปลายทาง'); need(p.go_at, 'กรุณาเลือกวันเวลาออก'); need(p.purpose, 'กรุณากรอกวัตถุประสงค์');
    if(!/^\d{4}-\d{2}-\d{2} \d{2}:\d{2}$/.test(String(p.go_at))) throw 'รูปแบบวันเวลาออกไม่ถูกต้อง';
    if(p.back_at && String(p.back_at) < String(p.go_at)) throw 'เวลากลับต้องหลังเวลาออก';
    // หัวหน้าแผนกจองเอง หรือแผนกยังไม่มีหัวหน้า -> ข้ามไปผู้บริหาร (บันทึกเหตุผลไว้ในช่องหัวหน้า)
    var heads = carDeptHeads_(e.department).filter(function(h){ return String(h.line_user_id) !== String(p.caller); });
    var skip = carIsHead_(e) ? 'หัวหน้าแผนกจองเอง' : (!heads.length ? 'แผนกนี้ยังไม่ได้ตั้งหัวหน้า' : '');
    var no = carNo_();
    appendObj(CAR_SHEET, {
      booking_no:no, status:(skip ? 'pending_exec' : 'pending_head'),
      requester_id:p.caller, requester_name:e.full_name||'', department:e.department||'', phone:String(p.phone||e.phone||''),
      destination:String(p.destination).trim(), origin:String(p.origin||'').trim(),
      go_at:String(p.go_at), back_at:String(p.back_at||''), pax:String(Number(p.pax)||1),
      need_driver:String(p.need_driver === true || String(p.need_driver) === 'true'), car_type:String(p.car_type||''),
      purpose:String(p.purpose).trim(), passengers:String(p.passengers||''), note:String(p.note||''),
      head_name:(skip ? ('— ' + skip) : ''), head_at:(skip ? now() : ''),
      created_at:now(), updated_at:now()
    });
    try{ carRouteNotify_(no); }catch(err){}   // แจ้งเตือนพลาด (เช่น โควตาไลน์เต็ม) ต้องไม่ทำให้การบันทึกล้ม
    return { booking_no:no, status:(skip ? 'pending_exec' : 'pending_head'), skipped_head:skip };
  },

  // scope: mine | queue (ใบที่ฉันต้องตัดสิน) | hr (งาน HR) | all (HR ดูทั้งหมด)
  car_list: function(p){
    var e = carEmp_(p.caller); denyIf(!e, 'ไม่พบบัญชีพนักงาน');
    var rows = carRows_(), sc = String(p.scope||'mine');
    var out;
    if(sc === 'queue') out = rows.filter(function(r){ return carStepFor_(r, e); });
    else if(sc === 'hr'){ denyIf(!carIsHR_(e)); out = rows.filter(function(r){ return ['pending_hr','assigned','in_use'].indexOf(String(r.status)) >= 0; }); }
    else if(sc === 'all'){ denyIf(!carIsHR_(e) && !/approv|exec|manager|บริหาร/i.test(String(e.role||''))); out = rows; }
    else out = rows.filter(function(r){ return String(r.requester_id) === String(p.caller) || String(r.driver_id) === String(e.id); });
    out.sort(function(a,b){ return String(b.created_at).localeCompare(String(a.created_at)); });
    var lim = Math.min(Number(p.limit) || 200, 2000);   // หน้าประวัติขอได้มากกว่ารายการปกติ
    return out.slice(0, lim).map(function(r){ var o = strip(r); o.can = carStepFor_(r, e); return o; });
  },

  car_get: function(p){
    var e = carEmp_(p.caller); denyIf(!e, 'ไม่พบบัญชีพนักงาน');
    var b = carGet_(p.booking_no); if(!b) throw 'ไม่พบใบจอง ' + p.booking_no;
    var mine = String(b.requester_id) === String(p.caller), drv = String(b.driver_id) === String(e.id);
    var step = carStepFor_(b, e), hr = carIsHR_(e);
    denyIf(!mine && !drv && !step && !hr && !/approv|exec|manager|บริหาร/i.test(String(e.role||'')) && !(carIsHead_(e) && e.department === b.department), 'ไม่มีสิทธิ์ดูใบนี้');
    var o = strip(b);
    o.perm = { decide:step, hr:hr, mine:mine, driver:drv,
      trip:(mine || drv || hr) && (b.status === 'assigned' || b.status === 'in_use'),
      cancel:mine && ['pending_head','pending_exec','pending_hr','assigned'].indexOf(String(b.status)) >= 0 };
    return o;
  },

  car_decide: function(p){
    return carDecide_(p.booking_no, p.caller, p.decision === 'approve' ? 'approve' : 'reject', p.reason, p.note);
  },

  // รถที่ HR เปิดให้จอง พร้อมบอกว่าช่วงเวลานี้คันไหนติดจอง
  car_fleet: function(p){
    var e = carEmp_(p.caller); denyIf(!carIsHR_(e), 'เฉพาะ HR');
    var list = carFleet_();
    if(String(p.all) !== '1') list = list.filter(function(v){ return v.bookable; });
    return list.map(function(v){
      var c = p.go_at ? carClashes_(v.vehicle_key, p.go_at, p.back_at, p.except) : [];
      delete v._row;
      v.busy = c.map(function(r){ return { booking_no:r.booking_no, go_at:r.go_at, back_at:r.back_at, requester:r.requester_name }; });
      return v;
    });
  },
  car_fleet_set: function(p){
    var e = carEmp_(p.caller); denyIf(!carIsHR_(e), 'เฉพาะ HR');
    var v = carFleet_().filter(function(x){ return x.vehicle_key === p.vehicle_key; })[0]; if(!v) throw 'ไม่พบรถ ' + p.vehicle_key;
    var s = sh(SHEETS.VEH), head = s.getRange(1,1,1,s.getLastColumn()).getValues()[0];
    var put = function(col, val){ var c = head.indexOf(col); if(c >= 0) s.getRange(v._row, c+1).setValue(val); };
    if(p.bookable !== undefined) put('car_bookable', String(p.bookable) === 'true');
    if(p.car_type !== undefined) put('car_type', String(p.car_type));
    if(p.seats !== undefined) put('car_seats', Number(p.seats) || '');
    return { ok:true };
  },
  // รายชื่อพนักงานให้ HR ค้นเลือกคนขับ
  car_people: function(p){
    var e = carEmp_(p.caller); denyIf(!carIsHR_(e), 'เฉพาะ HR');
    return getRows(SHEETS.EMP).filter(function(r){ return String(r.active) === 'true' && r.full_name; })
      .map(function(r){ return { id:r.id, name:r.full_name, nickname:r.nickname||'', dept:r.department||'', phone:String(r.phone||'') }; });
  },

  car_assign: function(p){
    var e = carEmp_(p.caller); denyIf(!carIsHR_(e), 'เฉพาะ HR');
    var b = carGet_(p.booking_no); if(!b) throw 'ไม่พบใบจอง';
    denyIf(['pending_hr','assigned'].indexOf(String(b.status)) < 0, 'ใบนี้ยังไม่ถึงขั้น HR หรือดำเนินการไปแล้ว');
    var v = carFleet_().filter(function(x){ return x.vehicle_key === p.vehicle_key; })[0]; if(!v) throw 'ไม่พบรถที่เลือก';
    var c = carClashes_(v.vehicle_key, b.go_at, b.back_at, b.booking_no);
    if(c.length) throw 'รถคันนี้ติดจอง ' + c[0].booking_no + ' ในช่วงเวลาเดียวกัน';
    var drv = p.driver_id ? getRows(SHEETS.EMP).filter(function(r){ return String(r.id) === String(p.driver_id); })[0] : null;
    denyIf(String(b.substitute || '') && !drv && !String(p.driver_name || '').trim(), 'ใบนี้ผู้อนุมัติให้หาคนไปแทน — กรุณาเลือกคนไปแทน (คนขับ)');
    carPatch_(b.booking_no, { status:'assigned', vehicle_key:v.vehicle_key, plate:v.plate, car_model:v.model,
      driver_id:drv ? drv.id : '', driver_name:drv ? drv.full_name : String(p.driver_name||''),
      driver_phone:drv ? String(drv.phone||'') : String(p.driver_phone||''),
      meet_point:String(p.meet_point||''), hr_id:p.caller, hr_name:e.full_name||'', hr_at:now(), hr_note:String(p.hr_note||'') });
    var nb = carGet_(b.booking_no);
    var ids = [nb.requester_id];
    if(drv && drv.line_user_id && String(drv.line_user_id) !== String(nb.requester_id)) ids.push(drv.line_user_id);
    try{ carPush_(ids, nb, 'ok', 'การจองรถได้รับอนุมัติ · พร้อมเดินทาง', '✅ จองรถสำเร็จ ' + nb.plate + ' · ' + carWhen_(nb)); }catch(err){}
    return { ok:true };
  },

  // บันทึกไมล์: kind=out (ออกรถ) | in (คืนรถ)
  car_trip: function(p){
    var e = carEmp_(p.caller); denyIf(!e, 'ไม่พบบัญชีพนักงาน');
    var b = carGet_(p.booking_no); if(!b) throw 'ไม่พบใบจอง';
    var ok = String(b.requester_id) === String(p.caller) || String(b.driver_id) === String(e.id) || carIsHR_(e);
    denyIf(!ok, 'เฉพาะผู้ขอ คนขับ หรือ HR');
    var m = Number(String(p.mileage||'').replace(/,/g,'')); denyIf(!(m > 0), 'กรุณากรอกเลขไมล์');
    if(p.kind === 'out'){
      denyIf(b.status !== 'assigned', 'ใบนี้ยังไม่พร้อมออกรถ หรือออกไปแล้ว');
      carPatch_(b.booking_no, { status:'in_use', out_at:now(), out_mileage:m });
      return { ok:true, status:'in_use' };
    }
    denyIf(b.status !== 'in_use', 'ต้องบันทึกออกรถก่อน');
    var out = Number(b.out_mileage) || 0; denyIf(m < out, 'เลขไมล์ตอนกลับน้อยกว่าตอนออก (' + out + ')');
    carPatch_(b.booking_no, { status:'returned', in_at:now(), in_mileage:m, in_condition:String(p.condition||'ปกติ'), distance_km:(out ? (m - out) : '') });
    // ไมล์ล่าสุดไปไว้ที่รถ ใช้กับหน้าอื่นต่อได้
    try{
      var v = carFleet_().filter(function(x){ return x.vehicle_key === b.vehicle_key; })[0];
      if(v){ var s = sh(SHEETS.VEH), head = s.getRange(1,1,1,s.getLastColumn()).getValues()[0];
        s.getRange(v._row, head.indexOf('last_mileage')+1).setValue(m);
        s.getRange(v._row, head.indexOf('last_mileage_at')+1).setValue(now()); }   // ห้าม setNumberFormat: ชีตที่เป็นตารางจะพังตอน flush
    }catch(err){}
    if(/ปัญหา/.test(String(p.condition||''))) notifyAdmins('⚠️ คืนรถ ' + b.plate + ' (' + b.booking_no + ') แจ้งว่ารถมีปัญหา\n' + liffUrl('car=' + b.booking_no));
    return { ok:true, status:'returned', distance:(out ? m - out : '') };
  },

  car_cancel: function(p){
    var b = carGet_(p.booking_no); if(!b) throw 'ไม่พบใบจอง';
    denyIf(String(b.requester_id) !== String(p.caller), 'ยกเลิกได้เฉพาะผู้ขอ');
    denyIf(['pending_head','pending_exec','pending_hr','assigned'].indexOf(String(b.status)) < 0, 'ใบนี้ยกเลิกไม่ได้แล้ว');
    carPatch_(b.booking_no, { status:'cancelled', cancelled_at:now() });
    if(b.status === 'assigned') notifyAdmins('↩️ ผู้ขอยกเลิกใบจองรถ ' + b.booking_no + ' (' + (b.plate||'') + ' · ' + carWhen_(b) + ')');
    return { ok:true };
  }
});

// งานที่เขียนชีตทั้งหมดต้องต่อคิวกัน (เลขใบ / สถานะ / รถ ไม่ชนกัน)
['car_create','car_assign','car_trip','car_cancel','car_resend'].forEach(function(k){
  var f = API[k]; API[k] = function(p){ return carLocked_(function(){ return f(p); }); };
});
// ซ่อมใบที่ค้างครึ่งทาง: สถานะเป็น "รอผู้บริหาร" แต่ไม่มีการอนุมัติของหัวหน้าบันทึกไว้เลย -> กลับไปรอหัวหน้า
// (แถวที่ข้ามขั้นหัวหน้าโดยตั้งใจจะมีข้อความ "— …" ในช่อง head_name จึงไม่โดนแตะ) เรียกซ้ำได้
API.car_repair = function(){
  return carLocked_(function(){
    var fixed = [];
    carRows_().forEach(function(r){
      if(r.status === 'pending_exec' && !r.head_id && !String(r.head_name||'').trim()){
        carPatch_(r.booking_no, { status:'pending_head' }); fixed.push(r.booking_no);
      }
      // ใบที่ถูกบันทึกเป็น "ไม่อนุมัติ · หาคนไปแทน" ก่อนเปลี่ยนกติกา -> เดินต่อไปให้ HR จัดคนไปแทนในใบเดิม
      if(r.status === 'rejected' && String(r.reject_reason) === CAR_RS.sub){
        var who = String(r.reject_by || ''), at = r.rejected_at || now(), head = r.reject_step === 'head';
        var ex = head && carIsExec_(getRows(SHEETS.EMP).filter(function(x){ return String(x.full_name) === who; })[0]);
        var sp = { status:(head && !ex) ? 'pending_exec' : 'pending_hr',
          substitute:'หาคนไปแทน · ' + who + (head ? ' (หัวหน้าแผนก)' : ' (ผู้อนุมัติ)') + (r.reject_note ? (' · ' + r.reject_note) : ''),
          reject_step:'', reject_by:'', reject_reason:'', reject_note:'', rejected_at:'' };
        if(head){ sp.head_name = who; sp.head_at = at; }
        if(!head || ex){ sp.exec_name = who; sp.exec_at = at; }
        carPatch_(r.booking_no, sp); fixed.push(r.booking_no + ' -> ' + sp.status);
        var nb = carGet_(r.booking_no);
        try{ carRouteNotify_(r.booking_no); }catch(err){}
        var nx = sp.status === 'pending_hr' ? 'HR จัดคนไปแทน' : 'ผู้อนุมัติ';
        try{ carPush_([nb.requester_id], nb, 'sub', 'ให้หาคนไปแทน · ส่งต่อ ' + nx, '👥 ใบ ' + nb.booking_no + ' ให้หาคนไปแทน · ส่งต่อ ' + nx); }catch(err){}
      }
    });
    return { fixed:fixed };
  });
};

// ---------- LINE: ปุ่มในการ์ด + ถามเหตุผลตอนไม่อนุมัติ ----------
function lineReplyMsgs_(token, messages){
  var tk = PropertiesService.getScriptProperties().getProperty('LINE_PUSH_TOKEN');
  if(!tk || !token) return;
  try{
    UrlFetchApp.fetch('https://api.line.me/v2/bot/message/reply', { method:'post', contentType:'application/json',
      headers:{ Authorization:'Bearer ' + tk }, payload:JSON.stringify({ replyToken:token, messages:messages }), muteHttpExceptions:true });
  }catch(e){}
}
var CAR_RS = { sub:'หาคนไปแทน', other:'อื่นๆ' };
/** postback ของใบจองรถ (act ขึ้นต้น car_) — คืน true ถ้าจัดการแล้ว */
function carPostback_(d, uid, reply){
  if(!/^car_/.test(String(d.act||''))) return false;
  var e = carEmp_(uid);
  if(!e){ lineReply(reply, '⛔ บัญชีนี้ยังไม่ได้ลงทะเบียน'); return true; }
  var b = carGet_(d.t);
  if(!b){ lineReply(reply, '❌ ไม่พบใบจอง ' + d.t); return true; }
  var step = carStepFor_(b, e);
  if(!step){
    lineReply(reply, /^pending_/.test(String(b.status)) ? '⛔ บัญชีนี้ไม่มีสิทธิ์ตัดสินใบนี้ในขั้นนี้'
      : ('ℹ️ ใบ ' + d.t + ' ดำเนินการไปแล้ว (' + (CAR_STEP_TH[b.status] || b.status) + ')'));
    return true;
  }
  if(d.act === 'car_ok'){
    var r = carDecide_(d.t, uid, 'approve');
    lineReply(reply, '✅ อนุมัติใบจองรถ ' + d.t + '\nโดย ' + r.by + '\n\nส่งต่อ: ' + (CAR_STEP_TH[r.status] || r.status));
    return true;
  }
  if(d.act === 'car_no'){
    // ถามเหตุผลด้วยปุ่มตอบกลับด่วน
    var qi = function(label, key){ return { type:'action', action:{ type:'postback', label:label, data:pbData('car_rs', d.t) + '&r=' + key, displayText:label } }; };
    lineReplyMsgs_(reply, [{ type:'text', text:'ไม่อนุมัติใบ ' + d.t + '\nเลือกเหตุผลด้านล่าง 👇',
      quickReply:{ items:[ qi('👥 หาคนไปแทน', 'sub'), qi('✏️ อื่นๆ (พิมพ์เหตุผล)', 'other'), qi('ยกเลิก', 'cancel') ] } }]);
    return true;
  }
  if(d.act === 'car_rs'){
    if(d.r === 'cancel'){ lineReply(reply, 'ยกเลิกแล้ว · ใบ ' + d.t + ' ยังรอพิจารณาอยู่'); return true; }
    if(d.r === 'other'){
      CacheService.getScriptCache().put('carrs_' + uid, d.t, 600);
      lineReply(reply, '✏️ พิมพ์เหตุผลที่ไม่อนุมัติใบ ' + d.t + ' ส่งมาได้เลย (ภายใน 10 นาที)');
      return true;
    }
    var rs = carDecide_(d.t, uid, 'reject', CAR_RS.sub, '');
    lineReply(reply, '👥 ใบ ' + d.t + ' ให้หาคนไปแทน\nส่งต่อ: ' + (rs.status === 'pending_hr' ? 'HR จัดคนไปแทน + รถ' : (CAR_STEP_TH[rs.status] || rs.status)) + '\nแจ้งผู้ขอเรียบร้อย');
    return true;
  }
  return false;
}
/** ข้อความที่พิมพ์ตามหลังการเลือก "อื่นๆ" = เหตุผลที่ไม่อนุมัติ — คืน true ถ้ากินข้อความนี้ไปแล้ว */
function carReasonText_(ev){
  var uid = ev.source && ev.source.userId; if(!uid) return false;
  var c = CacheService.getScriptCache(), no = c.get('carrs_' + uid);
  if(!no) return false;
  var txt = String(ev.message && ev.message.text || '').trim();
  if(!txt || /^เมนู:/.test(txt)) return false;   // แตะเมนูระหว่างรอ = ไม่ใช่เหตุผล
  c.remove('carrs_' + uid);
  try{
    carDecide_(no, uid, 'reject', CAR_RS.other, txt);
    lineReply(ev.replyToken, '✕ ไม่อนุมัติใบ ' + no + '\nเหตุผล: ' + txt + '\nแจ้งผู้ขอเรียบร้อย');
  }catch(err){ lineReply(ev.replyToken, 'ℹ️ ' + String(err)); }
  return true;
}
