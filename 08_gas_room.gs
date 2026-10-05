// ============================================================
//  จองห้องประชุม (บ้านกาแฟ) — ข้อมูลอยู่ใน Supabase ของ To-do List (rooms / bookings / meetings)
//  ใช้ตารางชุดเดียวกับ To-do: จองจากแอปไหนก็เห็นตรงกัน · กันเวลาซ้อนด้วย constraint no_overlap ในฐานข้อมูล
//  พนักงาน (มีชื่อในระบบ) จอง -> ผู้อนุมัติ (role Approver) อนุมัติ/ไม่อนุมัติ -> แจ้งผู้จอง
//  แจ้งแอดมินล่วงหน้า 2 วันก่อนวันประชุม (เกาะ trigger เดิมของ pollTodoDone วันละครั้ง)
//  status: tentative = รออนุมัติ (กันช่วงเวลาไว้แล้ว) · confirmed = อนุมัติ · cancelled = ไม่อนุมัติ/ยกเลิก (ดู decision)
// ============================================================
var RM_PURPOSE = { training:'อบรม / ประชุม', audit:'Audit', guest:'รับรองแขก', other:'อื่น ๆ' };
var RM_OPEN = '07:00', RM_CLOSE = '20:00';

// ---------- สิทธิ์ห้องประชุม (ติ๊กในทะเบียนพนักงาน) ----------
// room_admin = ได้แจ้งล่วงหน้า 2 วัน + จัดการ/ยกเลิกการจอง · room_approver = อนุมัติการจอง
// ยังไม่ติ๊กให้ใครเลย -> ใช้ role เดิม (Admin / Approver) ระบบไม่เงียบระหว่างเปลี่ยนผ่าน
function rmBool_(v){ return v === true || String(v).toUpperCase() === 'TRUE'; }
function rmTicked_(col){ return getRows(SHEETS.EMP).filter(function(r){ return String(r.active) === 'true' && rmBool_(r[col]); }); }
function rmIsApprover_(e){ if(!e) return false; return rmTicked_('room_approver').length ? rmBool_(e.room_approver) : carIsExec_(e); }
function rmIsAdmin_(e){ if(!e) return false; return rmTicked_('room_admin').length ? rmBool_(e.room_admin) : carIsHR_(e); }
function rmApproverIds_(){ var t = rmTicked_('room_approver'); return t.length ? t.map(function(r){ return r.line_user_id; }).filter(Boolean) : approverIds(); }
// room_food = คนจัดการสั่งอาหาร/เบรค (ได้การ์ดเมื่อใบที่มีอาหารอนุมัติ/ยกเลิก + แจ้งล่วงหน้า 2 วัน) · ไม่มีคนติ๊ก = ไม่ส่งใคร
function rmFoodIds_(){ return rmTicked_('room_food').map(function(r){ return r.line_user_id; }).filter(Boolean); }
function rmFoodNames_(){ return rmTicked_('room_food').map(function(r){ return r.full_name; }); }
// ยังไม่มีใครติ๊กผู้สั่งอาหาร -> ส่งแอดมินห้องแทน (การ์ดสั่งอาหารต้องไม่หายเงียบ · เคยเกิด 2 ต.ค. RM-202610-002)
function rmFoodTargets_(){ var ids = rmFoodIds_(); return ids.length ? ids : rmAdminIds_(); }
function rmHasFood_(b){ return !!(Number(b.food_break_am) || Number(b.food_lunch) || Number(b.food_break_pm)); }
// เบรค "น้ำ+ขนมปัง" ต้องสั่งล่วงหน้า 5 วัน (นับจากวันนี้ถึงวันประชุม)
var RM_BREAD = 'น้ำ+ขนมปัง', RM_BREAD_DAYS = 5;
function rmDaysAhead_(d){ var t = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd'); return Math.round((Date.parse(d + 'T00:00:00Z') - Date.parse(t + 'T00:00:00Z')) / 86400000); }
// ผู้ยืนยันการจอง (กดรับทราบ) = ผู้อนุมัติห้อง หรือ แอดมินบ้านกาแฟ — การจองเช็คห้องว่างให้แล้ว ขั้นนี้แค่ให้แอดมินรับทราบ (ผู้ใช้เปลี่ยน 5 ต.ค.)
function rmCanConfirm_(e){ return rmIsApprover_(e) || rmIsAdmin_(e); }
function rmAdminIds_(){ var t = rmTicked_('room_admin'); return t.length ? t.map(function(r){ return r.line_user_id; }).filter(Boolean) : adminIds(); }

// ---------- Supabase REST ----------
function rmReq_(method, path, body, prefer){
  var c = sbProps(); if(!c.url || !c.key) throw 'ยังไม่ได้ตั้งค่า Supabase';
  var h = sbHeaders(c.key); if(prefer) h.Prefer = prefer;
  var o = { method:method, headers:h, muteHttpExceptions:true };
  if(body !== undefined){ o.contentType = 'application/json'; o.payload = JSON.stringify(body); }
  var r = UrlFetchApp.fetch(c.url + '/rest/v1/' + path, o), txt = r.getContentText();
  var j = null; try{ j = JSON.parse(txt || 'null'); }catch(e){}
  return { code:r.getResponseCode(), data:j, text:txt };
}
function rmGet_(path){ var r = rmReq_('get', path); if(r.code >= 300) throw 'อ่านข้อมูลห้องไม่ได้ (' + r.code + ')'; return r.data || []; }

// ---------- เวลา (เก็บ timestamptz · แสดงเวลาไทย) ----------
function rmIso_(date, time){ return date + 'T' + time + ':00+07:00'; }
function rmLocal_(iso){ return Utilities.formatDate(new Date(iso), TZ, 'yyyy-MM-dd HH:mm'); }
function rmDayRange_(date){ return { from:rmIso_(date, '00:00'), to:rmIso_(Utilities.formatDate(new Date(new Date(date + 'T12:00:00+07:00').getTime() + 86400000), TZ, 'yyyy-MM-dd'), '00:00') }; }
function rmWhen_(b){
  var s = rmLocal_(b.start_at), e = rmLocal_(b.end_at), p = s.slice(0,10).split('-');
  var dow = ['อา','จ','อ','พ','พฤ','ศ','ส'][new Date(+p[0], p[1]-1, +p[2]).getDay()];
  return dow + ' ' + (+p[2]) + '/' + (+p[1]) + ' · ' + s.slice(11,16) + '–' + e.slice(11,16);
}
function rmStatus_(b){
  if(b.status === 'tentative') return 'pending';
  if(b.status === 'confirmed') return 'approved';
  return b.decision === 'rejected' ? 'rejected' : 'cancelled';
}
function rmPub_(b, rooms){
  var r = (rooms || []).filter(function(x){ return x.id === b.room_id; })[0] || {};
  return { id:b.id, booking_no:b.booking_no, room_id:b.room_id, room:r.name || '', color:r.color || '', capacity:r.capacity || 0,
    start:rmLocal_(b.start_at), end:rmLocal_(b.end_at), when:rmWhen_(b), state:rmStatus_(b),
    title:b.title || '', purpose:b.purpose || '', purpose_other:b.purpose_other || '', headcount:b.headcount || 0,
    food_break_am:b.food_break_am || 0, food_lunch:b.food_lunch || 0, food_break_pm:b.food_break_pm || 0, food_note:b.food_note || '',
    food_am_type:b.food_am_type || '', food_lunch_type:b.food_lunch_type || '', food_pm_type:b.food_pm_type || '',
    requester_name:b.requester_name || '', department:b.department || '', phone:b.phone || '', mine:false,
    approver_name:b.approver_name || '', decided_at:b.decided_at ? rmLocal_(b.decided_at) : '', reject_reason:b.reject_reason || '',
    source:b.source || 'myssb', meeting_id:b.meeting_id || '', created_at:b.created_at ? rmLocal_(b.created_at) : '' };
}
function rmRooms_(){ return rmGet_('rooms?active=eq.true&select=id,name,capacity,site,sort,color,description,facilities&order=sort.asc'); }
function rmGetBooking_(id){
  var a = rmGet_('bookings?id=eq.' + encodeURIComponent(id) + '&select=*'); return a[0] || null;
}
function rmPatch_(id, patch){
  patch.updated_at = new Date().toISOString();
  var r = rmReq_('patch', 'bookings?id=eq.' + encodeURIComponent(id), patch, 'return=representation');
  if(r.code >= 300) throw 'บันทึกไม่สำเร็จ (' + r.code + ')';
  return (r.data || [])[0];
}
function rmNo_(){
  var prefix = 'RM-' + Utilities.formatDate(new Date(), TZ, 'yyyyMM') + '-';
  var a = rmGet_('bookings?booking_no=like.' + encodeURIComponent(prefix + '*') + '&select=booking_no&order=booking_no.desc&limit=1');
  var n = a[0] ? parseInt(String(a[0].booking_no).slice(prefix.length), 10) || 0 : 0;
  return prefix + ('000' + (n + 1)).slice(-3);
}

// ---------- สำเนาประวัติลงชีต RoomBookings ----------
// ข้อมูลจริงอยู่ใน Supabase — ชีตนี้ไว้ให้ HR เปิดดู/ทำรายงานเอง (เขียนทับทั้งแถวตาม id ทุกครั้งที่สถานะเปลี่ยน)
// ห้าม setNumberFormat (ผู้ใช้มักแปลงชีตเป็นตาราง) · ชีตนี้ไม่ถูกอ่านกลับเข้าระบบ
var RM_SHEET = 'RoomBookings';
var RM_SHEET_HEAD = ['booking_no','status','room','date','start','end','headcount','purpose','title',
  'food_break_am','food_lunch','food_break_pm','food_note','requester_name','department','phone',
  'approver_name','decided_at','reject_reason','source','todo_meeting','created_at','updated_at','id'];
var RM_STATE_TH = { pending:'รอยืนยัน', approved:'ยืนยันแล้ว', rejected:'ไม่อนุมัติ', cancelled:'ยกเลิก' };
function rmSheetRow_(b, rooms){
  var p = rmPub_(b, rooms);
  return [p.booking_no, RM_STATE_TH[p.state] || p.state, p.room, p.start.slice(0,10), p.start.slice(11,16), p.end.slice(11,16), p.headcount,
    (p.purpose === 'other' ? ('อื่น ๆ: ' + p.purpose_other) : (RM_PURPOSE[p.purpose] || p.purpose)), p.title,
    p.food_break_am, p.food_lunch, p.food_break_pm, [rmFoodTypes_(p), p.food_note].filter(String).join(' · '), p.requester_name, p.department, p.phone,
    p.approver_name, p.decided_at, p.reject_reason, (p.source === 'todo' ? 'To-do' : 'MySSB'), (p.meeting_id ? 'มี' : ''),
    p.created_at, b.updated_at ? rmLocal_(b.updated_at) : p.created_at, p.id];
}
function rmSheet_(){
  var s = sh(RM_SHEET);
  if(!s){ s = ss().insertSheet(RM_SHEET); s.getRange(1,1,1,RM_SHEET_HEAD.length).setValues([RM_SHEET_HEAD]).setFontWeight('bold'); s.setFrozenRows(1); }
  return s;
}
function rmSyncSheet_(list){
  try{
    if(!list || !list.length) return 0;
    var s = rmSheet_(), rooms = rmRooms_(), last = s.getLastRow(), n = RM_SHEET_HEAD.length;
    var ids = last > 1 ? s.getRange(2, n, last - 1, 1).getValues().map(function(r){ return String(r[0]); }) : [];
    list.forEach(function(b){
      if(!b || !b.id) return;
      var row = rmSheetRow_(b, rooms), at = ids.indexOf(String(b.id));
      if(at >= 0) s.getRange(at + 2, 1, 1, n).setValues([row]);
      else { s.appendRow(row); ids.push(String(b.id)); }
    });
    return list.length;
  }catch(e){ return 0; }   // สำเนาพลาดต้องไม่ทำให้การจองล้ม
}

// ---------- การ์ดไลน์ ----------
function rmFood_(b){
  var f = [], t = function(x){ return x ? (' (' + x + ')') : ''; };
  if(b.food_break_am) f.push('เบรคเช้า ' + b.food_break_am + t(b.food_am_type));
  if(b.food_lunch) f.push('อาหารเที่ยง ' + b.food_lunch + t(b.food_lunch_type));
  if(b.food_break_pm) f.push('เบรคบ่าย ' + b.food_break_pm + t(b.food_pm_type));
  return f.join(' · ');
}
// ชนิดอาหาร (น้ำ+ขนมปัง / น้ำ / สไตล์คุณศุภชัย / ที่ระบุเอง) สำหรับชีตประวัติ
function rmFoodTypes_(b){
  return [b.food_break_am && b.food_am_type ? ('เช้า: ' + b.food_am_type) : '', b.food_lunch && b.food_lunch_type ? ('เที่ยง: ' + b.food_lunch_type) : '',
    b.food_break_pm && b.food_pm_type ? ('บ่าย: ' + b.food_pm_type) : ''].filter(String).join(' · ');
}
function rmBubble_(b, mode, rooms){
  var p = rmPub_(b, rooms);
  var color = { ask:['#0e7490','#1e3a8a'], ok:['#059669','#065f46'], rej:['#b91c1c','#7f1d1d'], food:['#b45309','#7c2d12'], foodx:['#b91c1c','#7f1d1d'] }[mode] || ['#334155','#0f172a'];
  var note = { ask:'รอแอดมินยืนยัน', ok:'ยืนยันแล้ว', rej:'ไม่อนุมัติ', food:'สั่งอาหาร', foodx:'ยกเลิกอาหาร' }[mode] || 'จองห้องประชุม';
  var title = { ask:'จองห้องประชุม · บ้านกาแฟ', ok:'การจองห้องได้รับการยืนยัน', rej:'การจองห้องไม่ได้รับอนุมัติ',
    food:'🍱 สั่งอาหารประชุม · บ้านกาแฟ', foodx:'✕ ยกเลิกการจอง · ยกเลิกอาหาร' }[mode] || 'จองห้องประชุม';
  var head = fxHead(title,
    p.room, p.when + ' · ' + (p.headcount || '-') + ' คน', color[0], color[1], note);
  var rows = [];
  if(mode === 'rej'){ rows.push(fxRow('เหตุผล', p.reject_reason || '-')); rows.push(fxRow('โดย', p.approver_name || '-')); }
  rows.push(fxRow('วัตถุประสงค์', (p.purpose === 'other' ? ('อื่น ๆ: ' + p.purpose_other) : (RM_PURPOSE[p.purpose] || '-')) + (p.title ? (' · ' + p.title) : '')));
  var food = rmFood_(p); if(food) rows.push(fxRow('อาหาร', food));
  if(p.food_note) rows.push(fxRow('หมายเหตุ', p.food_note));
  if(mode === 'ask' || mode === 'food' || mode === 'foodx') rows.push(fxRow('ผู้จอง', p.requester_name + (p.department ? (' · ' + p.department) : '') + (p.phone ? (' · ' + p.phone) : '')));
  if(mode === 'ok' || mode === 'food') rows.push(fxRow('ยืนยันโดย', p.approver_name || '-'));
  rows.push(fxRow('ใบจอง', p.booking_no));
  var btn = function(style, color, action){ var o = { type:'button', style:style, height:'sm', action:action }; if(color) o.color = color; return o; };
  var foot = [];
  if(mode === 'ask'){
    // ปุ่มเดียว: ยืนยันรับทราบ (ไม่มีปุ่มไม่อนุมัติแล้ว · การ์ดเก่าที่มีปุ่มไม่อนุมัติยังกดได้)
    foot.push({ type:'box', layout:'horizontal', spacing:'sm', contents:[
      btn('primary', '#22a06b', { type:'postback', label:'✓ ยืนยันรับทราบ', data:pbData('rm_ok', p.id), displayText:'✓ ยืนยันการจองห้อง ' + p.booking_no })
    ]});
  }
  foot.push(btn('secondary', null, { type:'uri', label:'ดูรายละเอียด', uri:liffUrl('room=' + p.id) }));
  return { type:'bubble', header:head, body:{ type:'box', layout:'vertical', paddingAll:'14px', spacing:'sm', contents:rows },
    footer:{ type:'box', layout:'vertical', spacing:'sm', paddingAll:'12px', contents:foot } };
}
function rmPush_(ids, b, mode, alt){
  var rooms = rmRooms_(), bubble = null;
  try{ bubble = rmBubble_(b, mode, rooms); }catch(e){}
  (ids || []).forEach(function(id){ if(!id) return; if(bubble) pushFlex(id, alt, bubble); else linePush(id, alt); });
}

// ---------- ตัดสิน (ใช้ทั้งเว็บและไลน์) ----------
function rmDecide_(id, uid, decision, reason){
  var e = carEmp_(uid); denyIf(!e, 'ไม่พบบัญชีพนักงาน');
  denyIf(!rmCanConfirm_(e), 'เฉพาะแอดมินบ้านกาแฟ / ผู้ยืนยันการจอง');
  var lk = LockService.getScriptLock(); if(!lk.tryLock(20000)) throw 'ระบบกำลังบันทึกรายการอื่น ลองใหม่อีกครั้ง';
  try{
    var b = rmGetBooking_(id); if(!b) throw 'ไม่พบการจอง';
    var who = String(e.full_name || '');
    // คนเดิมกดซ้ำ / แอปส่งซ้ำ -> ตอบว่าทำไปแล้ว
    if(b.status !== 'tentative'){
      if(String(b.approver_line) === String(uid)) return { ok:true, already:true, state:rmStatus_(b), by:who };
      throw 'การจองนี้ดำเนินการไปแล้ว (' + ({ approved:'ยืนยันแล้ว', rejected:'ไม่อนุมัติ', cancelled:'ยกเลิก' }[rmStatus_(b)]) + ')';
    }
    var now_ = new Date().toISOString(), nb;
    if(decision === 'approve'){
      nb = rmPatch_(id, { status:'confirmed', decision:'approved', approver_line:uid, approver_name:who, decided_at:now_ });
      try{ rmPush_([b.requester_line], nb, 'ok', '✅ จองห้องสำเร็จ · ' + rmWhen_(nb)); }catch(err){}
      if(rmHasFood_(nb)) try{ rmPush_(rmFoodTargets_(), nb, 'food', '🍱 สั่งอาหารประชุม · ' + rmWhen_(nb)); }catch(err){}
      rmSyncSheet_([nb]);
    } else {
      denyIf(!String(reason||'').trim(), 'ต้องระบุเหตุผลที่ไม่อนุมัติ');
      nb = rmPatch_(id, { status:'cancelled', decision:'rejected', approver_line:uid, approver_name:who, decided_at:now_, reject_reason:String(reason).trim() });
      try{ rmPush_([b.requester_line], nb, 'rej', '✕ การจองห้องไม่ได้รับอนุมัติ · ' + rmWhen_(nb)); }catch(err){}
      rmSyncSheet_([nb]);
    }
    return { ok:true, state:rmStatus_(nb), by:who };
  } finally { lk.releaseLock(); }
}

// ---------- แจ้งแอดมินล่วงหน้า 2 วัน (วันละครั้ง หลัง 08:00) ----------
function roomRemindDaily_(){
  var sp = PropertiesService.getScriptProperties(), today = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');
  if(Number(Utilities.formatDate(new Date(), TZ, 'H')) < 8 || sp.getProperty('RM_REMIND_DAY') === today) return;
  sp.setProperty('RM_REMIND_DAY', today);
  var target = Utilities.formatDate(new Date(Date.now() + 2 * 86400000), TZ, 'yyyy-MM-dd'), rg = rmDayRange_(target);
  var list = rmGet_('bookings?status=in.(tentative,confirmed)&start_at=gte.' + encodeURIComponent(rg.from) + '&start_at=lt.' + encodeURIComponent(rg.to) + '&select=*&order=start_at.asc');
  if(!list.length) return;
  var rooms = rmRooms_(), f = { am:0, lunch:0, pm:0 };
  var lines = list.map(function(b){
    var p = rmPub_(b, rooms); f.am += p.food_break_am; f.lunch += p.food_lunch; f.pm += p.food_break_pm;
    return fxRow(p.room.replace('ห้อง', ''), p.start.slice(11,16) + '–' + p.end.slice(11,16) + ' · ' + (p.headcount || '-') + ' คน'
      + (rmFood_(p) ? (' · ' + rmFood_(p)) : '') + (p.state === 'pending' ? ' (รอยืนยัน)' : ''));
  });
  var p0 = rmPub_(list[0], rooms);
  var head = fxHead('อีก 2 วัน · เตรียมห้องและอาหาร', p0.when.split(' · ')[0], list.length + ' การจอง · บ้านกาแฟ', '#b45309', '#7c2d12', 'จองห้องประชุม');
  var body = lines.concat([{ type:'separator', margin:'md' },
    fxRow('รวมอาหาร', 'เบรคเช้า ' + f.am + ' · เที่ยง ' + f.lunch + ' · เบรคบ่าย ' + f.pm)]);
  var bubble = { type:'bubble', header:head, body:{ type:'box', layout:'vertical', paddingAll:'14px', spacing:'sm', contents:body },
    footer:{ type:'box', layout:'vertical', paddingAll:'12px', contents:[{ type:'button', style:'secondary', height:'sm',
      action:{ type:'uri', label:'เปิดตารางห้อง', uri:liffUrl('go=roomboard') } }] } };
  var to = rmAdminIds_(); rmFoodIds_().forEach(function(id){ if(to.indexOf(id) < 0) to.push(id); });
  to.forEach(function(id){ pushFlex(id, '📅 อีก 2 วันมีประชุม ' + list.length + ' รายการ (บ้านกาแฟ)', bubble); });
  list.forEach(function(b){ try{ rmReq_('patch', 'bookings?id=eq.' + b.id, { admin_reminded_at:new Date().toISOString() }); }catch(e){} });
}

// ---------- API ----------
Object.assign(API, {
  room_init: function(p){
    var e = carEmp_(p.caller); denyIf(!e, 'ไม่พบบัญชีพนักงาน');
    var q = rmCanConfirm_(e) ? rmGet_('bookings?status=eq.tentative&select=id').length : 0;
    return { rooms:rmRooms_(), me:{ name:e.full_name, department:e.department||'', phone:String(e.phone||''), is_approver:rmIsApprover_(e), is_admin:rmIsAdmin_(e), can_confirm:rmCanConfirm_(e) },
             queue:q, open:RM_OPEN, close:RM_CLOSE, bread_days:RM_BREAD_DAYS };
  },
  // การจองในช่วงวันที่ (ใช้ทั้งหน้าจองและตารางห้อง) · from/to = 'yyyy-MM-dd' (to รวมวันนั้น)
  room_range: function(p){
    var e = carEmp_(p.caller); denyIf(!e, 'ไม่พบบัญชีพนักงาน');
    // ดูหลายวัน (ตารางห้อง สัปดาห์/เดือน) = เฉพาะแอดมิน/ผู้อนุมัติห้อง · วันเดียว = ทุกคน (ฟอร์มจองใช้เช็คเวลาว่าง)
    denyIf(p.to && p.to !== p.from && !rmIsAdmin_(e) && !rmIsApprover_(e), 'ตารางห้องดูได้เฉพาะแอดมินและผู้อนุมัติ');
    var from = rmDayRange_(p.from).from, to = rmDayRange_(p.to || p.from).to;
    var rooms = rmRooms_();
    var list = rmGet_('bookings?status=in.(tentative,confirmed)&start_at=lt.' + encodeURIComponent(to) + '&end_at=gt.' + encodeURIComponent(from) + '&select=*&order=start_at.asc');
    return { rooms:rooms, bookings:list.map(function(b){ var o = rmPub_(b, rooms); o.mine = String(b.requester_line) === String(p.caller); return o; }) };
  },
  room_create: function(p){
    var e = carEmp_(p.caller); denyIf(!e, 'ไม่พบบัญชีพนักงาน');
    var d = String(p.date||''), s = String(p.start||''), en = String(p.end||'');
    denyIf(!/^\d{4}-\d{2}-\d{2}$/.test(d) || !/^\d{2}:\d{2}$/.test(s) || !/^\d{2}:\d{2}$/.test(en), 'วันเวลาไม่ถูกต้อง');
    denyIf(en <= s, 'เวลาสิ้นสุดต้องหลังเวลาเริ่ม');
    denyIf(s < RM_OPEN || en > RM_CLOSE, 'จองได้ระหว่าง ' + RM_OPEN + '–' + RM_CLOSE);
    denyIf(new Date(rmIso_(d, s)).getTime() < Date.now() - 5 * 60000, 'เลือกเวลาที่ผ่านมาแล้วไม่ได้');
    var room = rmRooms_().filter(function(r){ return r.id === p.room_id; })[0]; denyIf(!room, 'ไม่พบห้อง');
    var hc = Math.max(1, Number(p.headcount) || 1);
    denyIf(room.capacity && hc > room.capacity, room.name + ' รับได้ ' + room.capacity + ' คน');
    denyIf(!RM_PURPOSE[p.purpose], 'กรุณาเลือกวัตถุประสงค์');
    denyIf(p.purpose === 'other' && !String(p.purpose_other||'').trim(), 'กรุณาระบุวัตถุประสงค์');
    denyIf(p.source !== 'todo' && !String(p.title||'').trim(), 'กรุณากรอกหัวข้อ');
    var n = function(v){ return Math.max(0, Math.min(500, Number(v) || 0)); };
    var ft = function(v, on){ return on ? (String(v || '').trim().slice(0, 100) || null) : null; };
    var bread = (n(p.food_break_am) && String(p.food_am_type||'').trim() === RM_BREAD) || (n(p.food_break_pm) && String(p.food_pm_type||'').trim() === RM_BREAD);
    denyIf(bread && rmDaysAhead_(d) < RM_BREAD_DAYS, 'ขนมปัง ต้องสั่งล่วงหน้า ' + RM_BREAD_DAYS + ' วัน — เลือก "น้ำ" หรือ "ระบุ" แทน');
    var lk = LockService.getScriptLock(); if(!lk.tryLock(20000)) throw 'ระบบกำลังบันทึกรายการอื่น ลองใหม่อีกครั้ง';
    var b;
    try{
      var no = rmNo_();
      var r = rmReq_('post', 'bookings', { room_id:room.id, start_at:rmIso_(d, s), end_at:rmIso_(d, en), status:'tentative',
        booking_no:no, title:String(p.title||'').trim(), purpose:p.purpose, purpose_other:String(p.purpose_other||'').trim(), headcount:hc,
        food_break_am:n(p.food_break_am), food_lunch:n(p.food_lunch), food_break_pm:n(p.food_break_pm), food_note:String(p.food_note||'').trim(),
        food_am_type:ft(p.food_am_type, n(p.food_break_am)), food_pm_type:ft(p.food_pm_type, n(p.food_break_pm)),
        food_lunch_type:ft(p.food_lunch_type, n(p.food_lunch) && p.purpose === 'guest'),   // แบบอาหารเที่ยง เฉพาะรับรองแขก
        requester_line:p.caller, requester_name:e.full_name||'', department:e.department||'', phone:String(p.phone || e.phone || ''),
        source:(p.source === 'todo' ? 'todo' : 'myssb'), meeting_id:(p.meeting_id || null) },
        'return=representation');
      // 23P01 = ชน constraint no_overlap (มีคนจองช่วงนี้ไปก่อนเสี้ยววินาที)
      if(r.code === 409 || (r.data && r.data.code === '23P01')) throw 'ช่วงเวลานี้มีคนจองไปแล้ว กรุณาเลือกเวลาใหม่';
      if(r.code >= 300) throw 'จองไม่สำเร็จ (' + r.code + ')';
      b = r.data[0];
    } finally { lk.releaseLock(); }
    // จองมาจากฟอร์มนัดประชุมของ To-do -> ผูกนัดกับการจองห้อง (meetings.room_booking_id)
    if(p.meeting_id){ try{ rmReq_('patch', 'meetings?id=eq.' + encodeURIComponent(p.meeting_id), { room_booking_id:b.id }); }catch(err){} }
    // นัดประชุมใน To-do (ถ้าติ๊กจากฝั่ง MySSB) — ผู้เข้าร่วมได้นัด + แจ้งเตือนจาก To-do
    var meeting = null;
    if(String(p.make_meeting) === 'true'){
      try{
        var me = rmGet_('users?line_user_id=eq.' + encodeURIComponent(p.caller) + '&select=id')[0];
        var mr = rmReq_('post', 'meetings', { title:String(p.title || RM_PURPOSE[p.purpose]).trim(), agenda:String(p.agenda||''),
          start_at:b.start_at, end_at:b.end_at, platform:'onsite', location:(room.site ? room.site + ' · ' : '') + room.name,
          organizer_id:me ? me.id : null, room_booking_id:b.id, reminder_minutes:Number(p.reminder) || 30 }, 'return=representation');
        if(mr.code < 300 && mr.data && mr.data[0]){
          meeting = mr.data[0];
          var att = (p.attendees || []).concat(me ? [me.id] : []).filter(function(v, i, a){ return v && a.indexOf(v) === i; });
          if(att.length) rmReq_('post', 'meeting_attendees', att.map(function(u){ return { meeting_id:meeting.id, user_id:u, status:(me && u === me.id) ? 'accepted' : 'invited' }; }));
          rmReq_('patch', 'bookings?id=eq.' + b.id, { meeting_id:meeting.id });
        }
      }catch(err){ meeting = { error:String(err) }; }
    }
    rmSyncSheet_([rmGetBooking_(b.id) || b]);
    // การ์ดให้แอดมินรับทราบ: ผู้อนุมัติห้อง + แอดมินบ้านกาแฟ (ไม่ซ้ำคน)
    var cto = rmApproverIds_(); rmAdminIds_().forEach(function(id){ if(cto.indexOf(id) < 0) cto.push(id); });
    try{ rmPush_(cto, b, 'ask', '🏢 จองห้อง ' + room.name + ' · ' + rmWhen_(b) + ' · รอยืนยัน'); }catch(err){}
    return { ok:true, id:b.id, booking_no:b.booking_no, meeting:!!(meeting && meeting.id) };
  },
  // scope: mine | queue (รออนุมัติ — ผู้อนุมัติ) | all (แอดมิน/ผู้อนุมัติ ย้อนหลัง/ล่วงหน้า)
  room_list: function(p){
    var e = carEmp_(p.caller); denyIf(!e, 'ไม่พบบัญชีพนักงาน');
    var sc = String(p.scope || 'mine'), q;
    if(sc === 'queue'){ denyIf(!rmCanConfirm_(e), 'เฉพาะแอดมินบ้านกาแฟ / ผู้ยืนยันการจอง'); q = 'bookings?status=eq.tentative&select=*&order=start_at.asc'; }
    else if(sc === 'all'){ denyIf(!rmIsApprover_(e) && !rmIsAdmin_(e)); q = 'bookings?select=*&order=start_at.desc&limit=500'; }
    else q = 'bookings?requester_line=eq.' + encodeURIComponent(p.caller) + '&select=*&order=start_at.desc&limit=100';
    var rooms = rmRooms_();
    return rmGet_(q).map(function(b){ var o = rmPub_(b, rooms); o.mine = String(b.requester_line) === String(p.caller); return o; });
  },
  room_get: function(p){
    var e = carEmp_(p.caller); denyIf(!e, 'ไม่พบบัญชีพนักงาน');
    var b = rmGetBooking_(p.id); if(!b) throw 'ไม่พบการจอง';
    var o = rmPub_(b, rmRooms_()); o.mine = String(b.requester_line) === String(p.caller);
    o.perm = { decide:rmCanConfirm_(e) && b.status === 'tentative', cancel:(o.mine || rmIsAdmin_(e)) && (b.status === 'tentative' || b.status === 'confirmed') && new Date(b.end_at) > new Date(),
      food_resend:rmIsAdmin_(e) && b.status === 'confirmed' && rmHasFood_(b) && new Date(b.end_at) > new Date() };
    if(o.perm.food_resend) o.food_to = rmFoodNames_();
    return o;
  },
  // แอดมินห้องส่งการ์ดสั่งอาหารถึงผู้ดูแลอาหารอีกครั้ง (เช่น เพิ่งติ๊กผู้สั่งอาหารหลังใบอนุมัติไปแล้ว)
  room_food_resend: function(p){
    var e = carEmp_(p.caller); denyIf(!rmIsAdmin_(e), 'เฉพาะแอดมินบ้านกาแฟ');
    var b = rmGetBooking_(p.id); if(!b) throw 'ไม่พบการจอง';
    denyIf(b.status !== 'confirmed' || !rmHasFood_(b), 'ใบนี้ไม่มีอาหาร หรือยังไม่ยืนยัน');
    var ids = rmFoodTargets_(); denyIf(!ids.length, 'ยังไม่มีผู้สั่งอาหาร — ติ๊ก "ผู้สั่งอาหารห้องประชุม" ในทะเบียนพนักงานก่อน');
    rmPush_(ids, b, 'food', '🍱 สั่งอาหารประชุม · ' + rmWhen_(b));
    var names = rmFoodNames_();
    return { ok:true, to:names.length ? names : ['แอดมินห้อง (ยังไม่ได้ติ๊กผู้สั่งอาหาร)'] };
  },
  room_decide: function(p){ return rmDecide_(p.id, p.caller, p.decision === 'approve' ? 'approve' : 'reject', p.reason); },
  room_cancel: function(p){
    var e = carEmp_(p.caller); denyIf(!e, 'ไม่พบบัญชีพนักงาน');
    var b = rmGetBooking_(p.id); if(!b) throw 'ไม่พบการจอง';
    denyIf(String(b.requester_line) !== String(p.caller) && !rmIsAdmin_(e), 'ยกเลิกได้เฉพาะผู้จองหรือแอดมิน');
    denyIf(b.status === 'cancelled', 'การจองนี้ยกเลิกไปแล้ว');
    var cb = rmPatch_(p.id, { status:'cancelled', decision:'cancelled' });
    rmSyncSheet_([cb]);
    if(b.status === 'confirmed' && rmHasFood_(b)) try{ rmPush_(rmFoodTargets_(), cb, 'foodx', '✕ ยกเลิกอาหาร · ' + rmWhen_(cb)); }catch(err){}
    return { ok:true };
  },
  // แอดมินบ้านกาแฟแก้คำอธิบายห้อง + อุปกรณ์ (แสดงใต้ชื่อห้องในหน้าจอง)
  room_set_info: function(p){
    var e = carEmp_(p.caller); denyIf(!rmIsAdmin_(e), 'เฉพาะแอดมินบ้านกาแฟ');
    denyIf(!p.id, 'ไม่พบห้อง');
    var fac = String(p.facilities || '').split(/[,\n]/).map(function(x){ return x.trim(); }).filter(Boolean).slice(0, 12);
    var r = rmReq_('patch', 'rooms?id=eq.' + encodeURIComponent(p.id), { description:String(p.description || '').trim().slice(0, 300) || null, facilities:fac }, 'return=representation');
    if(r.code >= 300 || !(r.data || []).length) throw 'บันทึกไม่สำเร็จ (' + r.code + ')';
    return { ok:true, room:r.data[0] };
  },
  // ผู้ใช้ To-do (ใช้เลือกผู้เข้าร่วมประชุม)
  room_people: function(p){
    var e = carEmp_(p.caller); denyIf(!e, 'ไม่พบบัญชีพนักงาน');
    return rmGet_('users?select=id,display_name&order=display_name.asc');
  },
  // คัดลอกการจองทั้งหมดลงชีต RoomBookings (ใช้ครั้งแรก / ซ่อมชีต) · เรียกซ้ำได้ ไม่ซ้ำแถว
  room_sync_all: function(){ return { synced:rmSyncSheet_(rmGet_('bookings?select=*&order=created_at.asc')) }; },
  room_diag: function(){
    var sp = PropertiesService.getScriptProperties();
    // ชื่อคนที่ได้รับแจ้งเตือน (ชื่อเท่านั้น ไม่คืน UID)
    var emps = getRows(SHEETS.EMP).filter(function(r){ return String(r.active) === 'true' && r.line_user_id; });
    var nameOf = function(ids){ return ids.map(function(id){ var x = emps.filter(function(r){ return String(r.line_user_id) === String(id); })[0]; return x ? x.full_name : '?'; }); };
    return { poll_last:sp.getProperty('POLL_LAST') || '', remind_day:sp.getProperty('RM_REMIND_DAY') || '', rooms:rmRooms_().length,
      room_admins:nameOf(rmAdminIds_()), room_approvers:nameOf(rmApproverIds_()), room_food:nameOf(rmFoodIds_()),
      using_ticks:{ admin:rmTicked_('room_admin').length > 0, approver:rmTicked_('room_approver').length > 0 } };
  }
});

// ---------- LINE ----------
var RM_RS = { busy:'ห้องมีภารกิจอื่น', move:'ขอเลื่อนวันหรือเวลา', other:'อื่น ๆ' };
function roomPostback_(d, uid, reply){
  if(!/^rm_/.test(String(d.act||''))) return false;
  var e = carEmp_(uid);
  if(!e || !rmCanConfirm_(e)){ lineReply(reply, '⛔ บัญชีนี้ไม่มีสิทธิ์ยืนยันการจองห้อง'); return true; }
  var b = rmGetBooking_(d.t); if(!b){ lineReply(reply, '❌ ไม่พบการจอง'); return true; }
  var no = b.booking_no || '';
  if(b.status !== 'tentative' && d.act !== 'rm_rs'){ lineReply(reply, 'ℹ️ การจอง ' + no + ' ดำเนินการไปแล้ว'); return true; }
  if(d.act === 'rm_ok'){ var r = rmDecide_(d.t, uid, 'approve'); lineReply(reply, (r.already ? 'ℹ️ ยืนยันไปแล้ว ' : '✅ ยืนยันรับทราบการจองห้อง ') + no + '\nแจ้งผู้จองเรียบร้อย'); return true; }
  if(d.act === 'rm_no'){
    var qi = function(label, key){ return { type:'action', action:{ type:'postback', label:label, data:pbData('rm_rs', d.t) + '&r=' + key, displayText:label } }; };
    lineReplyMsgs_(reply, [{ type:'text', text:'ไม่อนุมัติการจองห้อง ' + no + '\nเลือกเหตุผลด้านล่าง 👇',
      quickReply:{ items:[ qi('🏢 ห้องมีภารกิจอื่น', 'busy'), qi('🕘 ขอเลื่อนวัน/เวลา', 'move'), qi('✏️ อื่นๆ (พิมพ์เหตุผล)', 'other'), qi('ยกเลิก', 'cancel') ] } }]);
    return true;
  }
  if(d.act === 'rm_rs'){
    if(d.r === 'cancel'){ lineReply(reply, 'ยกเลิกแล้ว · การจอง ' + no + ' ยังรอพิจารณา'); return true; }
    if(d.r === 'other'){ CacheService.getScriptCache().put('rmrs_' + uid, d.t, 600); lineReply(reply, '✏️ พิมพ์เหตุผลที่ไม่อนุมัติการจอง ' + no + ' ส่งมาได้เลย (ภายใน 10 นาที)'); return true; }
    rmDecide_(d.t, uid, 'reject', RM_RS[d.r] || 'ไม่ระบุ');
    lineReply(reply, '✕ ไม่อนุมัติการจอง ' + no + '\nเหตุผล: ' + (RM_RS[d.r] || '-') + '\nแจ้งผู้จองเรียบร้อย');
    return true;
  }
  return false;
}
function roomReasonText_(ev){
  var uid = ev.source && ev.source.userId; if(!uid) return false;
  var c = CacheService.getScriptCache(), id = c.get('rmrs_' + uid); if(!id) return false;
  var txt = String(ev.message && ev.message.text || '').trim();
  if(!txt || /^เมนู:/.test(txt)) return false;
  c.remove('rmrs_' + uid);
  try{ rmDecide_(id, uid, 'reject', 'อื่น ๆ: ' + txt); lineReply(ev.replyToken, '✕ ไม่อนุมัติการจองห้อง\nเหตุผล: ' + txt + '\nแจ้งผู้จองเรียบร้อย'); }
  catch(err){ lineReply(ev.replyToken, 'ℹ️ ' + String(err)); }
  return true;
}
