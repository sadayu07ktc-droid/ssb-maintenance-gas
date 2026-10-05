// ============================================================
//  ทะเบียนพนักงาน (แอดมิน) — แก้ชีต Employees ผ่านแอป ไม่ต้องเปิดชีตเอง
//  ⚠️ ชีตนี้เป็น "ตาราง" (คอลัมน์มีประเภท) -> ห้าม setNumberFormat เด็ดขาด (พังตอน flush ดักไม่ได้)
//  ⚠️ ไม่ส่ง line_user_id / pin / token / signature ออกไปหน้าเว็บ (ส่งแค่ว่าผูก LINE แล้วหรือยัง)
//  อัปโหลดไฟล์ = เพิ่ม/แก้เท่านั้น ไม่ลบใคร · ช่องว่างในไฟล์ = คงค่าเดิม (กันมือไปโดนลบแล้วข้อมูลหาย)
// ============================================================
var EMP_F = [
  ['emp_code','รหัสพนักงาน'], ['full_name','ชื่อ-นามสกุล'], ['nickname','ชื่อเล่น'], ['department','แผนก'],
  ['position','ตำแหน่ง'], ['company','บริษัท'], ['phone','เบอร์โทร'], ['email','อีเมล'],
  ['birthdate','วันเกิด'], ['start_date','วันเริ่มงาน'], ['role','สิทธิ์'], ['active','สถานะ'], ['dept_approver','หัวหน้าแผนก'],
  ['room_admin','แอดมินห้องประชุม'], ['room_approver','ผู้อนุมัติห้องประชุม'], ['room_food','ผู้สั่งอาหารห้องประชุม']
];
// คอลัมน์กล่องติ๊ก (ใช่/ไม่ใช่)
var EMP_BOOL = { dept_approver:1, room_admin:1, room_approver:1, room_food:1 };
var EMP_LABEL = EMP_F.reduce(function(o, f){ o[f[0]] = f[1]; return o; }, {});

function empAdminOnly_(caller){ denyIf(!isAdminLine(caller), 'เฉพาะแอดมิน'); }
function empLocked_(fn){
  var lk = LockService.getScriptLock();
  if(!lk.tryLock(20000)) throw 'ระบบกำลังบันทึกรายการอื่น ลองใหม่อีกครั้ง';
  try{ return fn(); } finally { lk.releaseLock(); }
}
/**
 * หัวคอลัมน์แรกต้องชื่อ id (ตอนแปลงชีตเป็นตาราง ชื่อหัวถูกเปลี่ยนเป็น "คอลัมน์ 1" -> r.id หายทั้งระบบ)
 * และทุกคนต้องมี id (เดิมมีแค่คนที่ลงทะเบียนแล้ว) ไม่งั้นเลือกคนขับ/แก้ไขทีละคนไม่ได้
 */
function empEnsureIds_(){
  var s = sh(SHEETS.EMP), head = s.getRange(1,1,1,s.getLastColumn()).getValues()[0];
  var fixed = { header:false, ids:0 };
  if(head.indexOf('id') < 0 && /^(คอลัมน์|column)\s*1$/i.test(String(head[0]).trim())){ s.getRange(1,1).setValue('id'); head[0] = 'id'; fixed.header = true; }
  var ci = head.indexOf('id'); if(ci < 0) return fixed;
  var last = s.getLastRow(); if(last < 2) return fixed;
  var vals = s.getRange(2, 1, last-1, head.length).getValues();
  vals.forEach(function(r, i){
    var other = r.filter(function(v, j){ return j !== ci; });
    if(!empBlank_(other) && String(r[ci]).trim() === ''){ s.getRange(i+2, ci+1).setValue(uuid()); fixed.ids++; }
  });
  return fixed;
}
/**
 * สร้างคอลัมน์กล่องติ๊กที่ยังไม่มี (room_admin / room_approver ฯลฯ) — ทำครั้งเดียวตอนหัวคอลัมน์ยังไม่มี
 * ใส่ FALSE เฉพาะแถวที่มีพนักงาน (แถวว่างต้องว่าง ไม่งั้น getRows นับเป็นพนักงาน) · ห้าม setNumberFormat
 */
function empEnsureBoolCols_(){
  var s = sh(SHEETS.EMP), made = [];
  Object.keys(EMP_BOOL).forEach(function(col){
    var head = s.getRange(1,1,1,s.getLastColumn()).getValues()[0];
    if(head.indexOf(col) >= 0) return;
    s.getRange(1, s.getLastColumn()+1).setValue(col);
    var ci = s.getLastColumn(), last = s.getLastRow();
    s.getRange(2, ci, Math.max(s.getMaxRows()-1, 1), 1).setDataValidation(SpreadsheetApp.newDataValidation().requireCheckbox().build());
    if(last > 1){
      var all = s.getRange(2, 1, last-1, ci).getValues();
      s.getRange(2, ci, last-1, 1).setValues(all.map(function(r){ return [empBlank_(r.slice(0, ci-1)) ? '' : false]; }));
    }
    made.push(col);
  });
  return made;
}
// แถวว่าง: ทุกช่องว่างหรือ FALSE (กล่องติ๊กในตารางว่าง = FALSE)
function empBlank_(r){ return r.every(function(v){ return v === '' || v === false; }); }
/**
 * ล้าง FALSE ที่ติดอยู่ในแถวว่าง (คอลัมน์กล่องติ๊ก) -> ช่องว่าง
 * เคยเขียน FALSE ลงแถวว่างเกือบพันแถวตอนเพิ่มคอลัมน์ room_admin/room_approver (30 ก.ย.) เรียกซ้ำได้
 */
function empCleanBlankRows_(){
  var s = sh(SHEETS.EMP), head = s.getRange(1,1,1,s.getLastColumn()).getValues()[0], last = s.getLastRow();
  if(last < 2) return 0;
  var bools = head.map(function(h, i){ return EMP_BOOL[h] ? i : -1; }).filter(function(i){ return i >= 0; });
  var vals = s.getRange(2, 1, last-1, head.length).getValues(), n = 0;
  bools.forEach(function(ci){
    var dirty = false;
    var col = vals.map(function(r){
      var other = r.filter(function(v, j){ return bools.indexOf(j) < 0; });
      if(empBlank_(other) && r[ci] !== ''){ dirty = true; n++; return ['']; }
      return [r[ci]];
    });
    if(dirty) s.getRange(2, ci+1, col.length, 1).setValues(col);
  });
  return n;
}
function empHead_(){ var s = sh(SHEETS.EMP); return s.getRange(1,1,1,s.getLastColumn()).getValues()[0]; }

// ---------- แปลงค่าให้เทียบกันได้ ----------
function empRoleKey_(v){ v = String(v||'').toLowerCase(); if(/admin/.test(v)) return 'admin'; if(/approv|exec|manager|บริหาร|ผู้อนุมัติ/.test(v)) return 'approver'; return v ? 'requester' : ''; }
function empActiveKey_(v){
  if(v === true) return 'true'; if(v === false) return 'false';
  v = String(v||'').trim().toLowerCase();
  if(!v) return '';
  if(/^(true|ใช้งาน|active|yes|1)$/.test(v)) return 'true';
  if(/^(false|ปิด|ปิดการใช้งาน|inactive|no|0|ลาออก)$/.test(v)) return 'false';
  if(/^(pending|รออนุมัติ|รอ)$/.test(v)) return 'pending';
  return '?';
}
function empBoolKey_(v){
  if(v === true) return 'true'; if(v === false) return 'false';
  v = String(v||'').trim().toLowerCase();
  if(!v) return '';
  if(/^(true|ใช่|yes|y|1|✓|x|หัวหน้า)$/.test(v)) return 'true';
  if(/^(false|ไม่ใช่|no|n|0|-)$/.test(v)) return 'false';
  return '?';
}
function empDateKey_(v){ if(v instanceof Date) return Utilities.formatDate(v, TZ, 'yyyy-MM-dd'); return normDate(String(v||'').trim()); }
function empNorm_(k, v){
  if(k === 'role') return empRoleKey_(v);
  if(k === 'active') return empActiveKey_(v);
  if(EMP_BOOL[k]) return empBoolKey_(v);
  if(k === 'birthdate' || k === 'start_date') return empDateKey_(v);
  var s = String(v == null ? '' : v).trim();
  if(k === 'phone' || k === 'emp_code') s = s.replace(/^'/, '').replace(/\s+/g, '');
  return s;
}
function empShow_(k, key){
  if(k === 'role') return { admin:'Admin', approver:'Approver', requester:'User' }[key] || key;
  if(k === 'active') return { 'true':'ใช้งาน', 'false':'ปิดการใช้งาน', pending:'รออนุมัติ' }[key] || key;
  if(EMP_BOOL[k]) return key === 'true' ? 'ใช่' : 'ไม่ใช่';
  return key;
}
// ค่าที่เขียนลงชีต: สิทธิ์ใช้ตัวสะกดเดียวกับที่มีอยู่ในชีต (Admin/Approver) · เลขนำหน้า 0 ต้องมี ' ไม่งั้นชีตตัด 0 ทิ้ง
function empRoleWrite_(key, rows){
  var hit = rows.map(function(r){ return String(r.role||''); }).filter(function(v){ return v && empRoleKey_(v) === key; })[0];
  return hit || { admin:'Admin', approver:'Approver', requester:'requester' }[key];
}
function empWriteVal_(k, key, rows){
  if(k === 'role') return empRoleWrite_(key, rows);
  if(k === 'active') return key;                                   // 'true'/'false'/'pending' แบบเดียวกับโค้ดลงทะเบียนเดิม
  if(EMP_BOOL[k]) return key === 'true';
  if((k === 'phone' || k === 'emp_code') && /^0\d+$/.test(key)) return "'" + key;
  return key;
}
function empPublic_(r){
  var o = { id:String(r.id||''), linked:!!String(r.line_user_id||'').trim(), has_sign:!!String(r.signature_file_id||'').trim() };
  EMP_F.forEach(function(f){ o[f[0]] = empShow_(f[0], empNorm_(f[0], r[f[0]])); });
  o.role_key = empRoleKey_(r.role); o.active_key = empActiveKey_(r.active); o.head = empBoolKey_(r.dept_approver) === 'true';
  o.room_admin_on = empBoolKey_(r.room_admin) === 'true'; o.room_approver_on = empBoolKey_(r.room_approver) === 'true';
  o.room_food_on = empBoolKey_(r.room_food) === 'true';
  return o;
}

// ---------- เขียนแถว ----------
function empWriteRow_(row, patch){
  var s = sh(SHEETS.EMP), head = empHead_();
  head.forEach(function(h, c){ if(h && patch[h] !== undefined) s.getRange(row, c+1).setValue(patch[h]); });
}
// คนใหม่ลงแถวว่างแถวแรก (ชีตเคยมีแถวว่างคั่น) ไม่มีค่อยต่อท้าย
function empNewRow_(){
  var s = sh(SHEETS.EMP), head = empHead_(), last = s.getLastRow();
  if(last >= 2){
    var vals = s.getRange(2, 1, last-1, head.length).getValues();
    for(var i=0;i<vals.length;i++){ if(vals[i].join('') === '' || vals[i].every(function(v){ return v === '' || v === false; })) return i + 2; }
  }
  return last + 1;
}

// ---------- ตรวจแถว (ใช้ทั้งฟอร์มเว็บและไฟล์อัปโหลด) ----------
/**
 * rows: [{ _line, id, emp_code, ... }] · คืน { add, change, same, errors, missing, plan }
 * plan = รายการที่จะเขียนจริง (ใช้ตอน apply เท่านั้น ไม่ส่งออกหน้าเว็บ)
 */
function empDiff_(rows, fromFile){
  var cur = getRows(SHEETS.EMP).filter(function(r){ return String(r.full_name||'').trim() || String(r.emp_code||'').trim(); });
  var byId = {}, byCode = {};
  cur.forEach(function(r){ if(String(r.id||'').trim()) byId[String(r.id).trim()] = r; var c = empNorm_('emp_code', r.emp_code); if(c) byCode[c] = r; });
  var out = { add:[], change:[], same:0, errors:[], missing:[], blank:0, plan:[] }, seen = {}, touched = {};
  (rows || []).forEach(function(r, i){
    var line = r._line || (i + 1);
    var hasAny = EMP_F.some(function(f){ return String(r[f[0]] == null ? '' : r[f[0]]).trim() !== ''; });
    if(!hasAny){ out.blank++; return; }
    var id = String(r.id||'').trim(), code = empNorm_('emp_code', r.emp_code);
    var t = (id && byId[id]) || null;
    if(!t && code && byCode[code]) t = byCode[code];
    if(id && !t){ out.errors.push({ line:line, msg:'ไม่พบ id นี้ในระบบ (ห้ามแก้คอลัมน์ id)', name:String(r.full_name||'') }); return; }
    if(code){
      if(seen[code]){ out.errors.push({ line:line, msg:'รหัสพนักงาน ' + code + ' ซ้ำกับแถว ' + seen[code], name:String(r.full_name||'') }); return; }
      seen[code] = line;
    }
    var errs = [], patch = {}, diffs = [];
    EMP_F.forEach(function(f){
      var k = f[0]; if(!(k in r)) return;
      var raw = r[k];
      if(String(raw == null ? '' : raw).trim() === ''){
        // ไฟล์: ช่องว่าง = คงค่าเดิม (กันมือไปโดนลบ) · ฟอร์ม: ลบค่าได้ ยกเว้นช่องบังคับ
        if(fromFile) return;
        if(k === 'emp_code' || k === 'full_name'){ errs.push('ต้องกรอก' + f[1]); return; }
        if(t && empNorm_(k, t[k]) !== ''){ patch[k] = ''; diffs.push({ k:k, label:f[1], old:empShow_(k, empNorm_(k, t[k])), 'new':'(ว่าง)' }); }
        return;
      }
      var nv = empNorm_(k, raw);
      if(nv === '?'){ errs.push(f[1] + ' "' + raw + '" ไม่ถูกต้อง'); return; }
      if((k === 'birthdate' || k === 'start_date') && !/^\d{4}-\d{2}-\d{2}$/.test(nv)){ errs.push(f[1] + ' "' + raw + '" ไม่ใช่วันที่ (ใช้ ปปปป-ดด-วว)'); return; }
      if(k === 'email' && !/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(nv)){ errs.push('อีเมล "' + raw + '" ไม่ถูกต้อง'); return; }
      var ov = t ? empNorm_(k, t[k]) : '';
      if(nv !== ov){ patch[k] = empWriteVal_(k, nv, cur); diffs.push({ k:k, label:f[1], old:empShow_(k, ov), 'new':empShow_(k, nv) }); }
    });
    if(!t){
      if(!code) errs.push('ไม่มีรหัสพนักงาน');
      if(!String(r.full_name||'').trim()) errs.push('ไม่มีชื่อ-นามสกุล');
    }
    if(t && patch.emp_code !== undefined && byCode[code] && byCode[code] !== t) errs.push('รหัสพนักงาน ' + code + ' เป็นของ ' + byCode[code].full_name + ' อยู่แล้ว');
    if(errs.length){ out.errors.push({ line:line, msg:errs.join(' · '), name:String(r.full_name || (t && t.full_name) || '') }); return; }
    if(t) touched[t._row] = 1;
    var name = String(r.full_name || (t && t.full_name) || '').trim(), dept = String(r.department || (t && t.department) || '').trim();
    if(!t){
      if(patch.active === undefined) patch.active = 'true';
      if(patch.role === undefined) patch.role = empRoleWrite_('requester', cur);
      Object.keys(EMP_BOOL).forEach(function(k){ if(patch[k] === undefined) patch[k] = false; });
      out.add.push({ line:line, name:name, dept:dept }); out.plan.push({ row:0, patch:patch });
    } else if(diffs.length){
      out.change.push({ line:line, name:name, dept:dept, fields:diffs }); out.plan.push({ row:t._row, patch:patch });
    } else out.same++;
  });
  if(fromFile){
    out.missing = cur.filter(function(r){ return !touched[r._row]; }).map(function(r){ return String(r.full_name || r.emp_code || '-'); });
  }
  return out;
}
function empApplyPlan_(plan){
  var n = 0;
  plan.forEach(function(p){
    if(p.row){ empWriteRow_(p.row, p.patch); n++; return; }
    var row = empNewRow_(); p.patch.id = uuid(); empWriteRow_(row, p.patch); n++;
  });
  return n;
}

// ---------- API ----------
Object.assign(API, {
  // ซ่อมหัวคอลัมน์ id + ใส่ id ให้ทุกคน (เรียกซ้ำได้ ไม่ทำอะไรถ้าครบแล้ว)
  // ล้างแถวว่างก่อนเสมอ ไม่งั้นแถวว่างที่มี FALSE จะถูกใส่ id กลายเป็น "พนักงาน"
  emp_fix_ids: function(){ return empLocked_(function(){ var c = empCleanBlankRows_(); var r = empEnsureIds_(); r.cols = empEnsureBoolCols_(); r.cleaned = c; return r; }); },

  emp_admin_list: function(p){
    empAdminOnly_(p.caller);
    empLocked_(function(){ empCleanBlankRows_(); empEnsureIds_(); empEnsureBoolCols_(); });
    var rows = getRows(SHEETS.EMP).filter(function(r){ return String(r.full_name||'').trim() || String(r.emp_code||'').trim(); });
    var distinct = function(k){ var o = {}; rows.forEach(function(r){ var v = String(r[k]||'').trim(); if(v) o[v] = 1; }); return Object.keys(o).sort(); };
    return { rows:rows.map(function(r){ var o = empPublic_(r); o.self = String(r.line_user_id||'') === String(p.caller); return o; }), options:{ department:distinct('department'), company:distinct('company'), position:distinct('position') } };
  },

  // บันทึกจากฟอร์ม (เพิ่มใหม่เมื่อไม่มี id) — ใช้ตัวตรวจชุดเดียวกับไฟล์อัปโหลด
  emp_admin_save: function(p){
    empAdminOnly_(p.caller);
    var d = p.data || {}; d._line = 1; d.id = String(p.id || '');
    return empLocked_(function(){
      var r = empDiff_([d], false);
      if(r.errors.length) throw r.errors[0].msg;
      var n = empApplyPlan_(r.plan);
      return { ok:true, saved:n, added:r.add.length, changed:r.change.length ? r.change[0].fields.length : 0 };
    });
  },

  // ยกเลิกการผูก LINE (เปลี่ยนเครื่อง/เปลี่ยนไลน์) -> พนักงานลงทะเบียนใหม่ได้
  emp_admin_unlink: function(p){
    empAdminOnly_(p.caller);
    return empLocked_(function(){
      var t = getRows(SHEETS.EMP).filter(function(r){ return String(r.id) === String(p.id); })[0];
      if(!t) throw 'ไม่พบพนักงาน';
      denyIf(String(t.line_user_id) === String(p.caller), 'ยกเลิกการผูก LINE ของตัวเองไม่ได้');
      var patch = { line_user_id:'' }; if(empHead_().indexOf('auth_token') >= 0) patch.auth_token = '';
      empWriteRow_(t._row, patch);
      return { ok:true };
    });
  },

  // ลบพนักงานออกจากทะเบียน — เก็บสำเนาทั้งแถวไว้ในแผ่น EmployeesDeleted ก่อนลบ (กู้คืนได้ถ้าลบผิดคน)
  // ใบแจ้งซ่อม/ใบจองรถเดิมยังอยู่ครบ เพราะเก็บชื่อผู้ขอไว้ในใบแล้ว
  emp_admin_delete: function(p){
    empAdminOnly_(p.caller);
    return empLocked_(function(){
      var t = getRows(SHEETS.EMP).filter(function(r){ return String(r.id) === String(p.id) && String(p.id).trim(); })[0];
      if(!t) throw 'ไม่พบพนักงาน';
      denyIf(String(t.line_user_id) === String(p.caller), 'ลบบัญชีของตัวเองไม่ได้');
      var s = sh(SHEETS.EMP), head = empHead_();
      var vals = s.getRange(t._row, 1, 1, head.length).getValues()[0];
      var bk = sh('EmployeesDeleted');
      if(!bk){ bk = ss().insertSheet('EmployeesDeleted'); bk.getRange(1,1,1,head.length+2).setValues([['deleted_at','deleted_by'].concat(head)]); }
      var me = getRows(SHEETS.EMP).filter(function(r){ return String(r.line_user_id) === String(p.caller); })[0] || {};
      bk.appendRow([now(), String(me.full_name || '')].concat(vals));
      s.deleteRow(t._row);
      return { ok:true, name:String(t.full_name || '') };
    });
  },

  emp_import_preview: function(p){
    empAdminOnly_(p.caller);
    var r = empDiff_(p.rows || [], true); delete r.plan;
    return r;
  },
  emp_import_apply: function(p){
    empAdminOnly_(p.caller);
    return empLocked_(function(){
      empEnsureIds_();
      var r = empDiff_(p.rows || [], true);   // ตรวจซ้ำฝั่งเซิร์ฟเวอร์ ไม่เชื่อผลจากหน้าเว็บ
      var n = empApplyPlan_(r.plan);
      return { ok:true, applied:n, added:r.add.length, changed:r.change.length, skipped:r.errors.length };
    });
  }
});
