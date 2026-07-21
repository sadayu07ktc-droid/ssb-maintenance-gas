/**
 * ssb-maintenance — ระบบล็อกอิน (อีเมล OTP + PIN 6 หลัก)  [เพิ่มเติมจาก Code.gs]
 *
 * ติดตั้ง:
 *  1) วางไฟล์นี้เพิ่มใน Apps Script (โปรเจคเดียวกับ Code.gs) → Save
 *  2) เพิ่ม action ใน API ของ Code.gs (ดูหมายเหตุท้ายไฟล์) — หรือใช้ registerAuthActions() ที่ผมผูกไว้ให้แล้ว
 *  3) รัน ensureAuthCols() 1 ครั้ง → เพิ่มคอลัมน์ email/pin/token ใน Employees + สร้างแท็บ AuthOTP
 *  4) Project Settings → Script Properties → เพิ่ม AUTH_PEPPER = <สตริงสุ่มยาวๆ ลับ>
 *  5) เติม "email" ของพนักงานแต่ละคนในแท็บ Employees (ส่วนตัว/บริษัท ก็ได้) + ใส่ role (admin/approver/requester)
 *  6) Deploy → Manage deployments → แก้ version เป็น New version
 */

var OTP_TTL_MIN   = 10;   // OTP ใช้ได้กี่นาที
var OTP_MAX_TRIES = 5;    // ใส่ OTP ผิดได้กี่ครั้ง
var OTP_RESEND_SEC= 60;   // ขอ OTP ซ้ำได้ทุกกี่วินาที
var TOKEN_TTL_DAY = 30;   // token จำเครื่องกี่วัน
var PIN_MAX_FAIL  = 5;    // PIN ผิดกี่ครั้งแล้วล็อก (ต้องยืนยันอีเมลใหม่)

function authPepper(){ return PropertiesService.getScriptProperties().getProperty('AUTH_PEPPER') || 'ssb-change-this-pepper'; }
function sha256(s){ return Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(s), Utilities.Charset.UTF_8)); }
function randToken(){ return sha256(uuid()+uuid()+now()+Math.random()).replace(/[^A-Za-z0-9]/g,'').slice(0,40); }
function gen6(){ return ('00000' + Math.floor(Math.random()*1000000)).slice(-6); }
function nowMs(){ return new Date().getTime(); }
function plusMs(ms){ return Utilities.formatDate(new Date(nowMs()+ms), TZ, 'yyyy-MM-dd HH:mm:ss'); }

// เพิ่มคอลัมน์ auth ใน Employees + แท็บ AuthOTP (รัน 1 ครั้ง / idempotent)
function ensureAuthCols(){
  var s = sh(SHEETS.EMP);
  if(!s){ s = ss().insertSheet(SHEETS.EMP); s.getRange(1,1,1,HEADERS.Employees.length).setValues([HEADERS.Employees]); }
  var head = s.getRange(1,1,1,Math.max(1,s.getLastColumn())).getValues()[0];
  // flow ปัจจุบัน (รหัสพนักงาน + วันเกิด + อนุมัติ) ใช้แค่ birthdate เพิ่ม
  ['birthdate'].forEach(function(c){
    if(head.indexOf(c) < 0){ s.getRange(1, s.getLastColumn()+1).setValue(c); head.push(c); }
  });
  return 'auth columns ready (birthdate)';
}

function empByEmail(email){
  email = String(email||'').trim().toLowerCase();
  if(!email) return null;
  return getRows(SHEETS.EMP).filter(function(r){ return String(r.email||'').trim().toLowerCase() === email; })[0] || null;
}
function empByToken(token){
  token = String(token||''); if(!token) return null;
  var e = getRows(SHEETS.EMP).filter(function(r){ return String(r.auth_token) === token; })[0];
  if(!e) return null;
  if(e.token_exp && new Date(e.token_exp).getTime() < nowMs()) return null;
  return e;
}
function patchEmp(row, patch){
  var s = sh(SHEETS.EMP);
  var head = s.getRange(1,1,1,s.getLastColumn()).getValues()[0];
  head.forEach(function(h,c){
    if(patch[h] === undefined) return;
    var cell = s.getRange(row, c+1);
    // เบอร์โทร/รหัส ต้องเก็บเป็นข้อความ ไม่งั้น Sheets กิน 0 ตัวหน้า (0834847856 -> 834847856)
    if(h === 'phone' || h === 'emp_code') cell.setNumberFormat('@');
    cell.setValue(patch[h]);
  });
}
function pubEmp(e){ return { id:e.id, full_name:e.full_name, role:(e.role||'requester'), email:e.email, department:e.department, emp_code:e.emp_code, has_pin: !!e.pin_hash }; }

function upsertOtp(email, obj){
  var s = sh('AuthOTP'); var vals = s.getDataRange().getValues(); var head = vals[0];
  var ec = head.indexOf('email');
  for(var i=1;i<vals.length;i++){
    if(String(vals[i][ec]).toLowerCase() === email){
      head.forEach(function(h,c){ if(obj[h] !== undefined) s.getRange(i+1, c+1).setValue(obj[h]); });
      return;
    }
  }
  appendObj('AuthOTP', obj);
}
function otpRow(email){
  var s = sh('AuthOTP'); var vals = s.getDataRange().getValues(); var head = vals[0];
  var ec = head.indexOf('email');
  for(var i=1;i<vals.length;i++){ if(String(vals[i][ec]).toLowerCase() === email){ var o={_row:i+1}; head.forEach(function(h,c){o[h]=vals[i][c];}); return o; } }
  return null;
}

// ---- ACTIONS ----

// 1) ขอ OTP ไปที่อีเมล (ครั้งแรก / เครื่องใหม่ / ลืม PIN)
function authRequestOtp(p){
  ensureAuthCols();
  var email = String(p.email||'').trim().toLowerCase();
  if(!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(email)) return { ok:false, error:'อีเมลไม่ถูกต้อง' };
  var emp = empByEmail(email);
  if(!emp) return { ok:false, error:'ไม่พบอีเมลนี้ในระบบพนักงาน — ติดต่อ HR เพื่อลงทะเบียนอีเมล' };
  if(String(emp.active) === 'false' || String(emp.active) === '0') return { ok:false, error:'บัญชีถูกปิดใช้งาน' };
  var prev = otpRow(email);
  if(prev && prev.last_sent && (nowMs() - new Date(prev.last_sent).getTime() < OTP_RESEND_SEC*1000))
    return { ok:false, error:'เพิ่งส่งรหัสไป รอสักครู่แล้วลองใหม่' };
  var code = gen6();
  upsertOtp(email, { email:email, code_hash: sha256(code + authPepper()), expires: plusMs(OTP_TTL_MIN*60000), tries:0, last_sent: now(), emp_id: emp.id });
  MailApp.sendEmail({
    to: email,
    subject: 'รหัสเข้าใช้งาน MySSB Connect: ' + code,
    body: 'สวัสดีคุณ ' + (emp.full_name||'') + '\n\nรหัส OTP สำหรับเข้าใช้งาน MySSB Connect คือ\n\n    ' + code + '\n\nรหัสนี้ใช้ได้ภายใน ' + OTP_TTL_MIN + ' นาที · กรุณาอย่าเปิดเผยรหัสนี้กับผู้อื่น\n\n— ระบบ MySSB Connect (สุขสมบูรณ์ กรุ๊ป)'
  });
  return { ok:true, sent:true, name: emp.full_name, masked: email.replace(/^(.).*(@.*)$/, '$1***$2') };
}

// 2) ยืนยัน OTP → ได้ token (เก็บในเครื่อง) + บอกว่ามี PIN แล้วหรือยัง
function authVerifyOtp(p){
  var email = String(p.email||'').trim().toLowerCase();
  var code  = String(p.otp||'').trim();
  var rec = otpRow(email);
  if(!rec) return { ok:false, error:'ยังไม่ได้ขอรหัส หรือรหัสถูกใช้ไปแล้ว' };
  if(nowMs() > new Date(rec.expires).getTime()) return { ok:false, error:'รหัสหมดอายุ กรุณาขอใหม่' };
  if(Number(rec.tries) >= OTP_MAX_TRIES) return { ok:false, error:'ใส่รหัสผิดหลายครั้ง กรุณาขอรหัสใหม่' };
  if(sha256(code + authPepper()) !== rec.code_hash){
    sh('AuthOTP').getRange(rec._row, 4).setValue(Number(rec.tries)+1); // tries col = 4
    return { ok:false, error:'รหัสไม่ถูกต้อง' };
  }
  var emp = getRows(SHEETS.EMP).filter(function(r){ return String(r.id) === String(rec.emp_id); })[0];
  var token = randToken();
  var patch = { auth_token: token, token_exp: plusMs(TOKEN_TTL_DAY*86400000), pin_fail: 0 };
  if(p.line_user_id) patch.line_user_id = String(p.line_user_id); // ผูก LINE เข้ากับพนักงาน (ลงทะเบียน)
  patchEmp(emp._row, patch);
  sh('AuthOTP').deleteRow(rec._row);
  if(p.line_user_id) emp.line_user_id = String(p.line_user_id);
  return { ok:true, token: token, employee: pubEmp(emp) };
}

// 3) ตั้ง PIN 6 หลัก (หลังยืนยัน OTP)
function authSetPin(p){
  var emp = empByToken(p.token);
  if(!emp) return { ok:false, error:'need_login' };
  var pin = String(p.pin||'');
  if(!/^\d{6}$/.test(pin)) return { ok:false, error:'PIN ต้องเป็นตัวเลข 6 หลัก' };
  var salt = uuid();
  patchEmp(emp._row, { pin_salt: salt, pin_hash: sha256(salt + pin + authPepper()), pin_fail: 0 });
  return { ok:true, employee: pubEmp(emp) };
}

// 4) เข้าด้วย token ของเครื่อง (เช็คว่าใคร+ มี PIN ไหม) — ยังไม่ปล่อยเข้าถ้ายังไม่ใส่ PIN
function authLoginToken(p){
  var emp = empByToken(p.token);
  if(!emp) return { ok:false, error:'need_login' };
  return { ok:true, employee: pubEmp(emp) };
}

// 5) ปลดล็อกด้วย PIN (เครื่องที่จำ token ไว้)
function authLoginPin(p){
  var emp = empByToken(p.token);
  if(!emp) return { ok:false, error:'need_login' };
  if(!emp.pin_hash) return { ok:false, error:'ยังไม่ได้ตั้ง PIN' };
  if(Number(emp.pin_fail||0) >= PIN_MAX_FAIL){
    patchEmp(emp._row, { auth_token:'', token_exp:'' }); // ล้าง token → บังคับยืนยันอีเมลใหม่
    return { ok:false, error:'need_login', locked:true };
  }
  if(sha256(emp.pin_salt + String(p.pin||'') + authPepper()) !== emp.pin_hash){
    patchEmp(emp._row, { pin_fail: Number(emp.pin_fail||0)+1 });
    return { ok:false, error:'PIN ไม่ถูกต้อง', left: PIN_MAX_FAIL - (Number(emp.pin_fail||0)+1) };
  }
  patchEmp(emp._row, { pin_fail: 0 });
  return { ok:true, employee: pubEmp(emp) };
}

// ============================================================
//  ลงทะเบียนด้วย รหัสพนักงาน + วันเกิด + อนุมัติโดยแอดมิน (แนะนำ)
// ============================================================
function normDate(s){
  s = String(s||'').trim(); if(!s) return '';
  var m;
  if(m = s.match(/^(\d{4})[-\/.](\d{1,2})[-\/.](\d{1,2})$/)){ var y=+m[1]; if(y>2400)y-=543; return y+'-'+('0'+m[2]).slice(-2)+'-'+('0'+m[3]).slice(-2); }
  if(m = s.match(/^(\d{1,2})[-\/.](\d{1,2})[-\/.](\d{4})$/)){ var y2=+m[3]; if(y2>2400)y2-=543; return y2+'-'+('0'+m[2]).slice(-2)+'-'+('0'+m[1]).slice(-2); }
  return s;
}
function empByCode(code){
  code = String(code||'').trim();
  if(!code) return null;
  return getRows(SHEETS.EMP).filter(function(r){ return String(r.emp_code||'').trim() === code; })[0] || null;
}

// (1) เช็ครหัสพนักงาน → คืนชื่อ+แผนก ให้ยืนยัน
function regLookup(p){
  ensureAuthCols();
  var emp = empByCode(p.emp_code);
  if(!emp) return { ok:false, error:'ไม่พบรหัสพนักงานนี้ในระบบ' };
  if(emp.line_user_id && String(emp.active) === 'true') return { ok:false, error:'รหัสนี้ลงทะเบียนและใช้งานแล้ว — ติดต่อ HR หากไม่ใช่คุณ' };
  return { ok:true, full_name: emp.full_name, department: emp.department||'' };
}

// (2) ยืนยันวันเกิด → ผูก LINE + ตั้งสถานะรออนุมัติ (active='pending')
function regSubmit(p){
  ensureAuthCols();
  var emp = empByCode(p.emp_code);
  if(!emp) return { ok:false, error:'ไม่พบรหัสพนักงาน' };
  if(emp.line_user_id && String(emp.active) === 'true') return { ok:false, error:'รหัสนี้ลงทะเบียนแล้ว' };
  if(!emp.birthdate) return { ok:false, error:'ระบบยังไม่มีวันเกิดของพนักงานนี้ — ติดต่อ HR' };
  if(normDate(p.birthdate) !== normDate(emp.birthdate)) return { ok:false, error:'วันเดือนปีเกิดไม่ตรงกับข้อมูล' };
  var isAdmin = String(emp.role||'').toLowerCase().indexOf('admin') >= 0; // แอดมิน = อนุมัติอัตโนมัติ (bootstrap)
  var patch = { line_user_id: String(p.line_user_id||''), active: isAdmin ? 'true' : 'pending' };
  if(!emp.id) patch.id = uuid(); // เจน id อัตโนมัติถ้ายังว่าง
  if(p.phone) patch.phone = String(p.phone); // เบอร์โทร (กรอกตอนลงทะเบียน)
  if(p.email) patch.email = String(p.email); // อีเมล (ไม่บังคับ)
  patchEmp(emp._row, patch);
  if(!isAdmin){ try{ notifyAdmins('🔔 มีคำขอลงทะเบียนใหม่: ' + emp.full_name + ' (' + emp.emp_code + ')' + (emp.department?(' · '+emp.department):'')); }catch(e){} }
  return { ok:true, active: isAdmin, pending: !isAdmin, full_name: emp.full_name };
}

// (3) แอดมิน: รายชื่อรออนุมัติ
function pendingUsers(){
  return getRows(SHEETS.EMP).filter(function(r){ return String(r.active) === 'pending'; })
    .map(function(e){ return { id:e.id, emp_code:e.emp_code, full_name:e.full_name, department:e.department }; });
}
// (4) แอดมิน: อนุมัติ / ไม่อนุมัติ
function approveUser(p){
  var e = getRows(SHEETS.EMP).filter(function(r){ return String(r.id) === String(p.emp_id); })[0];
  if(!e) return { ok:false, error:'ไม่พบ' };
  patchEmp(e._row, { active:'true' });
  return { ok:true };
}
function rejectUser(p){
  var e = getRows(SHEETS.EMP).filter(function(r){ return String(r.id) === String(p.emp_id); })[0];
  if(!e) return { ok:false, error:'ไม่พบ' };
  patchEmp(e._row, { active:'false', line_user_id:'' });
  return { ok:true };
}

// ผูก action เข้ากับ API ของ Code.gs (เรียกตอนโหลดสคริปต์)
(function registerAuthActions(){
  if(typeof API === 'undefined') return;
  API.request_otp  = function(p){ return authRequestOtp(p); };
  API.verify_otp   = function(p){ return authVerifyOtp(p); };
  API.set_pin      = function(p){ return authSetPin(p); };
  API.login_token  = function(p){ return authLoginToken(p); };
  API.login_pin    = function(p){ return authLoginPin(p); };
  API.reg_lookup   = function(p){ return regLookup(p); };
  API.reg_submit   = function(p){ return regSubmit(p); };
  API.pending_users= function(p){ return pendingUsers(p); };
  API.approve_user = function(p){ return approveUser(p); };
  API.reject_user  = function(p){ return rejectUser(p); };
})();
