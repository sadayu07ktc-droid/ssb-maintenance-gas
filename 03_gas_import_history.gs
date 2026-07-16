/**
 * สร้างแท็บ MaintenanceRecords จากประวัติดิบ (xlsx ต้นฉบับ)
 * ไม่ต้องขนข้อมูลผ่านที่อื่น — GAS แตกคอลัมน์ + map vehicle_key ให้เอง
 *
 * ใช้งาน:
 *  1) ในชีต: File → Import → Upload → "สรุปค่าใช้จ่าย+ทะเบียนรถ.xlsx"
 *       เลือก "Insert new sheet(s)" → Import  (จะได้แท็บใหม่ 2 อัน)
 *  2) หาแท็บที่มีคอลัมน์ "อาการและรายละเอียดที่ซ่อม" → เปลี่ยนชื่อเป็น  RawHistory
 *       (แท็บทะเบียนรถที่ import ซ้ำมา ลบทิ้งได้)
 *  3) วางไฟล์นี้เพิ่มใน Apps Script (โปรเจคเดียวกับ Code.gs) → Save
 *  4) เลือกฟังก์ชัน buildMaintenanceFromRaw → Run → อนุญาตสิทธิ์
 *  5) เสร็จแล้วลบแท็บ RawHistory ได้ (ถ้าต้องการ)
 */
function buildMaintenanceFromRaw(){
  var raw = getRows('RawHistory');
  if(!raw.length) throw 'ไม่พบแท็บ RawHistory — import xlsx แล้วเปลี่ยนชื่อแท็บก่อน';

  // ── map ป้าย(ตัดช่องว่าง) -> vehicle_key จากแท็บ Vehicles ──
  var pmap = {};
  getRows(SHEETS.VEH).forEach(function(v){
    String(v.plate_all || '').split(';').forEach(function(p){
      p = p.replace(/\s/g,''); if(p) pmap[p] = v.vehicle_key;
    });
    [v.plate_current, v.plate_red].forEach(function(p){
      p = String(p||'').replace(/\s/g,''); if(p && !pmap[p]) pmap[p] = v.vehicle_key;
    });
  });

  function pick(o, names){ for(var i=0;i<names.length;i++){ if(o[names[i]] !== undefined && o[names[i]] !== '') return o[names[i]]; } return ''; }
  function toDate(v){
    if(v instanceof Date) return Utilities.formatDate(v, TZ, 'yyyy-MM-dd');
    if(String(v).match(/^\d+(\.\d+)?$/)){ var ms = Date.UTC(1899,11,30) + Math.round(Number(v))*86400000; return Utilities.formatDate(new Date(ms), TZ, 'yyyy-MM-dd'); }
    return String(v||'');
  }

  var HH = ['id','vehicle_key','ทะเบียนรถ','plate_type','วันที่ซ่อม','เลขที่ใบแจ้งซ่อม','ระยะเช็ค_กม','เลขไมล์','รายการซ่อม','จำนวนเงิน'];
  var rows = [];
  raw.forEach(function(r){
    var plate = String(pick(r,['ทะเบียนรถ','ทะเบียน'])).trim();
    var g     = String(pick(r,['อาการและรายละเอียดที่ซ่อม','รายการซ่อม','อาการ']));
    var money = pick(r,['จำนวนเงิน']);
    var refno = String(pick(r,['เลขที่ใบแจ้งซ่อม'])).trim();
    var dstr  = toDate(pick(r,['วัน/เดือน/ปีที่ซ่อม','วันที่ซ่อม','วันที่']));

    // แตกคอลัมน์ G → ระยะเช็ค / เลขไมล์ / รายการ
    var interval = '', odo = '';
    var mi = g.match(/(?:ตรวจ)?เช[็้ด]ค\s*ระยะ\s*([\d,\.]+)\s*ก(?:ม|ิโล)?/);
    if(mi) interval = mi[1].replace(/[,\.\s]/g,'');
    var mo = g.match(/(?:เลข)?\s*ไมล[์ื]?\s*[:：]?\s*\(?\s*([\d][\d,\.\s]*\d|\d)/);
    if(mo) odo = mo[1].replace(/[,\.\s]/g,'');
    var desc = g;
    if(mi) desc = desc.replace(mi[0], '');
    desc = desc.replace(/\(?\s*(?:เลข)?\s*ไมล[์ื]?\s*[:：]?\s*\(?\s*[\d][\d,\.\s]*\d\s*\)?\s*ก?ม?\.?/, '');
    desc = desc.replace(/\s+/g,' ').replace(/^[\s+\-\.\/()]+|[\s+\-\.\/()]+$/g,'').trim();

    var vk = pmap[plate.replace(/\s/g,'')] || '';
    var ptype = /^[ก-ฮ]\s*\d{1,4}$/.test(plate) ? 'แดง' : 'ขาว';
    rows.push([ Utilities.getUuid(), vk, plate, ptype, dstr, refno, interval, odo, desc, money ]);
  });

  var out = sh(SHEETS.HIST);
  out.clear();
  out.getRange(1,1,1,HH.length).setValues([HH]);
  if(rows.length) out.getRange(2,1,rows.length,HH.length).setValues(rows);
  return 'สร้าง MaintenanceRecords ' + rows.length + ' แถว';
}
