// ===== กันบันทึกซ้ำเมื่อหน้าเว็บส่งคำขอเดิมซ้ำ (idempotency) =====
// apiPost ในหน้าเว็บลองส่งใหม่เองสูงสุด 4 ครั้งเมื่อคำตอบอ่านไม่ได้ แต่ GAS มักบันทึกสำเร็จไปแล้วตั้งแต่ครั้งแรก
// → ได้ใบแจ้งซ่อม/ใบจองซ้ำ (เคยเกิด HRC-202609-019/030/031)
// หน้าเว็บจึงแนบ client_ref เดิมทุกครั้งที่ส่งซ้ำ ฝั่งนี้จำผลของ client_ref ไว้ 6 ชม. แล้วตอบผลเดิมแทนการทำซ้ำ
var IDEM_TTL = 21600;

function idem_(action, p, run){
  var ref = String(p && p.client_ref || '');
  if(!/^[\w-]{8,64}$/.test(ref)) return run();          // หน้าเว็บรุ่นเก่าที่ยังไม่ส่ง client_ref = ทำงานแบบเดิม
  var cache = CacheService.getScriptCache(), key = 'idem:' + action + ':' + ref;

  // จองคิวแบบสั้นๆ (ไม่ถือล็อกระหว่างทำงานจริง เพราะบางคำสั่งใช้ล็อกของตัวเองอยู่แล้ว)
  var lk = LockService.getScriptLock();
  if(!lk.tryLock(10000)) throw 'ระบบกำลังบันทึกรายการอื่น ลองใหม่อีกครั้ง';
  var hit;
  try {
    hit = cache.get(key);
    if(!hit) cache.put(key, 'RUNNING', 600);
  } finally { lk.releaseLock(); }

  if(hit){
    // คำขอแรกยังทำงานอยู่ → รอผล (สูงสุด ~25 วิ)
    for(var i = 0; hit === 'RUNNING' && i < 25; i++){ Utilities.sleep(1000); hit = cache.get(key); }
    if(hit === 'RUNNING') throw 'กำลังบันทึกรายการนี้อยู่ รอสักครู่แล้วเปิดดูอีกครั้ง';
    if(hit && hit !== 'RUNNING'){
      var saved = JSON.parse(hit);
      if(saved.big) throw 'บันทึกรายการนี้แล้ว กรุณารีเฟรชเพื่อดูผล';
      return saved.data;
    }
    // ผลหายจากแคช (เช่นคำขอแรกล้มเหลว) → ทำใหม่ตามปกติ
  }

  try {
    var data = run();
    var s = JSON.stringify({ data: data });
    cache.put(key, s.length < 90000 ? s : JSON.stringify({ big: true }), IDEM_TTL);
    return data;
  } catch(err){
    cache.remove(key);                                   // ล้มเหลว = ลองใหม่ได้
    throw err;
  }
}
