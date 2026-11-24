/*
 * ตั้งค่าหน้าเว็บ GitHub Pages
 * วางลิงก์ Web App ของ Apps Script (ลงท้ายด้วย /exec) ใน APP_URL
 * Deploy → Manage deployments → คัดลอก "Web app URL"
 */
window.ECO_CONFIG = {
  APP_URL: 'https://script.google.com/macros/s/AKfycbzrbfcVp3xfSEX08Hv-W2mgo3ITj-iv3D4K75JmF2WW_ubfgJQ2zu49LOgfEXaqkJPV/exec',                       // เช่น 'https://script.google.com/macros/s/xxxxxxxx/exec'
  ORG_NAME: 'Green Somdej',
  APP_NAME: 'ระบบติดตามโครงการสิ่งแวดล้อมโรงพยาบาล',

  /*
   * ประกาศความเป็นส่วนตัว (PDPA) — แก้ข้อความด้านล่างให้เป็นของหน่วยงานจริง และให้ผู้รับผิดชอบ/DPO ตรวจก่อนเปิดใช้
   * VERSION : เปลี่ยนทุกครั้งที่แก้เนื้อหาประกาศ → ผู้ใช้เดิมทุกคนจะถูกขอให้กดรับทราบใหม่เมื่อเข้าสู่ระบบ
   * RETENTION_YEARS : ระยะเก็บบันทึกการร่วมกิจกรรมของโครงการ (ปี) — ให้ตรงกับนโยบายของหน่วยงาน
   */
  PRIVACY: {
    VERSION: '2026-11-v1',
    CONTROLLER: 'โรงพยาบาล…… (ใส่ชื่อหน่วยงานผู้ควบคุมข้อมูลส่วนบุคคล)',
    CONTACT: '…… (ใส่อีเมลหรือเบอร์โทรของผู้ดูแลข้อมูลส่วนบุคคล / DPO)',
    RETENTION_YEARS: 3
  }
};
