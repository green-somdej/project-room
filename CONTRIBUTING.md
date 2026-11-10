# Contributing

ขอบคุณที่สนใจร่วมพัฒนา 🌱

## เริ่มต้น
```bash
git clone https://github.com/green-somdej/project-room.git
cd project-room
npm run check      # ตรวจ syntax Code.gs และ JS.html
```
ทดสอบด้วย Apps Script ของตนเอง: `cp .clasp.json.example .clasp.json` ใส่ Script ID แล้ว `npm install && npm run push`

## หลักการโค้ด
- Backend ทั้งหมดอยู่ใน `src/Code.gs` แบ่งเป็นโมดูลตามโครงการ (หัวข้อ `// ===== … =====`) ฟังก์ชัน private ลงท้ายด้วย `_`
- ฟังก์ชันเขียนข้อมูลต้องตรวจสิทธิ์ด้วย `assertAuthorized_` (หรือ helper ของโมดูล) เสมอ
- ค่าคงที่ที่ปรับตามบริบทควรอ่านจากชีต Settings ไม่ฝังในโค้ด
- ห้ามใส่ข้อมูลจริง (ชื่อ เบอร์โทร PIN Script ID) ในโค้ดหรือ issue

## เพิ่มโครงการใหม่
1. เพิ่มชีต/หัวคอลัมน์ใน `setupSheets()`
2. เขียนฟังก์ชัน `get…Overview`, `submit…`, `add/update/delete…` ในโมดูลใหม่
3. เพิ่มแท็บใน `Index.html` (`data-tab`) และตรรกะใน `JS.html`
4. เพิ่มธีมสี `.theme-xxx` ใน `CSS.html` และโครงการเชื่อมโยง (`LINKED_PROJECT_*`)

## Pull Request
Branch จาก `main` · ใช้ข้อความ commit แบบ `feat:`, `fix:`, `docs:` · ให้ `npm run check` ผ่าน · อธิบายวิธีทดสอบใน PR
