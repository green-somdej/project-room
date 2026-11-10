# Changelog
รูปแบบอิง [Keep a Changelog](https://keepachangelog.com/) · [SemVer](https://semver.org/)

## [1.2.0] - 2026-11-10
### Added
- หน้าแรก GitHub Pages เปิดระบบจริง (Apps Script Web App) ให้คลิกใช้งานได้ในหน้าเดียว พร้อมแถบเลือกโครงการ หน้าโหลดแบบเคลื่อนไหว โหมดเต็มจอ ธีมมืด/สว่าง และปุ่มเปิดในแท็บใหม่
- ตั้งค่าลิงก์ระบบที่ `docs/config.js` (`APP_URL`)
- ลิงก์ตรง `?page=paperless|envreport|archive|projects` เปิดแท็บนั้นได้ทันที
### Changed
- หน้าแนะนำโครงการย้ายไปที่ `docs/about.html`

## [1.1.0] - 2026-11-09
### Added
- ผู้ดูแลระบบเพิ่มชื่อหน่วยงานใหม่จากหน้าเว็บได้ (Paperless / ตลับหมึก / คัดแยก แบ่งปัน) ในหน้าต่าง "🏢 เลือกหน่วยงานที่ใช้" — เพิ่มพร้อมกันหลายโครงการได้, ชื่อซ้ำที่ปิดไว้จะถูกเปิดใช้กลับ, หน่วยงาน Paperless ตั้ง PIN ได้ทันที (เก็บเป็นแฮช)
- ฟังก์ชัน `addProjectDepartment()` ใน Code.gs

## [1.0.0] - 2026-11-09
### Added
- รีโปเริ่มต้น: Backend (Code.gs), หน้าเว็บ (Index/CSS/JS), `appsscript.json`
- 8 แคมเปญ: ทอดไม่ทิ้ง, Paperless, ตลับหมึก, จากครัวสู่ท้องฟ้า, ปฏิทินเก่า, คัดแยก แบ่งปัน, น้ำยา EM, ปุ๋ยหมักใบไม้
- เว็บไซต์ GitHub Pages (TH/EN, dark/light), CI, issue/PR templates
- สคริปต์ `npm run check`, ไฟล์ตัวอย่าง clasp
