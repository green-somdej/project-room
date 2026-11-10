# Architecture

```
Browser ──google.script.run──▶ Code.gs (doGet/doPost) ──▶ Google Sheets / Drive / Gmail
```

| ชั้น | ไฟล์ | หน้าที่ |
|---|---|---|
| Entry | `doGet`, `doPost` | เสิร์ฟ `Index` (template) และ REST API (`?api=ping`, `summary`, `departments`, `emDaily`) |
| Data | `getSheet_`, `sheetToObjects_`, แคชใน `READ_CACHE_` | อ่านชีตเป็นอ็อบเจ็กต์ พร้อมแคมต่อคำขอ |
| Auth | `verifyPin`, `assertAuthorized_`, `sha256Hex_` | ตรวจ PIN (แฮช SHA-256) และสิทธิ์ `IsAdmin` |
| Modules | หัวข้อ `// ===== โครงการ … =====` | อ่าน/เขียนข้อมูลรายโครงการ |
| Impact | `getEnvironmentalImpactSummary` | แปลงกิจกรรมเป็นต้นไม้/CO₂ จากชีต `EnvironmentalSettings` |
| Linked projects | `LINKED_PROJECT_*` | แถวตัวแทนแต่ละแคมเปญในชีต `Projects` เพื่อแสดงในภาพรวม/Gantt |
| Q&A | `askAssistant` | ตอบจากข้อมูลโครงการ, ส่งอีเมลต่อเมื่อตอบไม่ได้, ปฏิเสธคำขอ PIN |

## ชีตหลัก
`Departments`, `Projects`, `Logs`, `FryToFlyLog`, `Paperless*`, `Cartridge*`, `Share*`, `Buyers`, `FryToFlyShipments`, `KitchenOilShipments`, `CalendarDonations`, `ProjectQuestions`, `EnvironmentalSettings`, `IntegrationSettings` ฯลฯ (โครงสร้างคอลัมน์อยู่ในคอมเมนต์หัว `Code.gs`)

## Integration API
`GET <exec-url>?api=ping` (ไม่ต้องใช้คีย์) · `?api=summary&key=…&year=2026` · `?api=departments&key=…` · `?api=emDaily&key=…&days=60` · `POST` JSON `{key,type:"wasteDaily"|"departments",rows|names}`
