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
| Executive Summary | `ExecSlides.gs`: `setupExecutiveSummaries`, `executiveSummaryAutoUpdate`, `xsRun_` | สร้าง Google Slides 16:9 ≤10 หน้าต่อโครงการ ผ่าน Slides API `batchUpdate` อัปเดตเมื่อลายนิ้วมือข้อมูลเปลี่ยน (ทริกเกอร์รายชั่วโมง) ลงทะเบียนในชีต `ExecutiveDecks` |
| Executive Deck (ในระบบ) | `getExecutiveDeckData` (ExecSlides.gs) + `xd*` (JS.html) | สไลด์ 16:9 รวมทุกโครงการ ≤10 หน้า สร้างสดในเบราว์เซอร์จากรายงานสิ่งแวดล้อม + `getPeriodSummary` ทุกแคมเปญ + เป้าหมาย + สถานะโครงการ |
| Cartridge stock | `ctgLoad_`, `ctgApplyItems_`, `ctgAllocateFifo_`, `ctgRelease_`, `ctgMerge_` | การจำหน่ายตัดยอดรายการ "รอจำหน่าย" ที่เลือก (แยกจำหน่ายได้/ไม่ได้) ผูกด้วย `SaleID` แบ่งแถวด้วย `SplitFrom` |
| Eco Hub (ร่วมโครงการ) | `Hub.gs`: `hubHandleGet_`, `hubHandlePost_`, `hubCatalog_`, `hubJoinStaff_`, `hubDeptJoin_`, `hubRegister_`, `hubAdmin_` + `docs/join.html` | ผู้ดูแลกำหนดว่าโครงการไหนเปิดให้ เจ้าหน้าที่/หน่วยงาน/ประชาชน/บริษัท (`HubProjects`) · สมัครหลายโครงการครั้งเดียว · ดูรายละเอียดใน [HUB.md](HUB.md) |
| ปีงบประมาณ | `getPeriodSummary(key, year, basis)`, `periodYearOf_` | `basis='fiscal'` = 1 ต.ค.–30 ก.ย. (ปีงบ = ปี ค.ศ. ที่สิ้นสุด) ใช้ทั้ง Hub และการ์ดสรุปในแอป |
| Q&A | `askAssistant` | ตอบจากข้อมูลโครงการ, ส่งอีเมลต่อเมื่อตอบไม่ได้, ปฏิเสธคำขอ PIN |

## ชีตหลัก
`Departments`, `Projects`, `Logs`, `FryToFlyLog`, `Paperless*`, `Cartridge*`, `Share*`, `Buyers`, `FryToFlyShipments`, `KitchenOilShipments`, `CalendarDonations`, `ProjectQuestions`, `EnvironmentalSettings`, `IntegrationSettings` ฯลฯ (โครงสร้างคอลัมน์อยู่ในคอมเมนต์หัว `Code.gs`)

## Integration API
สาธารณะ (ไม่ต้องใช้คีย์): `?api=hubHome|hubProject|hubDepartments|hubAdminDepts` · POST `{action:"hub…"}` (ดู HUB.md)
`GET <exec-url>?api=ping` (ไม่ต้องใช้คีย์) · `?api=summary&key=…&year=2026` · `?api=departments&key=…` · `?api=emDaily&key=…&days=60` · `POST` JSON `{key,type:"wasteDaily"|"departments",rows|names}`
