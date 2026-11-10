/**
 * ระบบติดตามโครงการสิ่งแวดล้อมโรงพยาบาล
 * Backend: Google Apps Script + Google Sheets เป็นฐานข้อมูล
 *
 * โครงสร้างชีต:
 *   - Departments : DepartmentID, DepartmentName, DepartmentEmail, PinHash, IsAdmin, IsActive, SortOrder, CreatedAt, UpdatedAt
 *                   (IsAdmin = TRUE/FALSE — กำหนดว่าหน่วยงานนี้มีสิทธิ์ผู้ดูแลระบบ [เพิ่ม/แก้ไข/ลบโครงการ] หรือไม่
 *                    แก้ไขได้ตรงในชีตนี้เลย ไม่ต้องแก้โค้ด)
 *   - Projects    : ProjectID, ProjectName, Category, Description, Department, ResponsiblePerson, ResponsibleUnit,
 *                   StartDate, TargetDate, CurrentStatus, StatusCategory, Progress, IsArchived,
 *                   CreatedAt, UpdatedAt, CurrentStage
 *   - Logs        : LogID, ProjectID, LogDate, StatusText, StatusCategory, Progress,
 *                   RecordedByName, RecordedByUnit, CreatedAt, IsDeleted, Stage
 *   - FryToFlyLog : LogID, Timestamp, StaffCode, StaffName, Phone, ActionType, Quantity, Note
 *   - PaperlessDepartments : DepartmentID, DepartmentName, PinHash, IsActive, SortOrder, CreatedAt, UpdatedAt
 *   - PaperlessLog         : LogID, DepartmentID, DepartmentName, Year, Month, ItemCount, SheetsReduced,
 *                            Note, RecordedByName, CreatedAt, UpdatedAt  (1 แถวต่อ หน่วยงาน+ปี+เดือน, upsert)
 *   - ProjectQuestions     : QuestionID, ProjectID, ProjectName, AskerName, AskerContact, QuestionText,
 *                            Status (pending/answered-auto/answered-admin), AutoAnswer, AdminAnswer,
 *                            AnsweredByName, CreatedAt, AnsweredAt
 *   - CartridgeDepartments  : DepartmentID, DepartmentName, SortOrder, IsActive (รายชื่อหน่วยงานที่ส่งตลับหมึกเข้าโครงการ)
 *   - CartridgeContributions: ContributionID, DepartmentName, ContributionDate, Quantity, Note, CreatedAt, Status
 *                             (บันทึกสาธารณะ ไม่ต้องใช้ PIN — หน่วยงานแจ้งจำนวนตลับหมึกที่ส่งเข้าจุดรับ
 *                              Status = รอจำหน่าย / จำหน่ายแล้ว / จำหน่ายไม่ได้ — ผู้ดูแลระบบเปลี่ยนสถานะได้)
 *   - CartridgeSales        : SaleID, SaleDate, BuyerName, Quantity, Amount, Note, RecordedByName, CreatedAt, UpdatedAt
 *                             (บันทึกรอบจำหน่ายจริงให้ร้านรับซื้อ — ผู้ดูแลระบบ [IsAdmin] เท่านั้นที่เพิ่ม/แก้ไข/ลบได้)
 *   - CartridgeSettings     : Key, Value (ค่าที่ปรับได้โดยไม่ต้องแก้โค้ด เช่น WeightPerUnitKg, DisposalCostPerKg)
 *   - ShareItems / ShareSubmissions / ShareWithdrawals / ShareDepartments / ShareStockLog : โครงการ "คัดแยก แบ่งปัน" (ดูรายละเอียดในส่วนโครงการคัดแยก แบ่งปัน)
 *   - Buyers                : BuyerID, BuyerName, BuyerKey, Phone, Projects, Note, IsActive, CreatedAt
 *                             (ทะเบียนผู้รับซื้อ ใช้ร่วมกันทุกโครงการ — ระบบเพิ่มให้อัตโนมัติเมื่อบันทึกการส่งจำหน่าย และดึงรายชื่อที่มีอยู่แล้วมาให้เลือกอัตโนมัติ)
 *   - FryToFlyShipments     : ShipmentID, ShipmentDate, BuyerName, Quantity, Unit, Liters, UnitPrice, Amount, Note, RecordedByName, CreatedAt, UpdatedAt
 *                             (โครงการทอดไม่ทิ้ง: การส่งน้ำมันรายรอบ จำนวนลิตร + ราคา ซึ่งแต่ละรอบไม่เท่ากัน)
 *   - KitchenOilShipments   : โครงสร้างเดียวกับ FryToFlyShipments (โครงการ "จากครัวสู่ท้องฟ้า" ฝ่ายโภชนาการ — หน่วยเป็นลิตรหรือปี๊บ)
 *   - CalendarDonations     : DonationID, DonationDate, RoundName, DonorType, DonorName, WeightKg, Note, RecordedByName, CreatedAt, UpdatedAt
 *                             (โครงการ "ปฏิทินเก่าเราขอ": บริจาคปฏิทินเก่าทำอักษรเบรลล์ นับเป็นน้ำหนัก กก. ต่อรอบ/ต่อผู้บริจาค)
 *   - CartridgeSales เพิ่มคอลัมน์ UnitPrice (ราคาต่อตลับของรอบนั้น — ระบบเพิ่มคอลัมน์ให้เองเมื่อรัน setupSheets)
 */

// ============================= CONFIG =============================
const SHEET_PROJECTS    = 'Projects';
const SHEET_DEPARTMENTS = 'Departments';
const SHEET_LOGS        = 'Logs';
const SHEET_FTF_LOG     = 'FryToFlyLog';
const SHEET_PL_DEPTS    = 'PaperlessDepartments';
const SHEET_PL_LOG      = 'PaperlessLog';
const SHEET_CTG_DEPTS    = 'CartridgeDepartments';
const SHEET_CTG_CONTRIB  = 'CartridgeContributions';
const SHEET_CTG_SALES    = 'CartridgeSales';
const SHEET_CTG_SETTINGS = 'CartridgeSettings';

// สถานะของรายการ "ส่งตลับหมึกเข้าโครงการ" (CartridgeContributions.Status)
const CTG_STATUS_PENDING   = 'รอจำหน่าย';
const CTG_STATUS_SOLD      = 'จำหน่ายแล้ว';
const CTG_STATUS_REJECTED  = 'จำหน่ายไม่ได้';

const STATUS_CATEGORIES = ['ยังไม่เริ่ม', 'กำลังดำเนินการ', 'ล่าช้า', 'เสร็จสมบูรณ์', 'ระงับ'];

const PL_MONTH_NAMES = ['มกราคม', 'กุมภาพันธ์', 'มีนาคม', 'เมษายน', 'พฤษภาคม', 'มิถุนายน',
  'กรกฎาคม', 'สิงหาคม', 'กันยายน', 'ตุลาคม', 'พฤศจิกายน', 'ธันวาคม'];

// ProjectID คงที่ของ "โครงการเชื่อมโยง" — แถวตัวแทนของแคมเปญทอดไม่ทิ้ง/Paperless ในชีต Projects
// (ให้ปรากฏในภาพรวม/Timeline/สัดส่วนตามประเภทร่วมกับโครงการสิ่งแวดล้อมอื่นๆ แต่รายละเอียดเต็มยังอยู่ที่แท็บเดิม)
const LINKED_PROJECT_FTF_ID = 'PRJ-FTF-LINK';
const LINKED_PROJECT_PL_ID  = 'PRJ-PAPERLESS-LINK';
const LINKED_PROJECT_CTG_ID = 'PRJ-CARTRIDGE-LINK';
const LINKED_PROJECT_KIT_ID = 'PRJ-KITCHEN-LINK';
const LINKED_PROJECT_CAL_ID = 'PRJ-CALENDAR-LINK';
const LINKED_PROJECT_EM_ID  = 'PRJ-EM-LINK';
const LINKED_PROJECT_CP_ID  = 'PRJ-COMPOST-LINK';

// โครงการ "ตลับหมึกเก่าเราขอ": น้ำหนักตลับหมึก 1 ตลับ (กก.) และราคาค่าทำลายขยะอันตราย (บาท/กก.)
// ถ้าไม่ได้ขาย ต้องเสียค่าทำลายตามราคานี้ — ใช้คำนวณ "มูลค่าที่ประหยัดได้" เทียบกับการขาย
// ค่าเริ่มต้นนี้ใช้เป็น fallback เท่านั้น — ค่าจริงที่ใช้คำนวณอ่านจากชีต CartridgeSettings (แก้ไขได้โดยไม่ต้องแก้โค้ด ผ่าน getCartridgeSettings_())
const CARTRIDGE_WEIGHT_PER_UNIT_KG = 0.936;
const CARTRIDGE_DISPOSAL_COST_PER_KG = 55;

// แต้มต่อ 1 แกลลอนที่คืน/บริจาค และเกณฑ์ระดับสมาชิก Green Somdej
const FTF_POINTS_PER_GALLON = 50;
const FTF_TIERS = [
  { min: 400, key: 'gold',   label: 'ระดับทอง' },
  { min: 200, key: 'silver', label: 'ระดับเงิน' },
  { min: 100, key: 'bronze', label: 'ระดับทองแดง' },
  { min: 0,   key: 'member', label: 'Green Somdej' }
];
function ftfTierFor_(points) {
  return FTF_TIERS.find(t => points >= t.min) || FTF_TIERS[FTF_TIERS.length - 1];
}

// ============================= ผลกระทบด้านสิ่งแวดล้อม (สรุปรวมทุกโครงการ) =============================
// ค่าตัวแปรด้านล่างเป็นค่าประมาณ/สมมติฐานที่ใช้ทั่วไป ใช้เป็น fallback เท่านั้น — ค่าจริงอ่านจากชีต EnvironmentalSettings
// (แก้ไขได้โดยไม่ต้องแก้โค้ด ผ่าน getEnvironmentalSettings_()) เพื่อให้ปรับตามข้อมูลอ้างอิงที่โรงพยาบาลต้องการใช้ได้
const SHEET_ENV_SETTINGS = 'EnvironmentalSettings';
const ENV_PAPER_SHEETS_PER_TREE = 8333;       // กระดาษ A4 ที่ผลิตได้จากต้นไม้ 1 ต้น (ประมาณ)
const ENV_PAPER_CO2_PER_SHEET_KG = 0.0046;    // CO2 ที่ปล่อยจากการผลิตกระดาษ A4 1 แผ่น (กก.)
const ENV_FTF_LITERS_PER_GALLON = 4;          // 1 แกลลอน = 4 ลิตร (ตามที่ใช้ในโครงการทอดไม่ทิ้ง)
const ENV_FTF_CO2_PER_LITER_KG = 2.5;         // CO2 ที่ลดได้จากการนำน้ำมันใช้แล้ว 1 ลิตร ไปผลิตไบโอดีเซล/กำจัดอย่างถูกวิธี แทนทิ้งลงท่อ/พื้นดิน (ประมาณ)
const ENV_CARTRIDGE_CO2_PER_KG = 2;           // CO2 ที่ลดได้จากการรีไซเคิลตลับหมึก 1 กก. เทียบกับทิ้งเป็นขยะอันตราย/ผลิตใหม่ (ประมาณ)
const ENV_CO2_ABSORBED_PER_TREE_KG_YEAR = 20; // ต้นไม้โตเต็มที่ 1 ต้น ดูดซับ CO2 ได้ประมาณปีละ (กก.) — ใช้แปลง "CO2 ที่ลดได้รวม" เป็นจำนวนต้นไม้เทียบเท่า

// ============================= SETUP (รันครั้งเดียว) =============================
/**
 * รันจาก Apps Script editor (เลือก setupSheets แล้วกด ▶ Run)
 * สร้างชีตที่ยังไม่มี พร้อม Header ชีตที่มีอยู่แล้วจะไม่แตะต้อง
 */
function setupSheets() {
  const schema = {
    'Departments': ['DepartmentID', 'DepartmentName', 'DepartmentEmail', 'PinHash', 'IsAdmin', 'IsActive', 'SortOrder', 'CreatedAt', 'UpdatedAt'],
    'Projects': ['ProjectID', 'ProjectName', 'Category', 'Description', 'Department', 'ResponsiblePerson', 'ResponsibleUnit',
                 'StartDate', 'TargetDate', 'CurrentStatus', 'StatusCategory', 'Progress', 'IsArchived',
                 'CreatedAt', 'UpdatedAt', 'CurrentStage', 'AutoSummary', 'AutoSyncedAt', 'StatusSource', 'ArchivedAt'],
    'Logs': ['LogID', 'ProjectID', 'LogDate', 'StatusText', 'StatusCategory', 'Progress',
             'RecordedByName', 'RecordedByUnit', 'CreatedAt', 'IsDeleted', 'Stage'],
    'FryToFlyLog': ['LogID', 'Timestamp', 'StaffCode', 'StaffName', 'Phone', 'ActionType', 'Quantity', 'Note', 'VerifyStatus', 'VerifiedBy', 'VerifiedAt'],
    'PaperlessDepartments': ['DepartmentID', 'DepartmentName', 'PinHash', 'IsActive', 'SortOrder', 'CreatedAt', 'UpdatedAt'],
    'PaperlessLog': ['LogID', 'DepartmentID', 'DepartmentName', 'Year', 'Month', 'ItemCount', 'SheetsReduced',
                      'Note', 'RecordedByName', 'CreatedAt', 'UpdatedAt'],
    'ProjectQuestions': ['QuestionID', 'ProjectID', 'ProjectName', 'AskerName', 'AskerContact', 'QuestionText',
                          'Status', 'AutoAnswer', 'AdminAnswer', 'AnsweredByName', 'CreatedAt', 'AnsweredAt'],
    'CartridgeDepartments': ['DepartmentID', 'DepartmentName', 'SortOrder', 'IsActive'],
    'CartridgeContributions': ['ContributionID', 'DepartmentName', 'ContributionDate', 'Quantity', 'Note', 'CreatedAt', 'Status', 'VerifyStatus', 'VerifiedBy', 'VerifiedAt'],
    'CartridgeSales': ['SaleID', 'SaleDate', 'BuyerName', 'Quantity', 'Amount', 'Note', 'RecordedByName', 'CreatedAt', 'UpdatedAt', 'UnitPrice'],
    'Buyers': BUYER_HEADERS,
    'FryToFlyShipments': SHIP_HEADERS,
    'KitchenOilShipments': SHIP_HEADERS,
    'CalendarDonations': CAL_HEADERS,
    'Inquiries': INQUIRY_HEADERS,
    'ProjectSettings': PROJECT_SETTINGS_HEADERS,
    'ProjectTargets': ['TargetKey', 'ProjectKey', 'Year', 'Month', 'TargetValue', 'UpdatedBy', 'UpdatedAt'],
    'EMBatches': EM_BATCH_HEADERS,
    'EMChecks': EM_CHECK_HEADERS,
    'EMUsage': EM_USAGE_HEADERS,
    'EMSettings': ['Key', 'Value'],
    'EMDailyLogs': EM_LOG_HEADERS,
    'CompostBatches': CP_BATCH_HEADERS,
    'CompostChecks': CP_CHECK_HEADERS,
    'CompostUsage': CP_USAGE_HEADERS,
    'CompostSettings': ['Key', 'Value'],
    'CompostDailyLogs': CP_LOG_HEADERS,
    'IntegrationSettings': ['Key', 'Value'],
    'ExternalWasteDaily': EXT_WASTE_HEADERS,
    'ExternalDepartments': EXT_DEPT_HEADERS,
    'CalendarSubmissions': CAL_SUB_HEADERS,
    'ProjectPhotos': PHOTO_HEADERS,
    'CartridgeSettings': ['Key', 'Value'],
    'EnvironmentalSettings': ['Key', 'Value'],
    'ShareItems': ['ItemID', 'ItemName', 'Category', 'Unit', 'Quantity', 'UnitValue', 'Note', 'LastCheckedAt', 'LastCheckedBy', 'IsActive', 'CreatedAt', 'UpdatedAt'],
    'ShareSubmissions': ['SubmissionID', 'SubmitDate', 'DepartmentName', 'SubmitterName', 'Position', 'UsageStatus', 'ItemName', 'Category', 'Quantity', 'Unit', 'Note', 'Status', 'LinkedItemID', 'ReviewedBy', 'ReviewedAt', 'CreatedAt'],
    'ShareWithdrawals': ['WithdrawalID', 'WithdrawDate', 'Category', 'ItemName', 'Quantity', 'Unit', 'UnitPrice', 'DepartmentName', 'ItemID', 'Note', 'RecordedByName', 'CreatedAt'],
    'ShareDepartments': ['DepartmentID', 'DepartmentName', 'SortOrder', 'IsActive'],
    'ShareStockLog': ['LogID', 'Timestamp', 'ItemID', 'ItemName', 'Action', 'QtyBefore', 'QtyAfter', 'By', 'Note'],
    'PortalVisits': PORTAL_VISIT_HEADERS,
    'ProductRequests': PREQ_HEADERS,
    'RawMaterialMonthly': RAW_HEADERS
  };

  const ss = getSS();
  const existingNames = ss.getSheets().map(s => s.getName().trim().toLowerCase());
  const created = [];
  const skipped = [];

  Object.keys(schema).forEach(name => {
    if (existingNames.indexOf(name.trim().toLowerCase()) !== -1) {
      skipped.push(name);
      return;
    }
    const sh = ss.insertSheet(name);
    const headers = schema[name];
    sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
    sh.setFrozenRows(1);
    created.push(name);
  });

  if (created.indexOf('Departments') !== -1) {
    const sh = ss.getSheetByName('Departments');
    const now = new Date();
    // ลำดับคอลัมน์: DepartmentID, DepartmentName, DepartmentEmail, PinHash, IsAdmin, IsActive, SortOrder, CreatedAt, UpdatedAt
    sh.appendRow(['DEPT-001', 'งานสิ่งแวดล้อม', '', sha256Hex_('123456'), true, true, 1, now, now]);
    Logger.log('เพิ่มหน่วยงานตัวอย่าง "งานสิ่งแวดล้อม" PIN เริ่มต้น 123456 (สิทธิ์ผู้ดูแลระบบ = เปิด) (ควรเปลี่ยน PIN ภายหลังด้วย generatePinHash)');
  }

  // ชีต Departments ที่มีอยู่แล้วก่อนอัปเดตนี้ อาจยังไม่มีคอลัมน์ IsAdmin — เพิ่มให้โดยไม่กระทบข้อมูลเดิม
  ensureDepartmentIsAdminColumn_();

  // เติมรายชื่อหน่วยงานตั้งต้นในชีต CartridgeDepartments (ถ้ายังไม่มี) — ทำงานซ้ำได้อย่างปลอดภัย
  ensureCartridgeDepartmentsSeeded_();

  // ชีต CartridgeContributions ที่มีอยู่แล้วก่อนอัปเดตนี้ อาจยังไม่มีคอลัมน์ Status — เพิ่มให้โดยไม่กระทบข้อมูลเดิม
  ensureCartridgeContributionStatusColumn_();

  // เติมค่าตั้งต้น (น้ำหนักต่อตลับ/ราคาค่าทำลายขยะ) ในชีต CartridgeSettings ถ้ายังไม่มี — ทำงานซ้ำได้อย่างปลอดภัย
  ensureCartridgeSettingsSeeded_();

  // เติมค่าตั้งต้นของตัวแปรคำนวณผลกระทบด้านสิ่งแวดล้อมในชีต EnvironmentalSettings ถ้ายังไม่มี — ทำงานซ้ำได้อย่างปลอดภัย
  ensureEnvironmentalSettingsSeeded_();

  // เติมรายชื่อหน่วยงานตั้งต้นของโครงการ "คัดแยก แบ่งปัน" (ชีต ShareDepartments) — ทำงานซ้ำได้อย่างปลอดภัย
  ensureShareDepartmentsSeeded_();

  // ชีต CartridgeSales ที่มีอยู่แล้ว อาจยังไม่มีคอลัมน์ UnitPrice (ราคาต่อตลับของแต่ละรอบ) — เพิ่มให้โดยไม่กระทบข้อมูลเดิม
  ensureCartridgeSalesUnitPriceColumn_();

  // ทะเบียนผู้รับซื้อ: ดึงรายชื่อผู้ซื้อที่เคยบันทึกไว้แล้ว (เช่น จากรายการจำหน่ายตลับหมึกย้อนหลัง) มาลงทะเบียน — ทำงานซ้ำได้อย่างปลอดภัย
  ensureBuyersSeeded_();

  Logger.log('สร้างชีตใหม่: [' + created.join(', ') + ']');
  Logger.log('ชีตที่มีอยู่แล้ว (ไม่แตะต้อง): [' + skipped.join(', ') + ']');
  Logger.log('สเปรดชีต: ' + ss.getName() + '  |  ' + ss.getUrl());

  // เพิ่มแถว "โครงการทอดไม่ทิ้ง", "โครงการ Paperless" และ "โครงการตลับหมึกเก่าเราขอ" เข้าไปในรายการโครงการสิ่งแวดล้อม
  // (ทำให้นับรวมในภาพรวม/Timeline/สัดส่วนตามประเภท) — ทำงานซ้ำได้อย่างปลอดภัย ไม่สร้างแถวซ้ำ
  ensureProjectExtraColumns_();
  ensureVerifyColumns_(SHEET_FTF_LOG);
  ensureVerifyColumns_(SHEET_CTG_CONTRIB);
  ensureEmSettingsSeeded_();
  ensureCpSettingsSeeded_();
  ensureLinkedCampaignProjects();
}

// ============================= โครงการเชื่อมโยง: ทอดไม่ทิ้ง / Paperless =============================
/**
 * เพิ่มแถวตัวแทนของแคมเปญ "ทอดไม่ทิ้ง" และ "Paperless" ลงในชีต Projects (ถ้ายังไม่มี)
 * เพื่อให้นับรวมเป็นโครงการหนึ่งในภาพรวม/Timeline/สัดส่วนโครงการตามประเภท ร่วมกับโครงการสิ่งแวดล้อมอื่นๆ
 * รายละเอียดเต็มของแต่ละแคมเปญยังคงอยู่ที่แท็บเดิม (ทอดไม่ทิ้ง / Paperless) — แถวนี้เป็นเพียงตัวแทนสรุปสถานะ
 * รันซ้ำได้อย่างปลอดภัย (ไม่สร้างแถวซ้ำถ้ามี ProjectID นี้อยู่แล้ว)
 */
function ensureLinkedCampaignProjects() {
  const sh = getSheet_(SHEET_PROJECTS);
  const existing = sheetToObjects_(SHEET_PROJECTS);
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const now = new Date();
  const year = now.getFullYear();
  const jan1 = new Date(year, 0, 1);
  const dec31 = new Date(year, 11, 31);

  function appendIfMissing(projectId, fields) {
    if (existing.some(p => p.ProjectID === projectId)) return false;
    const rowObj = Object.assign({
      ProjectID: projectId,
      IsArchived: false,
      CreatedAt: now,
      UpdatedAt: now,
      StartDate: jan1,
      TargetDate: dec31,
      StatusCategory: 'กำลังดำเนินการ',
      Progress: 0,
      CurrentStage: 'ดำเนินการต่อเนื่องทั้งปี'
    }, fields);
    const row = headers.map(h => (h in rowObj ? rowObj[h] : ''));
    sh.appendRow(row);
    return true;
  }

  const createdFtf = appendIfMissing(LINKED_PROJECT_FTF_ID, {
    ProjectName: 'โครงการ "ทอดไม่ทิ้ง"',
    Category: 'การจัดการขยะ',
    Description: 'รวบรวมน้ำมันพืช/สัตว์ใช้แล้วจากบุคลากร เพื่อนำไปแปรรูปเป็นน้ำมันเชื้อเพลิงอากาศยานยั่งยืน (SAF) ลดการปล่อยก๊าซเรือนกระจก — ดูรายละเอียดเต็มได้ที่แท็บ "ทอดไม่ทิ้ง"',
    Department: 'ฝ่ายบริหารงานทั่วไป',
    ResponsiblePerson: 'นางสาวสุชาดา รัตนบัวแก้ว',
    ResponsibleUnit: 'ฝ่ายบริหารงานทั่วไป',
    CurrentStatus: 'ดูรายละเอียดเต็มได้ที่แท็บ "ทอดไม่ทิ้ง"'
  });

  const createdPl = appendIfMissing(LINKED_PROJECT_PL_ID, {
    ProjectName: 'โครงการ Paperless (ลดการใช้กระดาษ)',
    Category: 'การจัดการทรัพยากร',
    Description: 'แต่ละหน่วยงานบันทึกจำนวนเรื่องและปริมาณกระดาษที่ลดลงได้ทุกเดือน เพื่อติดตามภาพรวมการลดการใช้กระดาษของโรงพยาบาล — ดูรายละเอียดเต็มได้ที่แท็บ "Paperless"',
    Department: 'ฝ่ายบริหารงานทั่วไป',
    ResponsiblePerson: 'ฝ่ายบริหารงานทั่วไป',
    ResponsibleUnit: 'ฝ่ายบริหารงานทั่วไป',
    CurrentStatus: 'ดูรายละเอียดเต็มได้ที่แท็บ "Paperless"'
  });

  const createdCtg = appendIfMissing(LINKED_PROJECT_CTG_ID, {
    ProjectName: 'โครงการตลับหมึกเก่าเราขอ',
    Category: 'การจัดการขยะ',
    Description: 'รวบรวมตลับหมึกพิมพ์ใช้แล้วจากทุกหน่วยงาน ส่งขายให้ร้านรับซื้อ ลดปริมาณขยะอันตรายที่ต้องเสียค่าทำลาย และสร้างรายได้ให้โรงพยาบาล — ดูรายละเอียดเต็มได้ที่แท็บ "ตลับหมึกเก่าเราขอ"',
    Department: 'คณะอนุกรรมการด้านการจัดการพลังงาน และส่งเสริมกิจกรรม 5ส',
    ResponsiblePerson: 'คณะอนุกรรมการด้านการจัดการพลังงาน',
    ResponsibleUnit: 'ฝ่ายบริหารงานทั่วไป',
    CurrentStatus: 'ดูรายละเอียดเต็มได้ที่แท็บ "ตลับหมึกเก่าเราขอ"'
  });

  const createdShare = appendIfMissing(LINKED_PROJECT_SHARE_ID, {
    ProjectName: 'โครงการคัดแยก แบ่งปัน',
    Category: 'การจัดการทรัพยากร',
    Description: 'ให้ทุกหน่วยงานคัดแยกเครื่องใช้สำนักงานและวัสดุที่เกินจำเป็น/หมดความจำเป็นส่งเข้าโครงการ ตรวจสอบแล้วรับเข้าคลังรอแบ่งปัน เพื่อให้หน่วยงานอื่นนำไปใช้แทนการจัดซื้อใหม่ ลดการสั่งซื้อซ้ำซ้อน ลดของเหลือทิ้ง และประหยัดงบประมาณ — ดูรายละเอียดเต็มได้ที่แท็บ "คัดแยก แบ่งปัน"',
    Department: 'ฝ่ายบริหารงานทั่วไป',
    ResponsiblePerson: 'คณะอนุกรรมการด้านการจัดการพลังงาน',
    ResponsibleUnit: 'ฝ่ายบริหารงานทั่วไป',
    CurrentStatus: 'ดูรายละเอียดเต็มได้ที่แท็บ "คัดแยก แบ่งปัน"',
    StartDate: new Date(2025, 10, 26),
    TargetDate: dec31
  });

  const createdKit = appendIfMissing(LINKED_PROJECT_KIT_ID, {
    ProjectName: 'โครงการ "จากครัวสู่ท้องฟ้า"',
    Category: 'การจัดการขยะ',
    Description: 'ฝ่ายโภชนาการรวบรวมน้ำมันทอดอาหารใช้แล้วจากครัวโรงพยาบาล ส่งให้ผู้รับซื้อเป็นรอบๆ (นับเป็นลิตรหรือปี๊บ ราคาแต่ละรอบไม่เท่ากัน) เพื่อนำไปแปรรูปเป็นพลังงานทดแทน แทนการทิ้งลงท่อระบายน้ำ — ดูรายละเอียดเต็มได้ที่แท็บ "จากครัวสู่ท้องฟ้า"',
    Department: 'ฝ่ายโภชนาการ',
    ResponsiblePerson: 'ฝ่ายโภชนาการ',
    ResponsibleUnit: 'ฝ่ายโภชนาการ',
    CurrentStatus: 'ดูรายละเอียดเต็มได้ที่แท็บ "จากครัวสู่ท้องฟ้า"'
  });

  const createdCal = appendIfMissing(LINKED_PROJECT_CAL_ID, {
    ProjectName: 'โครงการปฏิทินเก่าเราขอ',
    Category: 'การจัดการขยะ',
    Description: 'รับบริจาคปฏิทินเก่าจากหน่วยงาน เจ้าหน้าที่ และบริษัทภายนอกเป็นรอบๆ เพื่อส่งไปจัดทำอักษรเบรลล์ให้ผู้พิการทางสายตา นับเป็นน้ำหนักเพื่อคำนวณผลการช่วยรักษาสิ่งแวดล้อม — ดูรายละเอียดเต็มได้ที่แท็บ "ปฏิทินเก่าเราขอ"',
    Department: 'ฝ่ายบริหารงานทั่วไป',
    ResponsiblePerson: 'คณะอนุกรรมการด้านการจัดการพลังงาน',
    ResponsibleUnit: 'ฝ่ายบริหารงานทั่วไป',
    CurrentStatus: 'ดูรายละเอียดเต็มได้ที่แท็บ "ปฏิทินเก่าเราขอ"'
  });

  const createdEm = appendIfMissing(LINKED_PROJECT_EM_ID, {
    ProjectName: 'โครงการน้ำยา EM จากเศษอาหาร',
    Category: 'การจัดการขยะ',
    Description: 'นำเศษอาหาร ผักและผลไม้ และเนื้อสัตว์จากโรงพยาบาลมาหมักเป็นน้ำยา EM (จุลินทรีย์ที่มีประโยชน์) ตามรอบการหมัก 7 / 15 / 45 วัน เพื่อนำไปใช้บำบัดน้ำเสีย ดับกลิ่น และปรับปรุงดินในพื้นที่ต่างๆ ของโรงพยาบาล — ดูรายละเอียดเต็มได้ที่แท็บ "น้ำยา EM"',
    Department: 'ฝ่ายบริหารงานทั่วไป',
    ResponsiblePerson: 'คณะอนุกรรมการด้านการจัดการพลังงาน',
    ResponsibleUnit: 'ฝ่ายบริหารงานทั่วไป',
    CurrentStatus: 'เริ่มโครงการ — ดูรายละเอียดเต็มได้ที่แท็บ "น้ำยา EM"'
  });

  const createdCp = appendIfMissing(LINKED_PROJECT_CP_ID, {
    ProjectName: 'โครงการปุ๋ยหมักจากขยะใบไม้',
    Category: 'การจัดการขยะ',
    Description: 'นำขยะใบไม้ ใบไม้แห้ง/สด เศษหญ้า และกิ่งไม้ที่ตัดแต่งในพื้นที่โรงพยาบาลมาหมักเป็นปุ๋ยหมักตามรอบการหมัก 7 / 15 / 45 วัน เพื่อนำไปปรับปรุงดินและปลูกต้นไม้ในสวนหย่อมของโรงพยาบาล แทนการเผาหรือทิ้งเป็นขยะ — ดูรายละเอียดเต็มได้ที่แท็บ "ปุ๋ยหมักใบไม้"',
    Department: 'ฝ่ายบริหารงานทั่วไป',
    ResponsiblePerson: 'คณะอนุกรรมการด้านการจัดการพลังงาน',
    ResponsibleUnit: 'ฝ่ายบริหารงานทั่วไป',
    CurrentStatus: 'เริ่มโครงการ — ดูรายละเอียดเต็มได้ที่แท็บ "ปุ๋ยหมักใบไม้"'
  });

  if (createdFtf || createdPl || createdCtg || createdShare || createdKit || createdCal || createdEm || createdCp) {
    Logger.log('เพิ่มโครงการเชื่อมโยงในชีต Projects: ' +
      (createdFtf ? '"ทอดไม่ทิ้ง" ' : '') + (createdPl ? '"Paperless" ' : '') + (createdCtg ? '"ตลับหมึกเก่าเราขอ" ' : '') +
      (createdShare ? '"คัดแยก แบ่งปัน" ' : '') + (createdKit ? '"จากครัวสู่ท้องฟ้า" ' : '') + (createdCal ? '"ปฏิทินเก่าเราขอ"' : ''));
  }

  // อัปเดตสถานะล่าสุดให้ตรงกับข้อมูลจริงทันที (ถ้ามีข้อมูลอยู่แล้ว)
  syncFtfLinkedProject_();
  syncPaperlessLinkedProject_();
  syncCartridgeLinkedProject_();
  syncShareLinkedProject_();
  syncKitchenLinkedProject_();
  syncCalendarLinkedProject_();
  syncEmLinkedProject_();
  syncCpLinkedProject_();
}

/** คำนวณ % ความคืบหน้าตามสัดส่วนวันที่ผ่านไปของปี (ใช้กับแคมเปญต่อเนื่องทั้งปีที่ไม่มี % งานจริงให้วัด) */
function yearProgressPct_() {
  const now = new Date();
  const start = new Date(now.getFullYear(), 0, 1);
  const end = new Date(now.getFullYear(), 11, 31);
  return Math.min(100, Math.max(0, Math.round(((now - start) / (end - start)) * 100)));
}

/**
 * ซิงก์ข้อมูลสดจากแท็บโครงการกลับมาที่แถวโครงการเชื่อมโยงในหน้าภาพรวม (ไม่ผ่าน PIN เพราะเป็นการ sync อัตโนมัติ)
 *  - ข้อความสรุปจากแท็บ (fields.CurrentStatus) บันทึกที่คอลัมน์ AutoSummary เสมอ → แสดงคู่กับข้อมูลที่ผู้ดูแลกรอก
 *  - CurrentStatus / Progress จะถูกเขียนทับ "เฉพาะเมื่อ" StatusSource ไม่ใช่ manual
 *    (ถ้าผู้ดูแลเคยบันทึกความคืบหน้า/แก้สถานะในหน้าภาพรวมแล้ว ข้อมูลของผู้ดูแลจะไม่ถูกทับหาย)
 */
function updateLinkedProjectFields_(projectId, fields) {
  ensureProjectExtraColumns_();
  const found = findRowById_(SHEET_PROJECTS, 'ProjectID', projectId);
  if (!found) return;
  const headers = found.headers;
  const sh = getSheet_(SHEET_PROJECTS);
  const srcCol = headers.indexOf('StatusSource');
  const manual = srcCol !== -1 && String(found.values[srcCol] || '').toLowerCase() === 'manual';
  const write = {};
  if (fields.Progress !== undefined) {
    // ถ้าตั้งเป้าหมายรายเดือนของปีนี้ไว้ → ความคืบหน้า = ร้อยละที่ทำได้เทียบเป้าหมายทั้งปี (แทนสัดส่วนวันที่ผ่านไป)
    const tp = targetProgressFor_(targetKeyOf_(projectId));
    if (tp && tp.hasTarget) fields = Object.assign({}, fields, { Progress: Math.min(100, Math.round(tp.annualPct)) });
  }
  Object.keys(fields).forEach(k => {
    if (k === 'CurrentStatus') write.AutoSummary = fields[k];
    if (manual && (k === 'CurrentStatus' || k === 'Progress' || k === 'StatusCategory' || k === 'CurrentStage')) return;
    write[k] = fields[k];
  });
  write.AutoSyncedAt = new Date();
  Object.keys(write).forEach(key => {
    const col = headers.indexOf(key);
    if (col !== -1) sh.getRange(found.rowIndex, col + 1).setValue(write[key]);
  });
  if (!manual) {
    const updatedCol = headers.indexOf('UpdatedAt');
    if (updatedCol !== -1) sh.getRange(found.rowIndex, updatedCol + 1).setValue(new Date());
  }
}

/** เพิ่มคอลัมน์ใหม่ของชีต Projects ให้ชีตเดิม (ไม่กระทบข้อมูลเดิม) — เรียกซ้ำได้ ทำงานจริงครั้งเดียวต่อการเรียกใช้ */
var PROJECT_COLS_CHECKED_ = false;
function ensureProjectExtraColumns_() {
  if (PROJECT_COLS_CHECKED_) return;
  PROJECT_COLS_CHECKED_ = true;
  const sh = findSheet_(SHEET_PROJECTS);
  if (!sh) return;
  const lastCol = sh.getLastColumn();
  const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  const need = ['AutoSummary', 'AutoSyncedAt', 'StatusSource', 'ArchivedAt'].filter(h => headers.indexOf(h) === -1);
  need.forEach((h, i) => sh.getRange(1, lastCol + i + 1).setValue(h).setFontWeight('bold'));
  if (need.length) invalidateReadCache_();
}

/** โครงการที่เชื่อมกับแท็บ (ใช้ร่วมกันทั้งหน้าภาพรวม/แท็บโครงการ/คลังโครงการ) */
function linkedKeyOf_(projectId) {
  const m = {};
  m[LINKED_PROJECT_FTF_ID] = 'frytofly'; m[LINKED_PROJECT_PL_ID] = 'paperless'; m[LINKED_PROJECT_CTG_ID] = 'cartridge';
  m[LINKED_PROJECT_SHARE_ID] = 'share'; m[LINKED_PROJECT_KIT_ID] = 'kitchen'; m[LINKED_PROJECT_CAL_ID] = 'calendar'; m[LINKED_PROJECT_EM_ID] = 'em'; m[LINKED_PROJECT_CP_ID] = 'compost';
  return m[projectId] || '';
}

/** ซิงก์สถานะแถวโครงการ "ทอดไม่ทิ้ง" ให้ตรงกับข้อมูลจริงล่าสุด — เรียกทุกครั้งที่มีการบันทึกแกลลอนใหม่ */
function syncFtfLinkedProject_() {
  try {
    const s = getFryToFlyData().summary;
    const env = getEnvironmentalSettings_();
    const liters = Math.round(s.totalReturned * env.ftfLitersPerGallon);
    updateLinkedProjectFields_(LINKED_PROJECT_FTF_ID, {
      CurrentStatus: `ผู้เข้าร่วม ${s.totalParticipants} คน • บริจาคแล้ว ${s.totalReturned} แกลลอน (${liters.toLocaleString()} ลิตร) ` +
        `(🥇${s.tierCounts.gold} 🥈${s.tierCounts.silver} 🥉${s.tierCounts.bronze})` + ftfShipStatusText_(),
      Progress: yearProgressPct_()
    });
  } catch (e) { /* ยังไม่มีชีต/ข้อมูล ก็ข้ามไปเงียบๆ */ }
}

/** ซิงก์สถานะแถวโครงการ "Paperless" ให้ตรงกับข้อมูลจริงล่าสุด — เรียกทุกครั้งที่มีการบันทึกข้อมูลรายเดือนใหม่ */
function syncPaperlessLinkedProject_(year) {
  try {
    const pl = getPaperlessBootstrap(year || new Date().getFullYear());
    updateLinkedProjectFields_(LINKED_PROJECT_PL_ID, {
      CurrentStatus: `ลดกระดาษแล้ว ${pl.grandTotalSheetsReduced.toLocaleString()} แผ่น • ` +
        `${pl.departments.length} หน่วยงานร่วมบันทึกข้อมูล`,
      Progress: yearProgressPct_()
    });
  } catch (e) { /* ยังไม่มีชีต/ข้อมูล ก็ข้ามไปเงียบๆ */ }
}

/** ซิงก์สถานะแถวโครงการ "ตลับหมึกเก่าเราขอ" ให้ตรงกับข้อมูลจริงล่าสุด — เรียกทุกครั้งที่มีการบันทึกรายการจำหน่าย/การส่งตลับหมึกใหม่ */
function syncCartridgeLinkedProject_() {
  try {
    const ov = getCartridgeOverview();
    updateLinkedProjectFields_(LINKED_PROJECT_CTG_ID, {
      CurrentStatus: `สะสม ${ov.allTime.totalQuantitySold.toLocaleString()} ตลับ • รายได้ ${ov.allTime.totalAmount.toLocaleString()} บาท • ` +
        `ประหยัดค่าทำลายขยะอันตรายได้ ${ov.allTime.disposalCostAvoided.toLocaleString()} บาท`,
      Progress: yearProgressPct_()
    });
  } catch (e) { /* ยังไม่มีชีต/ข้อมูล ก็ข้ามไปเงียบๆ */ }
}

// ============================= WEB APP ENTRY =============================
function doGet(e) {
  const params = (e && e.parameter) || {};
  if (params.api) return handleApiGet_(params); // จุดเชื่อมต่อให้แอปอื่นดึงข้อมูล (JSON) — ดูหัวข้อ "เชื่อมต่อระบบอื่น"
  const template = HtmlService.createTemplateFromFile('Index');
  // รองรับลิงก์ตรงเข้าฟอร์มแคมเปญ เช่น .../exec?page=frytofly&action=form
  template.initialPage = params.page || '';
  template.initialAction = params.action || '';
  template.initialRid = String(params.rid || '').slice(0, 40);
  template.initialToken = String(params.t || '').slice(0, 64);
  template.webAppUrl = ScriptApp.getService().getUrl() || '';
  return template.evaluate()
    .setTitle('ระบบติดตามโครงการสิ่งแวดล้อมโรงพยาบาล')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getFryToFlyFormLink() {
  const base = ScriptApp.getService().getUrl() || '';
  return base + '?page=frytofly&action=form';
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ============================= SHEET HELPERS =============================
function getSS() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

/** หาแท็บชีตแบบยืดหยุ่น (ไม่สนช่องว่างหน้า-หลัง/ตัวพิมพ์เล็ก-ใหญ่) */
function getSheet_(name) {
  return trackSheet_(findSheet_(name));
}
function findSheet_(name) {
  const ss = getSS();
  let sh = ss.getSheetByName(name);
  if (sh) return sh;

  const target = name.trim().toLowerCase();
  const all = ss.getSheets();
  sh = all.find(s => s.getName().trim().toLowerCase() === target);
  if (sh) return sh;

  const available = all.map(s => s.getName()).join(', ');
  throw new Error('ไม่พบชีตชื่อ: ' + name + '  |  ชีตที่มีอยู่ในสเปรดชีตนี้คือ: [' + available + ']');
}

/** รันจาก editor เพื่อตรวจว่าสคริปต์ผูกกับสเปรดชีตไฟล์ไหน */
function listSheetNames() {
  const ss = getSS();
  Logger.log('ไฟล์สเปรดชีตที่สคริปต์นี้ผูกอยู่: ' + ss.getName());
  Logger.log('ลิงก์: ' + ss.getUrl());
  Logger.log('รายชื่อแท็บทั้งหมด: ' + ss.getSheets().map(s => '"' + s.getName() + '"').join(', '));
}

/** อ่านทั้งชีตเป็น array ของ object โดยใช้แถวแรกเป็น header */
// ---- แคชการอ่านชีต (อยู่ในหน่วยความจำของคำขอเดียว: ทุกครั้งที่หน้าเว็บเรียกเซิร์ฟเวอร์ ตัวแปรนี้เริ่มใหม่) ----
// ฟังก์ชันสรุปผล (เช่น รายงานสิ่งแวดล้อม) อ่านชีตเดิมซ้ำหลายสิบครั้ง การแคชช่วยลดเวลาโหลดมาก
// เมื่อมีการเขียนชีตใดๆ (appendRow/setValue/deleteRow ฯลฯ) แคชจะถูกล้างทันที จึงไม่ได้ข้อมูลเก่า
let READ_CACHE_ = {};
function invalidateReadCache_() { READ_CACHE_ = {}; }
const SHEET_WRITE_METHODS_ = { appendRow: 1, deleteRow: 1, deleteRows: 1, insertRowAfter: 1, insertRowBefore: 1, insertRows: 1, insertColumns: 1,
  insertColumnAfter: 1, insertColumnBefore: 1, deleteColumn: 1, deleteColumns: 1, clear: 1, clearContents: 1, setName: 1, sort: 1, moveRows: 1 };
const RANGE_WRITE_METHODS_ = { setValue: 1, setValues: 1, setFormula: 1, setFormulas: 1, clear: 1, clearContent: 1, deleteCells: 1, insertCells: 1,
  setNumberFormat: 1, setNumberFormats: 1, setRichTextValue: 1, setRichTextValues: 1, sort: 1, copyTo: 1, moveTo: 1, merge: 1, breakApart: 1 };
function wrapWriteTracker_(obj, methods, wrapReturn) {
  return new Proxy(obj, {
    get: function (target, prop) {
      const v = target[prop];
      if (typeof v !== 'function') return v;
      return function () {
        if (methods[prop]) invalidateReadCache_();
        const r = v.apply(target, arguments);
        return wrapReturn && wrapReturn[prop] ? wrapReturn[prop](r) : r;
      };
    }
  });
}
function trackSheet_(sh) {
  if (!sh) return sh;
  return wrapWriteTracker_(sh, SHEET_WRITE_METHODS_, { getRange: r => wrapWriteTracker_(r, RANGE_WRITE_METHODS_, null), getDataRange: r => wrapWriteTracker_(r, RANGE_WRITE_METHODS_, null) });
}
function getSheetValuesCached_(sheetName) {
  const key = String(sheetName).trim().toLowerCase();
  if (Object.prototype.hasOwnProperty.call(READ_CACHE_, key)) return READ_CACHE_[key];
  const values = findSheet_(sheetName).getDataRange().getValues();
  READ_CACHE_[key] = values;
  return values;
}
function sheetToObjects_(sheetName) {
  const values = getSheetValuesCached_(sheetName);
  if (values.length < 2) return [];
  const headers = values[0];
  const rows = [];
  for (let r = 1; r < values.length; r++) {
    const row = values[r];
    if (row.every(c => c === '' || c === null)) continue;
    const obj = {};
    headers.forEach((h, i) => { obj[h] = row[i]; });
    obj.__row = r + 1;
    rows.push(obj);
  }
  return rows;
}

function findRowById_(sheetName, idField, idValue) {
  const sh = getSheet_(sheetName);
  const values = sh.getDataRange().getValues();
  const headers = values[0];
  const idCol = headers.indexOf(idField);
  if (idCol === -1) throw new Error('ไม่พบคอลัมน์ ' + idField + ' ในชีต ' + sheetName);
  for (let r = 1; r < values.length; r++) {
    if (values[r][idCol] === idValue) {
      return { rowIndex: r + 1, headers: headers, values: values[r] };
    }
  }
  return null;
}

function toIsoDate_(v) {
  if (!v) return null;
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone() || 'Asia/Bangkok', 'yyyy-MM-dd');
  return String(v);
}

function sha256Hex_(text) {
  const raw = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, text, Utilities.Charset.UTF_8);
  return raw.map(b => ('0' + (b & 0xFF).toString(16)).slice(-2)).join('');
}

function newId_(prefix) {
  return prefix + '-' + Utilities.getUuid().split('-')[0].toUpperCase();
}

// ============================= PUBLIC READ APIs =============================

/** ข้อมูลตั้งต้นสำหรับหน้าภาพรวม */
function getBootstrapData() {
  const departments = sheetToObjects_(SHEET_DEPARTMENTS)
    .filter(d => d.IsActive !== false)
    .sort((a, b) => (a.SortOrder || 0) - (b.SortOrder || 0))
    .map(d => ({ DepartmentID: d.DepartmentID, DepartmentName: d.DepartmentName }));

  // log ล่าสุด (ตามวันที่บันทึกจริง) ของแต่ละโครงการ ใช้ปักหมุดบน Gantt
  const lastLogByProject = {};
  sheetToObjects_(SHEET_LOGS).forEach(l => {
    if (l.IsDeleted || !l.ProjectID || !l.LogDate) return;
    const d = new Date(l.LogDate);
    const prev = lastLogByProject[l.ProjectID];
    if (!prev || d > new Date(prev.LogDate)) {
      lastLogByProject[l.ProjectID] = { LogDate: toIsoDate_(l.LogDate), StatusText: l.StatusText };
    }
  });

  ensureProjectExtraColumns_();
  const targetProgs = {};
  TARGET_KEYS.forEach(k => { try { const t = targetProgressFor_(k); if (t && t.hasTarget) targetProgs[k] = t; } catch (e) { /* ข้าม */ } });
  const projects = sheetToObjects_(SHEET_PROJECTS)
    .filter(p => p.ProjectID && !isTruthyFlag_(p.IsArchived))
    .map(p => {
      const lastLog = lastLogByProject[p.ProjectID];
      return {
        ProjectID: p.ProjectID,
        ProjectName: p.ProjectName,
        Category: p.Category,
        Department: p.Department,
        ResponsibleUnit: p.ResponsibleUnit,
        ResponsiblePerson: p.ResponsiblePerson,
        StartDate: toIsoDate_(p.StartDate),
        TargetDate: toIsoDate_(p.TargetDate),
        CurrentStatus: p.CurrentStatus,
        StatusCategory: p.StatusCategory,
        Progress: Number(p.Progress) || 0,
        CurrentStage: p.CurrentStage,
        Description: p.Description || '',
        LastLogDate: lastLog ? lastLog.LogDate : null,
        LinkKey: linkedKeyOf_(p.ProjectID),
        Target: targetProgs[targetKeyOf_(p.ProjectID)] || null,
        AutoSummary: p.AutoSummary || '',
        AutoSyncedAt: p.AutoSyncedAt ? toIsoDate_(p.AutoSyncedAt) : null,
        StatusSource: String(p.StatusSource || '').toLowerCase() === 'manual' ? 'manual' : 'auto'
      };
    });
  const archivedCount = sheetToObjects_(SHEET_PROJECTS).filter(p => p.ProjectID && isTruthyFlag_(p.IsArchived)).length;

  return { archivedCount: archivedCount, departments: departments, projects: projects, statusCategories: STATUS_CATEGORIES, projectModes: getProjectModes_() };
}

/** รายละเอียดโครงการเดียว + ประวัติความคืบหน้าทั้งหมด */
function getProjectDetail(projectId) {
  const found = findRowById_(SHEET_PROJECTS, 'ProjectID', projectId);
  if (!found) throw new Error('ไม่พบโครงการ ' + projectId);
  const obj = {};
  found.headers.forEach((h, i) => obj[h] = found.values[i]);
  obj.StartDate = toIsoDate_(obj.StartDate);
  obj.TargetDate = toIsoDate_(obj.TargetDate);
  obj.CreatedAt = toIsoDate_(obj.CreatedAt);
  obj.UpdatedAt = toIsoDate_(obj.UpdatedAt);
  obj.LinkKey = linkedKeyOf_(projectId);
  try { const t = targetProgressFor_(targetKeyOf_(projectId)); obj.Target = t && t.hasTarget ? t : null; } catch (e) { obj.Target = null; }
  obj.IsArchived = isTruthyFlag_(obj.IsArchived);
  obj.ArchivedAt = obj.ArchivedAt ? toIsoDate_(obj.ArchivedAt) : null;
  obj.AutoSyncedAt = obj.AutoSyncedAt ? toIsoDate_(obj.AutoSyncedAt) : null;
  obj.StatusSource = String(obj.StatusSource || '').toLowerCase() === 'manual' ? 'manual' : 'auto';

  const logs = sheetToObjects_(SHEET_LOGS)
    .filter(l => l.ProjectID === projectId && !l.IsDeleted)
    .map(l => ({
      LogID: l.LogID,
      LogDate: toIsoDate_(l.LogDate),
      StatusText: l.StatusText,
      StatusCategory: l.StatusCategory,
      Progress: Number(l.Progress) || 0,
      RecordedByName: l.RecordedByName,
      RecordedByUnit: l.RecordedByUnit,
      Stage: l.Stage
    }))
    .sort((a, b) => new Date(b.LogDate) - new Date(a.LogDate));

  return { project: obj, logs: logs };
}

/** ประวัติความคืบหน้าของทุกโครงการ จัดกลุ่มตาม ProjectID (ใช้สร้างรายงาน PDF ฉบับเต็ม) */
function getReportLogsByProject() {
  const map = {};
  sheetToObjects_(SHEET_LOGS).forEach(l => {
    if (l.IsDeleted || !l.ProjectID) return;
    (map[l.ProjectID] = map[l.ProjectID] || []).push({
      LogID: l.LogID,
      LogDate: toIsoDate_(l.LogDate),
      StatusText: l.StatusText,
      StatusCategory: l.StatusCategory,
      Progress: Number(l.Progress) || 0,
      RecordedByName: l.RecordedByName,
      Stage: l.Stage
    });
  });
  Object.keys(map).forEach(k => {
    map[k].sort((a, b) => new Date(b.LogDate) - new Date(a.LogDate));
  });
  return map;
}

// ============================= AUTH =============================

/** ตีความค่า TRUE/FALSE ของ Google Sheets ให้เป็น boolean ไม่ว่าจะเป็น checkbox จริงหรือพิมพ์ข้อความ TRUE/FALSE */
function isTruthyFlag_(v) {
  if (v === true) return true;
  if (typeof v === 'string') return v.trim().toUpperCase() === 'TRUE';
  return false;
}

/**
 * เพิ่มคอลัมน์ IsAdmin ให้ชีต Departments ที่มีอยู่แล้วก่อนอัปเดตนี้ (ถ้ายังไม่มี) โดยไม่กระทบคอลัมน์/ข้อมูลเดิม
 * หน่วยงานเดิมทุกแถวจะได้ IsAdmin = TRUE อัตโนมัติ (คงสิทธิ์เดิมไว้ก่อน) — ผู้ดูแลระบบสามารถเข้าไปแก้ในชีตเอง
 * ภายหลังเพื่อจำกัดว่าให้ฝ่ายไหนมีสิทธิ์ผู้ดูแลระบบบ้าง โดยไม่ต้องแก้โค้ด
 */
function ensureDepartmentIsAdminColumn_() {
  const sh = getSheet_(SHEET_DEPARTMENTS);
  const lastCol = sh.getLastColumn();
  const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  if (headers.indexOf('IsAdmin') !== -1) return; // มีคอลัมน์นี้อยู่แล้ว ไม่ต้องทำอะไร

  const newCol = lastCol + 1;
  sh.getRange(1, newCol).setValue('IsAdmin').setFontWeight('bold');
  const lastRow = sh.getLastRow();
  if (lastRow > 1) {
    const values = [];
    for (let r = 2; r <= lastRow; r++) values.push([true]);
    sh.getRange(2, newCol, values.length, 1).setValues(values);
  }
  Logger.log('เพิ่มคอลัมน์ IsAdmin ในชีต Departments (ตั้งค่าเริ่มต้น = TRUE ทุกหน่วยงานที่มีอยู่แล้ว) ' +
    'หากต้องการจำกัดสิทธิ์ ให้เข้าไปเปลี่ยนค่าในคอลัมน์นี้เป็น FALSE สำหรับหน่วยงานที่ไม่ต้องการให้แก้ไขโครงการได้ โดยไม่ต้องแก้โค้ด');
}

/** ตรวจสอบ PIN ของหน่วยงาน + สิทธิ์ผู้ดูแลระบบ (IsAdmin) คืนค่า {ok, departmentId, departmentName} */
function verifyPin(departmentId, pin) {
  const dept = sheetToObjects_(SHEET_DEPARTMENTS).find(d => d.DepartmentID === departmentId);
  if (!dept) return { ok: false, message: 'ไม่พบหน่วยงาน' };
  const hash = sha256Hex_(String(pin));
  if (hash !== dept.PinHash) return { ok: false, message: 'PIN ไม่ถูกต้อง' };
  if (!isTruthyFlag_(dept.IsAdmin)) {
    return { ok: false, message: 'หน่วยงานนี้ไม่มีสิทธิ์ผู้ดูแลระบบ (ผู้ดูแลระบบสามารถเปิดสิทธิ์ให้ได้ที่คอลัมน์ IsAdmin ในชีต Departments)' };
  }
  return { ok: true, departmentId: dept.DepartmentID, departmentName: dept.DepartmentName };
}

function assertAuthorized_(departmentId, pin) {
  const result = verifyPin(departmentId, pin);
  if (!result.ok) throw new Error(result.message || 'ไม่ได้รับอนุญาต');
  return result;
}

// ============================= WRITE APIs (ต้องมี departmentId + pin) =============================

/** เพิ่มโครงการใหม่ */
function addProject(departmentId, pin, data) {
  assertAuthorized_(departmentId, pin);
  const sh = getSheet_(SHEET_PROJECTS);
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const now = new Date();
  const projectId = newId_('PRJ');

  const rowObj = {
    ProjectID: projectId,
    ProjectName: data.ProjectName || '',
    Category: data.Category || '',
    Description: data.Description || '',
    Department: data.Department || '',
    ResponsiblePerson: data.ResponsiblePerson || '',
    ResponsibleUnit: data.ResponsibleUnit || '',
    StartDate: data.StartDate ? new Date(data.StartDate) : now,
    TargetDate: data.TargetDate ? new Date(data.TargetDate) : '',
    CurrentStatus: data.CurrentStatus || 'เริ่มโครงการ',
    StatusCategory: data.StatusCategory || 'ยังไม่เริ่ม',
    Progress: Number(data.Progress) || 0,
    IsArchived: false,
    CreatedAt: now,
    UpdatedAt: now,
    CurrentStage: data.CurrentStage || 'เริ่มโครงการ'
  };

  const row = headers.map(h => (h in rowObj ? rowObj[h] : ''));
  sh.appendRow(row);
  return { ok: true, projectId: projectId };
}

/** แก้ไขโครงการ (เฉพาะฟิลด์ที่ส่งมาใน data) */
function updateProject(departmentId, pin, projectId, data) {
  assertAuthorized_(departmentId, pin);
  ensureProjectExtraColumns_();
  const sh = getSheet_(SHEET_PROJECTS);
  const found = findRowById_(SHEET_PROJECTS, 'ProjectID', projectId);
  if (!found) throw new Error('ไม่พบโครงการ');

  const headers = found.headers;
  data = Object.assign({}, data);
  delete data.AutoSummary; delete data.AutoSyncedAt;
  if (linkedKeyOf_(projectId) && data.StatusSource === undefined) {
    const cur = k => { const i = headers.indexOf(k); return i === -1 ? '' : found.values[i]; };
    const changed = ['CurrentStatus', 'StatusCategory', 'Progress', 'CurrentStage'].some(k =>
      data[k] !== undefined && String(data[k]) !== String(cur(k)));
    if (changed) data.StatusSource = 'manual';
  }
  Object.keys(data).forEach(key => {
    const col = headers.indexOf(key);
    if (col === -1) return;
    let val = data[key];
    if (val === undefined) return;
    if ((key === 'StartDate' || key === 'TargetDate') && val) val = new Date(val);
    sh.getRange(found.rowIndex, col + 1).setValue(val);
  });
  const updatedCol = headers.indexOf('UpdatedAt');
  if (updatedCol !== -1) sh.getRange(found.rowIndex, updatedCol + 1).setValue(new Date());

  return { ok: true };
}

/** ลบโครงการถาวร (ใช้ระวัง) แนะนำใช้ archiveProject แทน */
function deleteProject(departmentId, pin, projectId) {
  assertAuthorized_(departmentId, pin);
  const sh = getSheet_(SHEET_PROJECTS);
  const found = findRowById_(SHEET_PROJECTS, 'ProjectID', projectId);
  if (!found) throw new Error('ไม่พบโครงการ');
  sh.deleteRow(found.rowIndex);
  return { ok: true };
}

/** เก็บเข้าคลัง (soft delete) / นำออกจากคลัง (archived = false) — ข้อมูลและประวัติความคืบหน้าทั้งหมดยังอยู่ครบ */
function archiveProject(departmentId, pin, projectId, archived) {
  assertAuthorized_(departmentId, pin);
  ensureProjectExtraColumns_();
  invalidateReadCache_();
  const on = (archived !== false);
  return updateProject(departmentId, pin, projectId, { IsArchived: on, ArchivedAt: on ? new Date() : '' });
}

/** คลังโครงการ: รายการโครงการที่เก็บเข้าคลังแล้ว (เปิดดูได้ทุกคน) */
function getArchivedProjects() {
  ensureProjectExtraColumns_();
  const logs = {};
  sheetToObjects_(SHEET_LOGS).forEach(l => {
    if (l.IsDeleted || !l.ProjectID) return;
    const o = logs[l.ProjectID] = logs[l.ProjectID] || { count: 0, last: null };
    o.count++;
    if (l.LogDate && (!o.last || new Date(l.LogDate) > new Date(o.last))) o.last = l.LogDate;
  });
  return sheetToObjects_(SHEET_PROJECTS)
    .filter(p => p.ProjectID && isTruthyFlag_(p.IsArchived))
    .map(p => ({
      ProjectID: p.ProjectID, ProjectName: p.ProjectName, Category: p.Category || '', Description: p.Description || '',
      ResponsibleUnit: p.ResponsibleUnit || '', ResponsiblePerson: p.ResponsiblePerson || '',
      StartDate: toIsoDate_(p.StartDate), TargetDate: toIsoDate_(p.TargetDate),
      StatusCategory: p.StatusCategory || '', Progress: Number(p.Progress) || 0,
      CurrentStatus: p.CurrentStatus || '', AutoSummary: p.AutoSummary || '',
      ArchivedAt: p.ArchivedAt ? toIsoDate_(p.ArchivedAt) : (p.UpdatedAt ? toIsoDate_(p.UpdatedAt) : null),
      LinkKey: linkedKeyOf_(p.ProjectID),
      LogCount: logs[p.ProjectID] ? logs[p.ProjectID].count : 0,
      LastLogDate: logs[p.ProjectID] && logs[p.ProjectID].last ? toIsoDate_(logs[p.ProjectID].last) : null
    }))
    .sort((a, b) => String(b.ArchivedAt || '').localeCompare(String(a.ArchivedAt || '')));
}

/** ผู้ดูแลระบบ: ให้โครงการเชื่อมโยงกลับไปใช้สถานะอัตโนมัติจากข้อมูลในแท็บโครงการ */
function setProjectStatusSource(departmentId, pin, projectId, source) {
  assertAuthorized_(departmentId, pin);
  ensureProjectExtraColumns_();
  invalidateReadCache_();
  const src = source === 'manual' ? 'manual' : 'auto';
  updateProject(departmentId, pin, projectId, { StatusSource: src });
  if (src === 'auto') {
    const f = { frytofly: syncFtfLinkedProject_, paperless: syncPaperlessLinkedProject_, cartridge: syncCartridgeLinkedProject_,
      share: syncShareLinkedProject_, kitchen: syncKitchenLinkedProject_, calendar: syncCalendarLinkedProject_ }[linkedKeyOf_(projectId)];
    if (f) { invalidateReadCache_(); f(); }
  }
  return { ok: true, source: src };
}

/** เพิ่มบันทึกความคืบหน้า และ sync สถานะล่าสุดกลับไปที่ตาราง Projects */
function addProgressLog(departmentId, pin, projectId, logData) {
  const auth = assertAuthorized_(departmentId, pin);
  const shLogs = getSheet_(SHEET_LOGS);
  const headers = shLogs.getRange(1, 1, 1, shLogs.getLastColumn()).getValues()[0];
  const now = new Date();
  const logId = newId_('LOG');

  const rowObj = {
    LogID: logId,
    ProjectID: projectId,
    LogDate: logData.LogDate ? new Date(logData.LogDate) : now,
    StatusText: logData.StatusText || '',
    StatusCategory: logData.StatusCategory || '',
    Progress: Number(logData.Progress) || 0,
    RecordedByName: logData.RecordedByName || auth.departmentName,
    RecordedByUnit: logData.RecordedByUnit || auth.departmentName,
    CreatedAt: now,
    IsDeleted: false,
    Stage: logData.Stage || ''
  };
  const row = headers.map(h => (h in rowObj ? rowObj[h] : ''));
  shLogs.appendRow(row);

  updateProject(departmentId, pin, projectId, {
    CurrentStatus: rowObj.StatusText,
    StatusCategory: rowObj.StatusCategory,
    Progress: rowObj.Progress,
    CurrentStage: rowObj.Stage,
    StatusSource: linkedKeyOf_(projectId) ? 'manual' : undefined
  });

  return { ok: true, logId: logId };
}

/** ลบบันทึกความคืบหน้า (soft delete) */
function deleteProgressLog(departmentId, pin, logId) {
  assertAuthorized_(departmentId, pin);
  const found = findRowById_(SHEET_LOGS, 'LogID', logId);
  if (!found) throw new Error('ไม่พบบันทึก');
  const col = found.headers.indexOf('IsDeleted');
  getSheet_(SHEET_LOGS).getRange(found.rowIndex, col + 1).setValue(true);
  return { ok: true };
}

// ============================= SETUP UTILITY =============================
/**
 * ตั้งค่า PIN ให้หน่วยงาน:
 * 1) แก้ค่า pin ด้านล่าง  2) เลือก generatePinHash > Run  3) คัดลอก hash จาก Logs ไปวางที่คอลัมน์ PinHash
 */
function generatePinHash() {
  const pin = '123456'; // <-- เปลี่ยนเป็นรหัส PIN ที่ต้องการก่อนรัน
  Logger.log('PIN: ' + pin + '  →  Hash: ' + sha256Hex_(pin));
}

// ============================= แคมเปญ "ทอดไม่ทิ้ง" (Fry to Fly) =============================
// เปิดให้บุคลากรลงทะเบียน/บันทึกการบริจาคเองได้ โดยไม่ต้องใช้ PIN

/** สรุป + อันดับผู้เข้าร่วม (รวมยอดต่อคนจาก StaffCode, แต้มคิดจากยอด "คืนแกลลอน" เท่านั้น) */
function getFryToFlyData() {
  const all = sheetToObjects_(SHEET_FTF_LOG);
  const rows = all.filter(r => isVerified_(r.VerifyStatus));
  const pending = all.filter(r => r.StaffCode && String(r.VerifyStatus || '') === VERIFY_PENDING);

  const byStaff = {};
  rows.forEach(r => {
    if (!r.StaffCode) return;
    const key = String(r.StaffCode).trim();
    if (!byStaff[key]) {
      byStaff[key] = { StaffCode: key, StaffName: r.StaffName || '', Phone: r.Phone || '', received: 0, returned: 0, lastAt: r.Timestamp };
    }
    const entry = byStaff[key];
    if (r.Timestamp && (!entry.lastAt || new Date(r.Timestamp) >= new Date(entry.lastAt))) {
      entry.StaffName = r.StaffName || entry.StaffName;
      entry.Phone = r.Phone || entry.Phone;
      entry.lastAt = r.Timestamp;
    }
    const qty = Number(r.Quantity) || 0;
    if (r.ActionType === 'รับแกลลอน') entry.received += qty;
    else if (r.ActionType === 'คืนแกลลอน') entry.returned += qty;
  });

  const participants = Object.values(byStaff).map(p => {
    const points = Math.round(p.returned * FTF_POINTS_PER_GALLON);
    const tier = ftfTierFor_(points);
    return {
      StaffCode: p.StaffCode, StaffName: p.StaffName, Phone: p.Phone,
      Received: p.received, Returned: p.returned, Points: points,
      TierKey: tier.key, TierLabel: tier.label
    };
  }).sort((a, b) => b.Points - a.Points || b.Returned - a.Returned);

  const summary = {
    totalParticipants: participants.length,
    totalReceived: participants.reduce((s, p) => s + p.Received, 0),
    totalReturned: participants.reduce((s, p) => s + p.Returned, 0),
    totalPoints: participants.reduce((s, p) => s + p.Points, 0),
    tierCounts: {
      gold: participants.filter(p => p.TierKey === 'gold').length,
      silver: participants.filter(p => p.TierKey === 'silver').length,
      bronze: participants.filter(p => p.TierKey === 'bronze').length,
      member: participants.filter(p => p.TierKey === 'member').length
    }
  };

  summary.pendingCount = pending.length;
  summary.pendingReturned = pending.filter(r => r.ActionType === 'คืนแกลลอน').reduce((s, r) => s + (Number(r.Quantity) || 0), 0);
  return { participants: participants, summary: summary };
}

/** บันทึกรายการเดี่ยว (รับแกลลอน หรือ คืนแกลลอน) */
function submitFryToFlyLog(staffCode, staffName, phone, actionType, quantity, note) {
  staffCode = String(staffCode || '').trim();
  staffName = String(staffName || '').trim();
  if (!staffCode) throw new Error('กรุณากรอกรหัส 7 หลัก');
  if (!staffName) throw new Error('กรุณากรอกชื่อ-นามสกุล');
  if (actionType !== 'รับแกลลอน' && actionType !== 'คืนแกลลอน') throw new Error('ประเภทรายการไม่ถูกต้อง');
  const qty = Number(quantity);
  if (!qty || qty <= 0) throw new Error('กรุณาระบุจำนวนให้ถูกต้อง');

  ensureVerifyColumns_(SHEET_FTF_LOG);
  const sh = getSheet_(SHEET_FTF_LOG);
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const now = new Date();
  const logId = newId_('FTF');

  const rowObj = {
    LogID: logId, Timestamp: now, StaffCode: staffCode, StaffName: staffName,
    Phone: String(phone || '').trim(), ActionType: actionType, Quantity: qty, Note: String(note || '').trim(),
    VerifyStatus: VERIFY_PENDING
  };
  const row = headers.map(h => (h in rowObj ? rowObj[h] : ''));
  sh.appendRow(row);
  syncFtfLinkedProject_();
  return { ok: true, logId: logId };
}

/** ฟอร์มรวม: บันทึก "รับแกลลอน" และ/หรือ "คืนแกลลอน" พร้อมกัน ต้องมีอย่างน้อย 1 ช่องมากกว่า 0 */
function submitFryToFlyCombined(staffCode, staffName, phone, receivedQty, returnedQty, note) {
  staffCode = String(staffCode || '').trim();
  staffName = String(staffName || '').trim();
  if (!staffCode) throw new Error('กรุณากรอกรหัส 7 หลัก');
  if (!staffName) throw new Error('กรุณากรอกชื่อ-นามสกุล');

  const received = Number(receivedQty) || 0;
  const returned = Number(returnedQty) || 0;
  if (received <= 0 && returned <= 0) {
    throw new Error('กรุณาระบุจำนวนแกลลอนที่รับ หรือจำนวนที่คืน อย่างน้อย 1 ช่อง');
  }

  ensureVerifyColumns_(SHEET_FTF_LOG);
  const sh = getSheet_(SHEET_FTF_LOG);
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const now = new Date();
  const phoneVal = String(phone || '').trim();
  const noteVal = String(note || '').trim();
  const logIds = [];

  function appendOne(actionType, qty) {
    const logId = newId_('FTF');
    const rowObj = {
      LogID: logId, Timestamp: now, StaffCode: staffCode, StaffName: staffName,
      Phone: phoneVal, ActionType: actionType, Quantity: qty, Note: noteVal, VerifyStatus: VERIFY_PENDING
    };
    const row = headers.map(h => (h in rowObj ? rowObj[h] : ''));
    sh.appendRow(row);
    logIds.push(logId);
  }

  if (received > 0) appendOne('รับแกลลอน', received);
  if (returned > 0) appendOne('คืนแกลลอน', returned);

  syncFtfLinkedProject_();
  return { ok: true, logIds: logIds, pending: true, pointsEarned: Math.round(returned * FTF_POINTS_PER_GALLON) };
}

// ============================= โครงการ Paperless (ลดการใช้กระดาษ) =============================
// แต่ละหน่วยงานบันทึกข้อมูลของตัวเองรายเดือน (จำนวนเรื่อง / ปริมาณกระดาษที่ลดลง) ผ่าน PIN เฉพาะของหน่วยงาน
// (สเปรดชีต PaperlessDepartments / PaperlessLog — ดู setupSheets ด้านบน)

/** ข้อมูลนำเข้าจริง ม.ค.-ส.ค. 2569 (37 หน่วยงาน) จากไฟล์เปรียบเทียบปริมาณการลดการใช้กระดาษ
 *  รูปแบบ logs: [เดือน(1-12), จำนวนเรื่อง, ปริมาณกระดาษที่ลดลง(แผ่น)] — ใช้ครั้งเดียวกับ importPaperlessInitialData() */
var PAPERLESS_IMPORT_DATA = [
  { name: 'ฝ่ายบริหารงานทั่วไป', logs: [[1,82,84], [2,116,122], [3,166,272], [4,170,442], [5,128,276], [6,126,225], [7,150,358], [8,137,219]] },
  { name: 'ฝ่ายแพทยศาสตรศึกษาชั้นคลินิก', logs: [[1,18,107], [2,8,36], [3,26,192], [4,14,29], [5,17,34], [6,22,45], [7,35,74], [8,12,23]] },
  { name: 'ฝ่ายทันตกรรม', logs: [[1,16,29], [2,11,23], [3,10,21]] },
  { name: 'ฝ่ายพัฒนาคุณภาพและบริหารความเสี่ยง', logs: [[1,5,8], [2,11,12], [3,18,52], [4,1,4], [5,1,1], [6,2,2], [7,4,52], [8,3,6]] },
  { name: 'ศูนย์ส่งเสริมฟื้นฟูสุขภาพผู้สูงอายุ', logs: [[1,3,3], [2,7,14], [3,1,18], [4,3,27], [5,3,17], [6,6,41], [7,7,41], [8,3,28]] },
  { name: 'ฝ่ายผู้ป่วยสัมพันธ์', logs: [[1,3,3], [6,1,1]] },
  { name: 'ฝ่ายทรัพยากรบุคคล', logs: [[1,3,4], [2,4,34], [3,9,46], [4,16,125], [5,13,88], [6,31,391], [7,25,175], [8,26,172]] },
  { name: 'ฝ่ายยาและเวชภัณฑ์', logs: [[1,3,40], [2,23,185], [3,10,84], [4,6,570], [5,9,56], [6,6,32], [7,5,43], [8,8,40]] },
  { name: 'ฝ่ายเวชศาสตร์ชันสูตร', logs: [[1,2,2], [2,6,7], [4,23,117], [5,22,207], [6,26,248], [7,20,157], [8,26,198]] },
  { name: 'ฝ่ายรังสีวิทยา', logs: [[1,2,18], [2,2,13], [8,1,2]] },
  { name: 'ฝ่ายโภชนาการ', logs: [[1,1,13], [4,1,1], [5,2,6], [6,1,2], [7,1,114], [8,3,7]] },
  { name: 'ฝ่ายการเงินและบัญชี', logs: [[1,1,1], [2,3,20], [5,20,22], [6,27,35], [7,10,24], [8,10,14]] },
  { name: 'ฝ่ายเวชสารสนเทศ', logs: [[1,1,2], [3,1,1], [4,2,10], [6,1,2], [7,7,22], [8,2,3]] },
  { name: 'ฝ่ายสื่อสารองค์กรและเวชนิทัศน์', logs: [[1,1,6], [2,2,126], [3,3,19], [4,5,23], [5,2,15], [6,3,45], [7,4,53], [8,3,37]] },
  { name: 'ฝ่ายการพยาบาล', logs: [[1,1,21], [2,7,18], [3,32,109], [4,32,154], [5,25,93], [6,34,131], [7,44,203], [8,35,205]] },
  { name: 'ฝ่ายเวชศาสตร์ฉุกเฉิน', logs: [[2,2,7], [3,3,8], [4,1,4], [5,3,24], [8,1,2]] },
  { name: 'ศูนย์เครื่องมือแพทย์', logs: [[2,1,14], [3,1,11], [4,5,15], [5,1,1], [7,2,3]] },
  { name: 'ฝ่ายผู้ป่วยนอก', logs: [[2,1,1], [5,2,32], [7,1,37]] },
  { name: 'องค์กรแพทย์', logs: [[2,1,1], [3,1,19], [4,2,4], [5,1,3], [6,1,2]] },
  { name: 'ฝ่ายบริหารอาคารสถานที่และสิ่งแวดล้อม', logs: [[2,4,10], [3,3,6], [4,5,7], [5,4,5], [6,4,7], [7,4,9], [8,4,8]] },
  { name: 'ฝ่ายเทคโนโลยีสารสนเทศ', logs: [[2,8,9], [3,7,9], [4,11,24], [5,8,106], [6,6,24], [7,6,40], [8,6,30]] },
  { name: 'ศูนย์บริหารงานกายภาพอาคารพิเศษ', logs: [[2,1,4], [3,5,7], [4,2,6]] },
  { name: 'ฝ่ายจิตเวช', logs: [[2,1,2], [3,1,1], [7,1,2]] },
  { name: 'ฝ่ายเวชกรรมฟื้นฟู', logs: [[3,1,1], [7,1,1]] },
  { name: 'ฝ่ายวิสัญญีวิทยา', logs: [[5,1,2], [7,1,1], [8,1,1]] },
  { name: 'ฝ่ายพยาธิวิทยา', logs: [[1,0,0], [2,0,0], [3,0,0], [4,0,0], [5,0,0], [6,0,0], [7,0,0], [8,0,0], [9,0,0], [10,0,0], [11,0,0], [12,0,0]] },
  { name: 'ฝ่ายอายุรกรรม', logs: [[3,1,13], [4,3,7], [5,1,8], [6,1,4], [7,3,12], [8,1,1]] },
  { name: 'ฝ่ายศัลยกรรม', logs: [[4,3,3], [5,4,6], [6,4,12], [7,3,12], [8,4,32]] },
  { name: 'ฝ่ายศัลยกรรมกระดูกและข้อ', logs: [[4,1,1], [5,1,2], [6,4,38], [7,1,7]] },
  { name: 'ฝ่ายสูติ-นรีเวชกรรม', logs: [[3,2,8], [4,1,6], [6,1,1], [7,3,11]] },
  { name: 'ฝ่ายกุมารเวชกรรม', logs: [[6,1,1]] },
  { name: 'ฝ่ายจักษุ โสต ศอ นาสิก ลาริงซ์', logs: [[6,3,9]] },
  { name: 'ฝ่ายเวชกรรมสังคม', logs: [[3,8,200], [4,3,15], [5,6,39], [6,9,245], [7,8,55], [8,3,19]] },
  { name: 'ฝ่ายบริหารงานพัสดุและจัดซื้อ', logs: [[5,3,6], [6,10,12], [7,8,9], [8,5,7]] },
  { name: 'ฝ่ายวิชาการและวิจัย', logs: [[3,2,7], [4,3,9], [5,3,4], [6,1,2], [7,2,25], [8,1,1]] },
  { name: 'ฝ่ายบริหารจัดการเบิกจ่ายค่าบริการทางการแพทย์', logs: [[3,4,4], [6,4,64], [7,2,4], [8,4,14]] },
  { name: 'ศูนย์บริการสุขภาพ สาขา', logs: [[3,2,2], [5,5,10], [6,1,1]] },
];

/**
 * นำเข้าข้อมูลจริง ม.ค.-ส.ค. 2569 ครั้งเดียว (รันจาก Apps Script editor: เลือก importPaperlessInitialData > Run)
 * - สร้างหน่วยงานใน PaperlessDepartments (ถ้ายังไม่มีชื่อนี้) PIN เริ่มต้น 123456 ทุกหน่วยงาน
 * - เขียนแถวข้อมูลเข้า PaperlessLog แบบ upsert (ถ้ารันซ้ำจะไม่สร้างแถวซ้ำ)
 */
function importPaperlessInitialData() {
  const year = 2026; // ค.ศ. ตรงกับ พ.ศ. 2569
  const deptSheet = getSheet_(SHEET_PL_DEPTS);
  const existingDepts = sheetToObjects_(SHEET_PL_DEPTS);
  const now = new Date();
  let nextSort = existingDepts.length
    ? Math.max.apply(null, existingDepts.map(d => Number(d.SortOrder) || 0)) + 1
    : 1;

  const nameToId = {};
  existingDepts.forEach(d => { nameToId[d.DepartmentName] = d.DepartmentID; });

  let deptsCreated = 0;
  let logsWritten = 0;
  let logsSkipped = 0;

  PAPERLESS_IMPORT_DATA.forEach(entry => {
    let deptId = nameToId[entry.name];
    if (!deptId) {
      deptId = newId_('PLD');
      deptSheet.appendRow([deptId, entry.name, sha256Hex_('123456'), true, nextSort, now, now]);
      nameToId[entry.name] = deptId;
      nextSort++;
      deptsCreated++;
    }

    entry.logs.forEach(tuple => {
      const month = tuple[0], itemCount = tuple[1], sheetsReduced = tuple[2];
      const existing = sheetToObjects_(SHEET_PL_LOG).find(l =>
        l.DepartmentID === deptId && Number(l.Year) === year && Number(l.Month) === month);
      if (existing) { logsSkipped++; return; }
      const logSheet = getSheet_(SHEET_PL_LOG);
      logSheet.appendRow([
        newId_('PLL'), deptId, entry.name, year, month, itemCount, sheetsReduced,
        'นำเข้าข้อมูลเริ่มต้นจากไฟล์รายงาน', 'ระบบ (นำเข้าข้อมูล)', now, now
      ]);
      logsWritten++;
    });
  });

  Logger.log('หน่วยงานที่สร้างใหม่: ' + deptsCreated + '  |  แถวข้อมูลที่เขียน: ' + logsWritten + '  |  ข้ามเพราะมีอยู่แล้ว: ' + logsSkipped);
  Logger.log('หน่วยงานทั้งหมดได้รับ PIN เริ่มต้น "123456" ควรแจ้งให้แต่ละหน่วยงานเปลี่ยนภายหลัง');
}

/** ตรวจสอบ PIN ของหน่วยงาน Paperless คืนค่า {ok, departmentId, departmentName} */
function verifyPaperlessPin(departmentId, pin) {
  const dept = sheetToObjects_(SHEET_PL_DEPTS).find(d => d.DepartmentID === departmentId);
  if (!dept) return { ok: false, message: 'ไม่พบหน่วยงาน' };
  const hash = sha256Hex_(String(pin));
  if (hash !== dept.PinHash) return { ok: false, message: 'PIN ไม่ถูกต้อง' };
  return { ok: true, departmentId: dept.DepartmentID, departmentName: dept.DepartmentName };
}

function assertPaperlessAuthorized_(departmentId, pin) {
  const result = verifyPaperlessPin(departmentId, pin);
  if (!result.ok) throw new Error(result.message || 'ไม่ได้รับอนุญาต');
  return result;
}

/**
 * ข้อมูลตั้งต้นหน้าภาพรวม Paperless ของปีที่ระบุ (ค.ศ.)
 * คืนค่า: หน่วยงานทั้งหมด + ตารางข้อมูลรายเดือนต่อหน่วยงาน + ยอดรวมต่อหน่วยงาน (เรียงมากไปน้อยตามปริมาณกระดาษที่ลดลง)
 * + ยอดรวมทั้งโรงพยาบาลต่อเดือน + รายชื่อปีที่มีข้อมูล
 */
function getPaperlessBootstrap(year) {
  const departments = plVisibleDepts_();

  const allLogs = sheetToObjects_(SHEET_PL_LOG);

  const years = Array.from(new Set(allLogs.map(l => Number(l.Year)).filter(Boolean))).sort((a, b) => b - a);
  const targetYear = Number(year) || (years[0] || new Date().getFullYear());

  const monthlyTotals = Array.from({ length: 12 }, () => ({ itemCount: 0, sheetsReduced: 0 }));

  const rows = departments.map(d => {
    const months = Array.from({ length: 12 }, () => ({ itemCount: 0, sheetsReduced: 0 }));
    let totalItemCount = 0;
    let totalSheetsReduced = 0;

    allLogs
      .filter(l => l.DepartmentID === d.DepartmentID && Number(l.Year) === targetYear)
      .forEach(l => {
        const m = Number(l.Month);
        if (m < 1 || m > 12) return;
        const itemCount = Number(l.ItemCount) || 0;
        const sheetsReduced = Number(l.SheetsReduced) || 0;
        months[m - 1] = { itemCount: itemCount, sheetsReduced: sheetsReduced };
        totalItemCount += itemCount;
        totalSheetsReduced += sheetsReduced;
        monthlyTotals[m - 1].itemCount += itemCount;
        monthlyTotals[m - 1].sheetsReduced += sheetsReduced;
      });

    return {
      DepartmentID: d.DepartmentID,
      DepartmentName: d.DepartmentName,
      months: months,
      totalItemCount: totalItemCount,
      totalSheetsReduced: totalSheetsReduced
    };
  });

  rows.sort((a, b) => b.totalSheetsReduced - a.totalSheetsReduced);
  rows.forEach((r, i) => { r.rank = i + 1; });

  const grandTotalItemCount = rows.reduce((s, r) => s + r.totalItemCount, 0);
  const grandTotalSheetsReduced = rows.reduce((s, r) => s + r.totalSheetsReduced, 0);

  return {
    year: targetYear,
    years: years.length ? years : [targetYear],
    monthNames: PL_MONTH_NAMES,
    departments: rows,
    monthlyTotals: monthlyTotals,
    grandTotalItemCount: grandTotalItemCount,
    grandTotalSheetsReduced: grandTotalSheetsReduced,
    useDepartments: isDeptMode_('paperless')
  };
}

/** ข้อมูลรายละเอียดของหน่วยงานเดียวในปีที่ระบุ (ใช้หน้ารายงานสรุปแต่ละหน่วยงาน) */
function getPaperlessDepartmentDetail(departmentId, year) {
  const dept = sheetToObjects_(SHEET_PL_DEPTS).find(d => d.DepartmentID === departmentId);
  if (!dept) throw new Error('ไม่พบหน่วยงาน ' + departmentId);

  const targetYear = Number(year) || new Date().getFullYear();
  const months = Array.from({ length: 12 }, (_, i) => ({ month: i + 1, monthName: PL_MONTH_NAMES[i], itemCount: 0, sheetsReduced: 0, note: '', recordedByName: '', updatedAt: null }));

  sheetToObjects_(SHEET_PL_LOG)
    .filter(l => l.DepartmentID === departmentId && Number(l.Year) === targetYear)
    .forEach(l => {
      const m = Number(l.Month);
      if (m < 1 || m > 12) return;
      months[m - 1] = {
        month: m,
        monthName: PL_MONTH_NAMES[m - 1],
        itemCount: Number(l.ItemCount) || 0,
        sheetsReduced: Number(l.SheetsReduced) || 0,
        note: l.Note || '',
        recordedByName: l.RecordedByName || '',
        updatedAt: toIsoDate_(l.UpdatedAt)
      };
    });

  const totalItemCount = months.reduce((s, m) => s + m.itemCount, 0);
  const totalSheetsReduced = months.reduce((s, m) => s + m.sheetsReduced, 0);

  return {
    departmentId: dept.DepartmentID,
    departmentName: dept.DepartmentName,
    year: targetYear,
    months: months,
    totalItemCount: totalItemCount,
    totalSheetsReduced: totalSheetsReduced
  };
}

/**
 * บันทึก/แก้ไขข้อมูลประจำเดือนของหน่วยงาน (upsert ตาม DepartmentID+Year+Month)
 * รูปแบบสิทธิ์เหมือนโครงการอื่นๆ (Projects/ทอดไม่ทิ้ง): ผู้ที่ล็อกอินด้วย PIN ของหน่วยงานใดก็ได้
 * (เช่น ฝ่ายบริหารงานทั่วไป ทำหน้าที่เป็นผู้ดูแลระบบ) สามารถบันทึก/แก้ไขข้อมูลของหน่วยงานปลายทาง
 * (targetDepartmentId) ได้ ไม่จำกัดเฉพาะหน่วยงานของตัวเอง — departmentId/pin ใช้ยืนยันตัวตนผู้บันทึกเท่านั้น
 */
function submitPaperlessLog(departmentId, pin, targetDepartmentId, year, month, itemCount, sheetsReduced, note, recordedByName) {
  // โหมดไม่ใช้หน่วยงาน: ผู้ดูแลระบบ (Departments/IsAdmin) บันทึกยอดรวมทั้งโรงพยาบาลเอง; โหมดใช้หน่วยงาน: ใช้ PIN ของหน่วยงานเหมือนเดิม
  const useDepts = isDeptMode_('paperless');
  const auth = useDepts ? assertPaperlessAuthorized_(departmentId, pin) : assertAuthorized_(departmentId, pin);
  if (!useDepts) { ensurePaperlessAllDept_(); targetDepartmentId = PL_ALL_ID; }

  const targetDept = sheetToObjects_(SHEET_PL_DEPTS).find(d => d.DepartmentID === targetDepartmentId);
  if (!targetDept) throw new Error('ไม่พบหน่วยงานปลายทางที่ต้องการบันทึกข้อมูล');

  const y = Number(year);
  const m = Number(month);
  if (!y || m < 1 || m > 12) throw new Error('ปีหรือเดือนไม่ถูกต้อง');

  const itemCountVal = Math.max(0, Number(itemCount) || 0);
  const sheetsReducedVal = Math.max(0, Number(sheetsReduced) || 0);
  const noteVal = String(note || '').trim();
  const recordedByVal = String(recordedByName || '').trim() || auth.departmentName;
  const now = new Date();

  const sh = getSheet_(SHEET_PL_LOG);
  const values = sh.getDataRange().getValues();
  const headers = values[0];
  const idx = {
    DepartmentID: headers.indexOf('DepartmentID'),
    Year: headers.indexOf('Year'),
    Month: headers.indexOf('Month')
  };

  let foundRow = -1;
  for (let r = 1; r < values.length; r++) {
    if (values[r][idx.DepartmentID] === targetDepartmentId &&
        Number(values[r][idx.Year]) === y &&
        Number(values[r][idx.Month]) === m) {
      foundRow = r + 1;
      break;
    }
  }

  if (foundRow !== -1) {
    const rowObj = {
      ItemCount: itemCountVal, SheetsReduced: sheetsReducedVal,
      Note: noteVal, RecordedByName: recordedByVal, UpdatedAt: now
    };
    Object.keys(rowObj).forEach(key => {
      const col = headers.indexOf(key);
      if (col !== -1) sh.getRange(foundRow, col + 1).setValue(rowObj[key]);
    });
    syncPaperlessLinkedProject_(y);
    return { ok: true, mode: 'updated' };
  }

  const logId = newId_('PLL');
  const rowObj = {
    LogID: logId, DepartmentID: targetDepartmentId, DepartmentName: targetDept.DepartmentName,
    Year: y, Month: m, ItemCount: itemCountVal, SheetsReduced: sheetsReducedVal,
    Note: noteVal, RecordedByName: recordedByVal, CreatedAt: now, UpdatedAt: now
  };
  const row = headers.map(h => (h in rowObj ? rowObj[h] : ''));
  sh.appendRow(row);
  syncPaperlessLinkedProject_(y);
  return { ok: true, mode: 'created', logId: logId };
}

/**
 * รายงานสรุปประจำเดือนแบบสะสม (ม.ค. .. เดือนที่เลือก) เรียงหน่วยงานที่ลดกระดาษได้มากที่สุดก่อน
 */
function getPaperlessMonthlyReportData(year, month) {
  const targetYear = Number(year) || new Date().getFullYear();
  const targetMonth = Math.min(12, Math.max(1, Number(month) || 12));

  const departments = plVisibleDepts_();

  const allLogs = sheetToObjects_(SHEET_PL_LOG)
    .filter(l => Number(l.Year) === targetYear && Number(l.Month) >= 1 && Number(l.Month) <= targetMonth);

  const rows = departments.map(d => {
    const deptLogs = allLogs.filter(l => l.DepartmentID === d.DepartmentID);
    const totalItemCount = deptLogs.reduce((s, l) => s + (Number(l.ItemCount) || 0), 0);
    const totalSheetsReduced = deptLogs.reduce((s, l) => s + (Number(l.SheetsReduced) || 0), 0);
    const monthsReported = deptLogs.length;
    return {
      DepartmentID: d.DepartmentID,
      DepartmentName: d.DepartmentName,
      cumulativeItemCount: totalItemCount,
      cumulativeSheetsReduced: totalSheetsReduced,
      monthsReported: monthsReported
    };
  });

  rows.sort((a, b) => b.cumulativeSheetsReduced - a.cumulativeSheetsReduced);
  rows.forEach((r, i) => { r.rank = i + 1; });

  return {
    year: targetYear,
    month: targetMonth,
    monthName: PL_MONTH_NAMES[targetMonth - 1],
    monthNames: PL_MONTH_NAMES,
    rows: rows,
    grandTotalItemCount: rows.reduce((s, r) => s + r.cumulativeItemCount, 0),
    grandTotalSheetsReduced: rows.reduce((s, r) => s + r.cumulativeSheetsReduced, 0)
  };
}

/**
 * ตั้งค่า PIN ให้หน่วยงาน Paperless:
 * 1) แก้ค่า pin ด้านล่าง  2) เลือก generatePaperlessPinHash > Run  3) คัดลอก hash ไปวางที่คอลัมน์ PinHash ของ PaperlessDepartments
 */
function generatePaperlessPinHash() {
  const pin = '123456'; // <-- เปลี่ยนเป็นรหัส PIN ที่ต้องการก่อนรัน
  Logger.log('PIN: ' + pin + '  →  Hash: ' + sha256Hex_(pin));
}

// ============================= โครงการ "ตลับหมึกเก่าเราขอ" (รีไซเคิลตลับหมึกพิมพ์) =============================
// ทุกหน่วยงานส่งตลับหมึกใช้แล้วเข้าจุดรับได้โดยไม่ต้องมี PIN (CartridgeContributions)
// โรงพยาบาลรวบรวมแล้วขายให้ร้านรับซื้อเป็นรอบๆ ราคา/จำนวนแต่ละรอบไม่เท่ากัน มีผู้ซื้อหลายราย
// บันทึก/แก้ไข/ลบรายการจำหน่ายได้เฉพาะผู้ดูแลระบบ (ต้องมีสิทธิ์ IsAdmin เหมือนโครงการอื่นๆ)

/** รายชื่อหน่วยงานตั้งต้นสำหรับฟอร์ม "ส่งตลับหมึกเข้าโครงการ" (เพิ่ม/แก้ไขเพิ่มเติมได้ตรงในชีต CartridgeDepartments ไม่ต้องแก้โค้ด) */
var CARTRIDGE_DEPARTMENT_NAMES = [
  "หน่วยอุบัติเหตุและฉุกเฉิน ER",
  "ฝ่ายผู้ป่วยสัมพันธ์",
  "ฝ่ายรังสีวิทยา",
  "ตึกอัยยิกาเจ้า ARI",
  "ฝ่ายเวชกรรมฟื้นฟู",
  "ฝ่ายเวชศาสตร์ชันสูตร",
  "ฝ่ายยาและเวชภัณฑ์",
  "ศูนย์เครื่องมือแพทย์",
  "หน่วยจ่ายกลาง + โรงซักฟอก",
  "หอผู้ป่วย ผ่าตัดเล็กและห้องส่องกล้อง",
  "ห้องสวนหัวใจ Cath Lab",
  "หอผู้ป่วย ตึกผ่าตัดใหญ่",
  "ฝ่ายวิสัญญีวิทยา",
  "หอผู้ป่วย ตึกคลอด",
  "หอผู้ป่วย หน่วยไอซียูศัลยกรรม",
  "หอผู้ป่วย หน่วยไอซียูอายุรกรรม",
  "หอผู้ป่วยวิกฤติทารกแรกเกิด/NICU",
  "หอผู้ป่วย หน่วยไตเทียม",
  "หอผู้ป่วย แผลไหม้ - น้ำร้อนลวก (14D)",
  "ฝ่ายเวชสารสนเทศ (สถิติ)",
  "ฝ่ายเทคโนโลยีสารสนเทศ",
  "ฝ่ายบริหารจัดการเบิกจ่ายค่าบริการทางการแพทย์",
  "ศูนย์ประกันสังคม",
  "ฝ่ายการเงินและบัญชี",
  "ฝ่ายบริหารงานทั่วไป",
  "สำนักงานฝ่ายการพยาบาล",
  "หอผู้ป่วยสูติกรรม (สามัญ) ชั้น 18A",
  "หอผู้ป่วยนรีเวชกรรม (สามัญ) (ชั้น 18B)",
  "หอผู้ป่วยสามัญกุมารเวชกรรม ชั้น 18D",
  "หอผู้ป่วยอายุรกรรมชาย 3 (19A)",
  "หอผู้ป่วยสามัญศัลยกรรมกระดูกและข้อ (สามัญ) ชั้น 19B",
  "หอผู้ป่วยศัลยกรรมชาย (สามัญ) ชั้น19C",
  "หอผู้ป่วยหญิง (สามัญ) ชั้น 19D",
  "หอผู้ป่วยเคมีบำบัด (สามัญ)/Day care ชั้น 21D",
  "หอผู้ป่วยอายุรกรรมหญิง 1 (สามัญ) ชั้น 22A",
  "หอผู้ป่วยอายุกรรมชาย 1 (สามัญ) ชั้น 22B",
  "หอผู้ป่วยสามัญอายุรกรรมชาย 2 (22C)",
  "หอผู้ป่วยสามัญอายุรกรรมหญิง 2 (22D)",
  "หอผู้ป่วย พิเศษสูติกรรม (23A)",
  "หอผู้ป่วย พิเศษอายุรกรรม-(24A)",
  "หอผู้ป่วย พิแศษศัลยกรรมกระดูกและข้อ (25A)",
  "หอผู้ป่วย พิเศษศัลยกรรม (25B)",
  "หอผู้ป่วย พิเศษทั่วไป (26B)",
  "เวรเปล",
  "ฝ่ายวิชาการและงานวิจัย",
  "ฝ่ายทรัพยากรบุคคล",
  "ฝ่ายพัฒนาคุณภาพและบริหารความเสี่ยง",
  "งานอาชีวอนามัย",
  "งานระบาด",
  "ศูนย์ส่งเสริมแลพัฒนาการวิจัย",
  "งานชุมชน",
  "ฝ่ายทันตกรรม",
  "บริจาคโลหิต",
  "ฝ่ายโภชนาการ",
  "หน่วยซ่อมบำรุง ฝ่ายอาคารฯ",
  "หน่วยธุรการทั่วไป ฝ่ายอาคารฯ",
  "คนงานขยะ",
  "ฝ่ายบริหารงานพัสดุและจัดซื้อ",
  "ฝ่ายพยาธิวิทยากายวิภาค",
  "ตึกสว่างวัฒนา",
  "ฝ่ายแพทยศาสตร์ศึกษาชั้นคลินิก",
  "ห้องสมุด",
  "บริษัทรักษาความปลอดภัย",
  "ทีมขนย้ายขยะ",
  "บริษัท แม่บ้านมืออาชีพ จำกัด",
  "งานเวชศาสตร์ชันสูตร ศูนย์ส่งเสริมสุขภาพผู้สูงอายุ",
  "รพ.ส่วนกลาง",
  "มูลนิธิ",
  "หน่วยบริการโทรศัพท์และสื่อสาร",
  "เรือนน้ำ + เรือนพักพยาบาล",
  "หน่วยงาน N-Health",
  "หน่วยงานช่าง FMU",
  "หน่วยป้องกันและควบคุมการติดเชื้อในร.พ",
  "หอผู้ป่วยสามัญศัลยกรรมอุบัติเหตุชาย (Trauma ward) (20B)",
  "หอผู้ป่วยสามัญศัลยกรรมชาย (20C)",
  "หอผู้ป่วยสามัญ EENT (18 C)",
  "หอผู้ป่วยวิกฤตกุมารเวชกรรม PICU",
  "หอผู้ป่วยวิกฤตโรคหัวใจและหลอดเลือด (CCU)",
  "ฝ่ายจิตเวช",
  "ศูนย์บำบัดด้วยออกซิเจนแรงดันสูง"
];

/** เติมรายชื่อหน่วยงานตั้งต้นในชีต CartridgeDepartments (เฉพาะชื่อที่ยังไม่มี) — ทำงานซ้ำได้อย่างปลอดภัย */
function ensureCartridgeDepartmentsSeeded_() {
  const sh = getSheet_(SHEET_CTG_DEPTS);
  const existing = sheetToObjects_(SHEET_CTG_DEPTS);
  const existingNames = existing.map(d => d.DepartmentName);
  let nextSort = existing.length ? Math.max.apply(null, existing.map(d => Number(d.SortOrder) || 0)) + 1 : 1;
  let added = 0;
  CARTRIDGE_DEPARTMENT_NAMES.forEach(name => {
    if (existingNames.indexOf(name) !== -1) return;
    sh.appendRow([newId_('CTGD'), name, nextSort, true]);
    nextSort++;
    added++;
  });
  if (added) Logger.log('เพิ่มหน่วยงานในชีต CartridgeDepartments: ' + added + ' หน่วยงาน');
}

/**
 * เพิ่มคอลัมน์ Status ให้ชีต CartridgeContributions ที่มีอยู่แล้วก่อนอัปเดตนี้ (ถ้ายังไม่มี) โดยไม่กระทบคอลัมน์/ข้อมูลเดิม
 * แถวเดิมทุกแถว (คือรายการที่หน่วยงานแจ้งส่งเข้าโครงการจริงผ่านฟอร์ม ก่อนจะมีคอลัมน์ Status) จะได้ Status = "รอจำหน่าย"
 * เพราะยังไม่เคยถูกนำไปจับคู่กับรายการจำหน่ายใดเลย — ผู้ดูแลระบบสามารถเข้าไปแก้สถานะเป็นรายตัวได้ภายหลัง (ในชีตเอง หรือในหน้าเว็บ)
 */
function ensureCartridgeContributionStatusColumn_() {
  const sh = getSheet_(SHEET_CTG_CONTRIB);
  const lastCol = sh.getLastColumn();
  const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  if (headers.indexOf('Status') !== -1) return; // มีคอลัมน์นี้อยู่แล้ว ไม่ต้องทำอะไร

  const newCol = lastCol + 1;
  sh.getRange(1, newCol).setValue('Status').setFontWeight('bold');
  const lastRow = sh.getLastRow();
  if (lastRow > 1) {
    const values = [];
    for (let r = 2; r <= lastRow; r++) values.push([CTG_STATUS_PENDING]);
    sh.getRange(2, newCol, values.length, 1).setValues(values);
  }
  Logger.log('เพิ่มคอลัมน์ Status ในชีต CartridgeContributions (ตั้งค่าเริ่มต้น = "' + CTG_STATUS_PENDING + '" สำหรับแถวที่มีอยู่แล้ว)');
}

/** เติมค่าตั้งต้นในชีต CartridgeSettings (Key/Value) ถ้ายังไม่มี — ทำงานซ้ำได้อย่างปลอดภัย ไม่ทับค่าที่ผู้ดูแลระบบแก้ไปแล้ว */
function ensureCartridgeSettingsSeeded_() {
  const sh = getSheet_(SHEET_CTG_SETTINGS);
  const existing = sheetToObjects_(SHEET_CTG_SETTINGS);
  const existingKeys = existing.map(r => r.Key);
  const defaults = [
    ['WeightPerUnitKg', CARTRIDGE_WEIGHT_PER_UNIT_KG],
    ['DisposalCostPerKg', CARTRIDGE_DISPOSAL_COST_PER_KG]
  ];
  defaults.forEach(pair => {
    if (existingKeys.indexOf(pair[0]) !== -1) return;
    sh.appendRow(pair);
  });
}

/**
 * อ่านค่าที่ปรับได้ของโครงการตลับหมึกจากชีต CartridgeSettings (น้ำหนักต่อตลับ, ราคาค่าทำลายขยะ/กก.)
 * ถ้าชีตยังไม่มีหรือยังไม่ได้ตั้งค่า จะใช้ค่าเริ่มต้น (CARTRIDGE_WEIGHT_PER_UNIT_KG / CARTRIDGE_DISPOSAL_COST_PER_KG) แทน
 * แก้ค่าได้โดยเข้าไปเปลี่ยนในชีต CartridgeSettings เอง ไม่ต้องแก้โค้ด
 */
function getCartridgeSettings_() {
  let weightPerUnitKg = CARTRIDGE_WEIGHT_PER_UNIT_KG;
  let disposalCostPerKg = CARTRIDGE_DISPOSAL_COST_PER_KG;
  try {
    const rows = sheetToObjects_(SHEET_CTG_SETTINGS);
    rows.forEach(r => {
      if (r.Key === 'WeightPerUnitKg' && r.Value !== '' && r.Value !== undefined && r.Value !== null && !isNaN(Number(r.Value))) {
        weightPerUnitKg = Number(r.Value);
      }
      if (r.Key === 'DisposalCostPerKg' && r.Value !== '' && r.Value !== undefined && r.Value !== null && !isNaN(Number(r.Value))) {
        disposalCostPerKg = Number(r.Value);
      }
    });
  } catch (e) { /* ยังไม่มีชีต CartridgeSettings — ใช้ค่าเริ่มต้น */ }
  return { weightPerUnitKg: weightPerUnitKg, disposalCostPerKg: disposalCostPerKg };
}

/** เติมค่าตั้งต้นในชีต EnvironmentalSettings (Key/Value) ถ้ายังไม่มี — ทำงานซ้ำได้อย่างปลอดภัย ไม่ทับค่าที่ผู้ดูแลระบบแก้ไปแล้ว */
function ensureEnvironmentalSettingsSeeded_() {
  const sh = getSheet_(SHEET_ENV_SETTINGS);
  const existing = sheetToObjects_(SHEET_ENV_SETTINGS);
  const existingKeys = existing.map(r => r.Key);
  const defaults = [
    ['PaperSheetsPerTree', ENV_PAPER_SHEETS_PER_TREE],
    ['PaperCo2PerSheetKg', ENV_PAPER_CO2_PER_SHEET_KG],
    ['FtfLitersPerGallon', ENV_FTF_LITERS_PER_GALLON],
    ['FtfCo2PerLiterKg', ENV_FTF_CO2_PER_LITER_KG],
    ['CartridgeCo2PerKg', ENV_CARTRIDGE_CO2_PER_KG],
    ['Co2AbsorbedPerTreeKgYear', ENV_CO2_ABSORBED_PER_TREE_KG_YEAR],
    ['SharePaperSheetsPerReam', ENV_SHARE_SHEETS_PER_REAM],
    ['KitchenLitersPerCan', ENV_KITCHEN_LITERS_PER_CAN],
    ['KitchenCo2PerLiterKg', ENV_KITCHEN_CO2_PER_LITER_KG],
    ['CalendarTreesPerTon', ENV_CALENDAR_TREES_PER_TON],
    ['CalendarCo2PerKg', ENV_CALENDAR_CO2_PER_KG],
    ['InquiryEmail', '']   // อีเมลผู้รับคำถามจากกล่องสอบถาม (คั่นหลายอีเมลด้วย ,) — ถ้าเว้นว่างจะใช้อีเมลของหน่วยงานผู้ดูแลระบบ
  ];
  defaults.forEach(pair => {
    if (existingKeys.indexOf(pair[0]) !== -1) return;
    sh.appendRow(pair);
  });
}

/**
 * อ่านตัวแปรคำนวณผลกระทบด้านสิ่งแวดล้อมจากชีต EnvironmentalSettings
 * ถ้าชีตยังไม่มีหรือยังไม่ได้ตั้งค่า จะใช้ค่าเริ่มต้น (ENV_* ด้านบน) แทน — แก้ค่าได้โดยเข้าไปเปลี่ยนในชีตเอง ไม่ต้องแก้โค้ด
 */
function getEnvironmentalSettings_() {
  const result = {
    paperSheetsPerTree: ENV_PAPER_SHEETS_PER_TREE,
    paperCo2PerSheetKg: ENV_PAPER_CO2_PER_SHEET_KG,
    ftfLitersPerGallon: ENV_FTF_LITERS_PER_GALLON,
    ftfCo2PerLiterKg: ENV_FTF_CO2_PER_LITER_KG,
    cartridgeCo2PerKg: ENV_CARTRIDGE_CO2_PER_KG,
    co2AbsorbedPerTreeKgYear: ENV_CO2_ABSORBED_PER_TREE_KG_YEAR,
    sharePaperSheetsPerReam: ENV_SHARE_SHEETS_PER_REAM,
    kitchenLitersPerCan: ENV_KITCHEN_LITERS_PER_CAN,
    kitchenCo2PerLiterKg: ENV_KITCHEN_CO2_PER_LITER_KG,
    calendarTreesPerTon: ENV_CALENDAR_TREES_PER_TON,
    calendarCo2PerKg: ENV_CALENDAR_CO2_PER_KG
  };
  const keyMap = {
    PaperSheetsPerTree: 'paperSheetsPerTree',
    PaperCo2PerSheetKg: 'paperCo2PerSheetKg',
    FtfLitersPerGallon: 'ftfLitersPerGallon',
    FtfCo2PerLiterKg: 'ftfCo2PerLiterKg',
    CartridgeCo2PerKg: 'cartridgeCo2PerKg',
    Co2AbsorbedPerTreeKgYear: 'co2AbsorbedPerTreeKgYear',
    SharePaperSheetsPerReam: 'sharePaperSheetsPerReam',
    KitchenLitersPerCan: 'kitchenLitersPerCan',
    KitchenCo2PerLiterKg: 'kitchenCo2PerLiterKg',
    CalendarTreesPerTon: 'calendarTreesPerTon',
    CalendarCo2PerKg: 'calendarCo2PerKg'
  };
  try {
    sheetToObjects_(SHEET_ENV_SETTINGS).forEach(r => {
      const field = keyMap[r.Key];
      if (field && r.Value !== '' && r.Value !== undefined && r.Value !== null && !isNaN(Number(r.Value))) {
        result[field] = Number(r.Value);
      }
    });
  } catch (e) { /* ยังไม่มีชีต EnvironmentalSettings — ใช้ค่าเริ่มต้น */ }
  return result;
}

/** ปัดเศษทศนิยม 2 ตำแหน่ง (ใช้ทั่วไปสำหรับตัวเลขผลกระทบสิ่งแวดล้อม) */
function round2_(n) {
  return Math.round((Number(n) || 0) * 100) / 100;
}

/** รวมยอด Paperless สะสมทุกปี (ไม่กรองเฉพาะปีที่เลือกแบบ getPaperlessBootstrap) — ใช้คำนวณผลกระทบสิ่งแวดล้อมสะสมทั้งหมด */
function getPaperlessAllTimeTotals_() {
  const logs = sheetToObjects_(SHEET_PL_LOG);
  return {
    totalItemCount: logs.reduce((s, r) => s + (Number(r.ItemCount) || 0), 0),
    totalSheetsReduced: logs.reduce((s, r) => s + (Number(r.SheetsReduced) || 0), 0)
  };
}

/**
 * สรุปผลกระทบด้านสิ่งแวดล้อมของโรงพยาบาล รวมทุกโครงการ (Paperless, ทอดไม่ทิ้ง, ตลับหมึกเก่าเราขอ)
 * ใช้กับหน้ารายงานสรุปผลด้านสิ่งแวดล้อมแยกเฉพาะ — แสดงทั้งแยกรายโครงการและภาพรวมรวมกัน
 */
function getEnvironmentalImpactSummary() {
  const env = getEnvironmentalSettings_();

  // --- Paperless: กระดาษที่ลดได้สะสมทุกปี ---
  const pl = getPaperlessAllTimeTotals_();
  const plTreesSaved = pl.totalSheetsReduced / env.paperSheetsPerTree;
  const plCo2SavedKg = pl.totalSheetsReduced * env.paperCo2PerSheetKg;

  // --- ทอดไม่ทิ้ง: น้ำมันใช้แล้วที่เก็บได้สะสม (แกลลอน -> ลิตร) ---
  const ftfSummary = getFryToFlyData().summary;
  const ftfGallons = ftfSummary.totalReturned;
  const ftfLiters = ftfGallons * env.ftfLitersPerGallon;
  const ftfCo2SavedKg = ftfLiters * env.ftfCo2PerLiterKg;

  // --- ตลับหมึกเก่าเราขอ: น้ำหนัก/ยอดสะสม ---
  const ctgAllTime = getCartridgeOverview().allTime;
  const ctgCo2AvoidedKg = ctgAllTime.totalWeightKg * env.cartridgeCo2PerKg;

  // --- คัดแยก แบ่งปัน: กระดาษ/ตลับหมึกที่แบ่งปันไปใช้แทนการซื้อใหม่ ---
  let shareImpact = { paperReams: 0, paperSheets: 0, treesSaved: 0, paperCo2Kg: 0, cartridges: 0, cartridgeWeightKg: 0, cartridgeCo2Kg: 0, totalCo2Kg: 0 };
  let shareKpi = { withdrawalCount: 0, withdrawalValue: 0, withdrawalDeptCount: 0, stockLines: 0, stockValue: 0 };
  try {
    const shareOv = getShareOverview();
    shareImpact = shareOv.impact;
    shareKpi = shareOv.kpi;
  } catch (e) { /* ยังไม่มีชีตของโครงการคัดแยก แบ่งปัน — นับเป็น 0 */ }

  // --- ส่งน้ำมัน/จำหน่ายรายรอบของทอดไม่ทิ้ง ---
  let ftfShip = { rounds: 0, liters: 0, amount: 0 };
  try { const a = getFryToFlyShipments().allTime; ftfShip = { rounds: a.rounds, liters: a.qty, amount: a.amount }; } catch (e) { /* ยังไม่มีข้อมูล */ }

  // --- จากครัวสู่ท้องฟ้า (ฝ่ายโภชนาการ) ---
  let kit = { rounds: 0, liters: 0, cans: 0, amount: 0, avgPrice: 0, co2SavedKg: 0 };
  try {
    const a = getKitchenOverview().allTime;
    kit = { rounds: a.rounds, liters: a.qty, cans: a.cans, amount: a.amount, avgPrice: a.avgPrice, co2SavedKg: round2_(a.qty * env.kitchenCo2PerLiterKg) };
  } catch (e) { /* ยังไม่มีข้อมูล */ }

  // --- ปฏิทินเก่าเราขอ (บริจาคทำอักษรเบรลล์) ---
  let cal = { rounds: 0, donors: 0, totalKg: 0, treesSaved: 0, co2SavedKg: 0 };
  try {
    const a = getCalendarOverview().allTime;
    cal = { rounds: a.rounds, donors: a.donors, totalKg: a.totalKg, treesSaved: a.treesSaved, co2SavedKg: a.co2SavedKg };
  } catch (e) { /* ยังไม่มีข้อมูล */ }

  const totalCo2SavedKg = plCo2SavedKg + ftfCo2SavedKg + ctgCo2AvoidedKg + shareImpact.totalCo2Kg + kit.co2SavedKg + cal.co2SavedKg;
  const treesEquivalentFromCo2 = totalCo2SavedKg / env.co2AbsorbedPerTreeKgYear;

  return {
    settings: env,
    paperless: {
      totalItemCount: pl.totalItemCount,
      totalSheetsReduced: pl.totalSheetsReduced,
      treesSaved: round2_(plTreesSaved),
      co2SavedKg: round2_(plCo2SavedKg)
    },
    fryToFly: {
      totalGallons: ftfGallons,
      totalLiters: round2_(ftfLiters),
      co2SavedKg: round2_(ftfCo2SavedKg),
      totalParticipants: ftfSummary.totalParticipants,
      shipRounds: ftfShip.rounds,
      shipLiters: round2_(ftfShip.liters),
      shipAmount: round2_(ftfShip.amount)
    },
    kitchen: kit,
    calendar: cal,
    cartridge: {
      totalQuantitySold: ctgAllTime.totalQuantitySold,
      totalContributed: ctgAllTime.totalContributed,
      totalWeightKg: ctgAllTime.totalWeightKg,
      co2AvoidedKg: round2_(ctgCo2AvoidedKg),
      disposalCostAvoided: ctgAllTime.disposalCostAvoided
    },
    share: {
      withdrawalCount: shareKpi.withdrawalCount,
      withdrawalValue: shareKpi.withdrawalValue,
      departmentCount: shareKpi.withdrawalDeptCount,
      stockLines: shareKpi.stockLines,
      stockValue: shareKpi.stockValue,
      paperReams: shareImpact.paperReams,
      paperSheets: shareImpact.paperSheets,
      treesSaved: shareImpact.treesSaved,
      cartridges: shareImpact.cartridges,
      cartridgeWeightKg: shareImpact.cartridgeWeightKg,
      co2SavedKg: shareImpact.totalCo2Kg
    },
    overview: {
      totalCo2SavedKg: round2_(totalCo2SavedKg),
      totalCo2SavedTon: round2_(totalCo2SavedKg / 1000),
      treesSavedDirect: round2_(plTreesSaved + shareImpact.treesSaved + cal.treesSaved),
      treesEquivalentFromCo2: round2_(treesEquivalentFromCo2)
    }
  };
}

/**
 * ข้อมูลภาพรวมโครงการตลับหมึกของปีที่ระบุ (ค.ศ.) — ใช้กับหน้าภาพรวม/สรุปรายเดือน/สรุปรายปี
 * คืนค่า: รายชื่อหน่วยงาน (สำหรับฟอร์ม), ยอดรายเดือนของปีที่เลือก, รายการจำหน่ายของปีที่เลือก,
 * สรุปรายปีทุกปีที่มีข้อมูล, อันดับหน่วยงานที่ส่งตลับหมึกเข้าโครงการมากที่สุด (สะสมทุกปี), และยอดสะสมทั้งหมด
 */
function getCartridgeOverview(year) {
  const settings = getCartridgeSettings_();
  const weightPerUnitKg = settings.weightPerUnitKg;
  const disposalCostPerKg = settings.disposalCostPerKg;

  const departments = sheetToObjects_(SHEET_CTG_DEPTS)
    .filter(d => d.IsActive !== false)
    .sort((a, b) => (Number(a.SortOrder) || 0) - (Number(b.SortOrder) || 0))
    .map(d => ({ DepartmentID: d.DepartmentID, DepartmentName: d.DepartmentName }));

  const sales = sheetToObjects_(SHEET_CTG_SALES).filter(s => s.SaleDate);
  const ctgAll = sheetToObjects_(SHEET_CTG_CONTRIB);
  const ctgPendingVerify = ctgAll.filter(c => String(c.VerifyStatus || '') === VERIFY_PENDING);
  const contributions = ctgAll.filter(c => isVerified_(c.VerifyStatus)).map(c => ({
    ContributionID: c.ContributionID,
    DepartmentName: c.DepartmentName,
    ContributionDate: c.ContributionDate,
    Quantity: Number(c.Quantity) || 0,
    Note: c.Note || '',
    CreatedAt: c.CreatedAt,
    // ข้อมูลเก่าก่อนมีคอลัมน์ Status ให้ถือเป็น "รอจำหน่าย" ไปก่อน (ผู้ดูแลระบบปรับได้)
    Status: c.Status || CTG_STATUS_PENDING
  }));

  const years = Array.from(new Set(sales.map(s => new Date(s.SaleDate).getFullYear()).filter(Boolean))).sort((a, b) => b - a);
  const targetYear = Number(year) || (years[0] || new Date().getFullYear());

  const monthlyTotals = Array.from({ length: 12 }, () => ({ quantity: 0, amount: 0 }));
  const salesThisYear = [];

  sales.forEach(s => {
    const d = new Date(s.SaleDate);
    const qty = Number(s.Quantity) || 0;
    const amt = Number(s.Amount) || 0;
    if (d.getFullYear() === targetYear) {
      monthlyTotals[d.getMonth()].quantity += qty;
      monthlyTotals[d.getMonth()].amount += amt;
      salesThisYear.push({
        SaleID: s.SaleID, SaleDate: toIsoDate_(s.SaleDate), BuyerName: s.BuyerName,
        Quantity: qty, Amount: amt, UnitPrice: round2_(Number(s.UnitPrice) > 0 ? Number(s.UnitPrice) : (qty > 0 ? amt / qty : 0)),
        Note: s.Note || '', RecordedByName: s.RecordedByName || ''
      });
    }
  });
  salesThisYear.sort((a, b) => new Date(b.SaleDate) - new Date(a.SaleDate));

  // สรุปรายปี (ทุกปีที่มีข้อมูลรายการจำหน่าย)
  const yearlyMap = {};
  sales.forEach(s => {
    const y = new Date(s.SaleDate).getFullYear();
    if (!y) return;
    if (!yearlyMap[y]) yearlyMap[y] = { year: y, totalQuantity: 0, totalAmount: 0 };
    yearlyMap[y].totalQuantity += Number(s.Quantity) || 0;
    yearlyMap[y].totalAmount += Number(s.Amount) || 0;
  });
  const yearlySummary = Object.values(yearlyMap).map(y => ({
    year: y.year,
    totalQuantity: y.totalQuantity,
    totalAmount: y.totalAmount,
    totalWeightKg: Math.round(y.totalQuantity * weightPerUnitKg * 100) / 100,
    disposalCostAvoided: Math.round(y.totalQuantity * weightPerUnitKg * disposalCostPerKg)
  })).sort((a, b) => b.year - a.year);

  // อันดับหน่วยงานที่ส่งตลับหมึกเข้าโครงการมากที่สุด (สะสมทุกปี ทุกสถานะ)
  // ไม่นับรายการที่ระบบสร้างอัตโนมัติคู่กับรายการจำหน่าย (ขึ้นต้นด้วย "ระบบ") เพราะไม่ใช่การส่งเข้าโครงการจริงของหน่วยงานใดหน่วยงานหนึ่ง
  const contribByDept = {};
  contributions.forEach(c => {
    const name = c.DepartmentName || '(ไม่ระบุหน่วยงาน)';
    if (name.indexOf('ระบบ (') === 0) return;
    contribByDept[name] = (contribByDept[name] || 0) + c.Quantity;
  });
  const contributionLeaderboard = Object.keys(contribByDept)
    .map(name => ({ DepartmentName: name, totalQuantity: contribByDept[name] }))
    .sort((a, b) => b.totalQuantity - a.totalQuantity)
    .map((r, i) => Object.assign(r, { rank: i + 1 }));

  // รายการ "ตลับหมึกส่งเข้าโครงการ" แยกตามสถานะ — รอจำหน่าย / จำหน่ายไม่ได้ (เพื่อแสดงในแท็บย่อย ให้ผู้ดูแลระบบจัดการสถานะได้)
  const toContribRow_ = c => ({
    ContributionID: c.ContributionID, DepartmentName: c.DepartmentName,
    ContributionDate: toIsoDate_(c.ContributionDate), Quantity: c.Quantity,
    Note: c.Note, Status: c.Status
  });
  const pendingContributions = contributions.filter(c => c.Status === CTG_STATUS_PENDING)
    .map(toContribRow_).sort((a, b) => new Date(b.ContributionDate) - new Date(a.ContributionDate));
  const unsellableContributions = contributions.filter(c => c.Status === CTG_STATUS_REJECTED)
    .map(toContribRow_).sort((a, b) => new Date(b.ContributionDate) - new Date(a.ContributionDate));
  const totalPendingQty = pendingContributions.reduce((s, r) => s + r.Quantity, 0);
  const totalUnsellableQty = unsellableContributions.reduce((s, r) => s + r.Quantity, 0);

  // ยอดสะสม "ส่งเข้าโครงการ" ทั้งหมด (ทุกสถานะ) และเฉพาะที่ "จำหน่ายแล้ว" — ตัวเลขหลังนี้ควรเท่ากับยอดจำหน่ายสะสมเสมอ
  // เพราะทุกรายการจำหน่ายจริงจะมีรายการ "ส่งเข้าโครงการ" สถานะจำหน่ายแล้ว รองรับอยู่เสมอ (ดู addCartridgeSale/importCartridgeHistoricalSales)
  const allTimeQuantitySold = sales.reduce((s, r) => s + (Number(r.Quantity) || 0), 0);
  const allTimeAmount = sales.reduce((s, r) => s + (Number(r.Amount) || 0), 0);
  const allTimeWeightKg = allTimeQuantitySold * weightPerUnitKg;
  const allTimeContributed = contributions.reduce((s, r) => s + r.Quantity, 0);
  const allTimeContributedSold = contributions.filter(c => c.Status === CTG_STATUS_SOLD).reduce((s, r) => s + r.Quantity, 0);

  // มุมมอง "การส่งจำหน่ายรายรอบ" (จำนวน + ราคาต่อตลับของแต่ละรอบ) — โครงสร้างเดียวกับทอดไม่ทิ้ง/จากครัวสู่ท้องฟ้า
  const shipItems = sales.map(s => {
    const q = Number(s.Quantity) || 0, a = Number(s.Amount) || 0;
    return { ID: s.SaleID, Date: new Date(s.SaleDate), Buyer: s.BuyerName || '', Quantity: q, Unit: 'ตลับ', Base: q,
      UnitPrice: round2_(Number(s.UnitPrice) > 0 ? Number(s.UnitPrice) : (q > 0 ? a / q : 0)), Amount: a, Note: s.Note || '', RecordedByName: s.RecordedByName || '' };
  });
  const shipments = buildShipView_(shipItems, targetYear);
  shipments.buyerOptions = safeBuyerOptions_();

  return {
    year: targetYear,
    pendingVerifyCount: ctgPendingVerify.length,
    pendingVerifyQty: ctgPendingVerify.reduce((s, c) => s + (Number(c.Quantity) || 0), 0),
    years: years.length ? years : [targetYear],
    monthNames: PL_MONTH_NAMES,
    departments: departments,
    shipments: shipments,
    monthlyTotals: monthlyTotals,
    salesThisYear: salesThisYear,
    yearlySummary: yearlySummary,
    contributionLeaderboard: contributionLeaderboard,
    pendingContributions: pendingContributions,
    unsellableContributions: unsellableContributions,
    totalPendingQty: totalPendingQty,
    totalUnsellableQty: totalUnsellableQty,
    allTime: {
      totalQuantitySold: allTimeQuantitySold,
      totalAmount: allTimeAmount,
      totalWeightKg: Math.round(allTimeWeightKg * 100) / 100,
      disposalCostAvoided: Math.round(allTimeWeightKg * disposalCostPerKg),
      totalContributed: allTimeContributed,
      totalContributedSold: allTimeContributedSold,
      totalPendingQty: totalPendingQty,
      totalUnsellableQty: totalUnsellableQty,
      weightPerUnitKg: weightPerUnitKg,
      disposalCostPerKg: disposalCostPerKg
    }
  };
}

/** ลิงก์ฟอร์ม "ส่งตลับหมึกเข้าโครงการ" สำหรับแชร์/สร้าง QR Code (ไม่ต้องมี PIN) */
function getCartridgeContributionFormLink() {
  const base = ScriptApp.getService().getUrl() || '';
  return base + '?page=cartridge&action=form';
}

/** หน่วยงานแจ้งจำนวนตลับหมึกที่ส่งเข้าจุดรับของโครงการ — เปิดให้ทุกคนบันทึกได้โดยไม่ต้องมี PIN */
function submitCartridgeContribution(departmentName, quantity, contributionDate, note) {
  let name = String(departmentName || '').trim();
  if (!name && !isDeptMode_('cartridge')) name = NO_DEPT_TEXT;
  if (!name) throw new Error('กรุณาเลือกหน่วยงาน');
  const qty = Number(quantity);
  if (!qty || qty <= 0) throw new Error('กรุณาระบุจำนวนตลับหมึกให้ถูกต้อง');

  const sh = getSheet_(SHEET_CTG_CONTRIB);
  const now = new Date();
  const contributionId = newId_('CTGC');
  let dateVal = now;
  if (contributionDate) {
    const d = new Date(contributionDate);
    if (!isNaN(d)) dateVal = d;
  }
  ensureVerifyColumns_(SHEET_CTG_CONTRIB);
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const rowObj = { ContributionID: contributionId, DepartmentName: name, ContributionDate: dateVal, Quantity: qty,
    Note: String(note || '').trim(), CreatedAt: now, Status: CTG_STATUS_PENDING, VerifyStatus: VERIFY_PENDING };
  sh.appendRow(headers.map(h => (h in rowObj ? rowObj[h] : '')));
  return { ok: true, contributionId: contributionId, pending: true };
}

/**
 * เปลี่ยนสถานะรายการ "ส่งตลับหมึกเข้าโครงการ" — ผู้ดูแลระบบ (IsAdmin) เท่านั้น
 * status ต้องเป็นหนึ่งใน: รอจำหน่าย / จำหน่ายแล้ว / จำหน่ายไม่ได้
 */
function updateCartridgeContributionStatus(departmentId, pin, contributionId, status) {
  assertAuthorized_(departmentId, pin);
  const validStatuses = [CTG_STATUS_PENDING, CTG_STATUS_SOLD, CTG_STATUS_REJECTED];
  if (validStatuses.indexOf(status) === -1) throw new Error('สถานะไม่ถูกต้อง');

  const found = findRowById_(SHEET_CTG_CONTRIB, 'ContributionID', contributionId);
  if (!found) throw new Error('ไม่พบรายการนี้');
  const col = found.headers.indexOf('Status');
  if (col === -1) throw new Error('ชีตยังไม่มีคอลัมน์ Status — กรุณารัน setupSheets() อีกครั้ง');
  getSheet_(SHEET_CTG_CONTRIB).getRange(found.rowIndex, col + 1).setValue(status);
  syncCartridgeLinkedProject_();
  return { ok: true };
}

/** เพิ่มรายการ "ส่งเข้าโครงการ" (สถานะจำหน่ายแล้ว) ให้ตรงกับรายการจำหน่ายที่ระบุ — ใช้เชื่อมยอดสะสมให้ตรงกันเสมอ (เรียกจาก add/update/deleteCartridgeSale) */
function syncContributionForSale_(saleId, saleDate, quantity) {
  const sh = getSheet_(SHEET_CTG_CONTRIB);
  const noteTag = 'AUTO:' + saleId;
  const existing = sheetToObjects_(SHEET_CTG_CONTRIB);
  const match = existing.find(c => String(c.Note || '').trim() === noteTag);
  if (match) {
    const found = findRowById_(SHEET_CTG_CONTRIB, 'ContributionID', match.ContributionID);
    if (found) {
      const colDate = found.headers.indexOf('ContributionDate');
      const colQty = found.headers.indexOf('Quantity');
      if (colDate !== -1) sh.getRange(found.rowIndex, colDate + 1).setValue(saleDate);
      if (colQty !== -1) sh.getRange(found.rowIndex, colQty + 1).setValue(quantity);
      return;
    }
  }
  sh.appendRow([newId_('CTGC'), 'ระบบ (บันทึกจากรายการจำหน่าย)', saleDate, quantity, noteTag, new Date(), CTG_STATUS_SOLD]);
}

/** ลบรายการ "ส่งเข้าโครงการ" ที่ผูกกับรายการจำหน่ายที่ระบุ (ถ้ามี) — ใช้เมื่อลบรายการจำหน่าย */
function deleteContributionForSale_(saleId) {
  const noteTag = 'AUTO:' + saleId;
  const existing = sheetToObjects_(SHEET_CTG_CONTRIB);
  const match = existing.find(c => String(c.Note || '').trim() === noteTag);
  if (!match) return;
  const found = findRowById_(SHEET_CTG_CONTRIB, 'ContributionID', match.ContributionID);
  if (found) getSheet_(SHEET_CTG_CONTRIB).deleteRow(found.rowIndex);
}

/** เพิ่มรายการจำหน่ายตลับหมึก (รายรอบ: จำนวน + ราคาต่อตลับ + จำนวนเงิน) — ผู้ดูแลระบบ (IsAdmin) เท่านั้น เพราะเป็นข้อมูลรายได้/การเงินของโรงพยาบาล */
function addCartridgeSale(departmentId, pin, data) {
  const auth = assertAuthorized_(departmentId, pin);
  data = data || {};

  const saleDate = data.SaleDate ? new Date(data.SaleDate) : new Date();
  if (isNaN(saleDate)) throw new Error('วันที่จำหน่ายไม่ถูกต้อง');
  let buyerName = String(data.BuyerName || '').trim();
  if (!buyerName) throw new Error('กรุณากรอกชื่อผู้ซื้อ/ร้านรับซื้อ');
  const quantity = Number(data.Quantity);
  if (!quantity || quantity <= 0) throw new Error('กรุณาระบุจำนวนตลับหมึกให้ถูกต้อง');
  const price = shipResolvePrice_(quantity, data.UnitPrice, data.Amount);

  ensureCartridgeSalesUnitPriceColumn_();
  buyerName = ensureBuyer_(buyerName, 'ตลับหมึก'); // ผู้รับซื้อที่มีอยู่แล้วจะใช้ชื่อเดิมในระบบ ถ้าเป็นรายใหม่จะเพิ่มเข้าทะเบียนอัตโนมัติ
  const now = new Date();
  const saleId = newId_('CTGS');
  appendByHeaders_(SHEET_CTG_SALES, {
    SaleID: saleId, SaleDate: saleDate, BuyerName: buyerName, Quantity: quantity, Amount: price.amount,
    Note: String(data.Note || '').trim(), RecordedByName: auth.departmentName, CreatedAt: now, UpdatedAt: now, UnitPrice: price.unitPrice
  });
  // เพิ่มรายการ "ส่งเข้าโครงการ" สถานะจำหน่ายแล้วให้ตรงกับรายการจำหน่ายนี้ — ทำให้ยอดสะสม "ส่งเข้าโครงการ (จำหน่ายแล้ว)" เท่ากับยอดจำหน่ายสะสมเสมอ
  syncContributionForSale_(saleId, saleDate, quantity);
  syncCartridgeLinkedProject_();
  return { ok: true, saleId: saleId, buyerName: buyerName };
}

/** แก้ไขรายการจำหน่ายตลับหมึก — ผู้ดูแลระบบ (IsAdmin) เท่านั้น */
function updateCartridgeSale(departmentId, pin, saleId, data) {
  assertAuthorized_(departmentId, pin);
  data = data || {};
  ensureCartridgeSalesUnitPriceColumn_();

  const found = findRowById_(SHEET_CTG_SALES, 'SaleID', saleId);
  if (!found) throw new Error('ไม่พบรายการจำหน่ายนี้');
  const headers = found.headers;
  const sh = getSheet_(SHEET_CTG_SALES);
  const cur = {};
  headers.forEach((h, i) => { cur[h] = found.values[i]; });

  const fields = {};
  if (data.SaleDate) {
    const d = new Date(data.SaleDate);
    if (isNaN(d)) throw new Error('วันที่จำหน่ายไม่ถูกต้อง');
    fields.SaleDate = d;
  }
  if (data.BuyerName !== undefined) {
    const buyerName = String(data.BuyerName).trim();
    if (!buyerName) throw new Error('กรุณากรอกชื่อผู้ซื้อ/ร้านรับซื้อ');
    fields.BuyerName = ensureBuyer_(buyerName, 'ตลับหมึก');
  }
  const quantity = (data.Quantity !== undefined) ? Number(data.Quantity) : Number(cur.Quantity);
  if (data.Quantity !== undefined) {
    if (!quantity || quantity <= 0) throw new Error('กรุณาระบุจำนวนตลับหมึกให้ถูกต้อง');
    fields.Quantity = quantity;
  }
  if (data.Amount !== undefined || data.UnitPrice !== undefined) {
    const price = shipResolvePrice_(quantity, data.UnitPrice, data.Amount);
    fields.Amount = price.amount;
    fields.UnitPrice = price.unitPrice;
  }
  if (data.Note !== undefined) fields.Note = String(data.Note).trim();
  fields.UpdatedAt = new Date();

  Object.keys(fields).forEach(key => {
    const col = headers.indexOf(key);
    if (col !== -1) sh.getRange(found.rowIndex, col + 1).setValue(fields[key]);
  });
  // ปรับรายการ "ส่งเข้าโครงการ (จำหน่ายแล้ว)" ที่ผูกกับรายการนี้ให้ตรงกับวันที่/จำนวนล่าสุด
  const updatedRow = sheetToObjects_(SHEET_CTG_SALES).find(r => r.SaleID === saleId);
  if (updatedRow) syncContributionForSale_(saleId, new Date(updatedRow.SaleDate), Number(updatedRow.Quantity) || 0);
  syncCartridgeLinkedProject_();
  return { ok: true };
}

/** ลบรายการจำหน่ายตลับหมึก — ผู้ดูแลระบบ (IsAdmin) เท่านั้น */
function deleteCartridgeSale(departmentId, pin, saleId) {
  assertAuthorized_(departmentId, pin);
  const found = findRowById_(SHEET_CTG_SALES, 'SaleID', saleId);
  if (!found) throw new Error('ไม่พบรายการจำหน่ายนี้');
  getSheet_(SHEET_CTG_SALES).deleteRow(found.rowIndex);
  deleteContributionForSale_(saleId);
  syncCartridgeLinkedProject_();
  return { ok: true };
}

/** ข้อมูลย้อนหลังจริงจากรายงานการจำหน่ายตลับหมึก ก.ค. 2565 – ก.ย. 2569 (15 รอบ/56 รายการ) — ใช้ครั้งเดียวกับ importCartridgeHistoricalSales() */
var CARTRIDGE_SALES_IMPORT_DATA = [
  { date: "2022-07-19", buyer: "นางสาวพิริญาพร บัวลา", qty: 52, amount: 1010 },
  { date: "2022-07-19", buyer: "นายจันทร์เพ็ญ ทะเที่ยง", qty: 25, amount: 440 },
  { date: "2022-08-16", buyer: "นางสาวพิริญาพร บัวลา", qty: 29, amount: 940 },
  { date: "2022-09-05", buyer: "นางสาวพิริญาพร บัวลา", qty: 55, amount: 1090 },
  { date: "2022-09-14", buyer: "นางสาวพิริญาพร บัวลา", qty: 47, amount: 880 },
  { date: "2022-10-17", buyer: "นางสาวพิริญาพร บัวลา", qty: 79, amount: 1520 },
  { date: "2022-11-15", buyer: "ร้านเลิศมงคลคุรุภัณฑ์", qty: 50, amount: 1500 },
  { date: "2022-11-16", buyer: "นางสาวพิริญาพร บัวลา", qty: 51, amount: 970 },
  { date: "2022-12-16", buyer: "นางสาวพิริญาพร บัวลา", qty: 26, amount: 440 },
  { date: "2023-01-10", buyer: "นางสาวพิริญาพร บัวลา", qty: 44, amount: 810 },
  { date: "2023-01-18", buyer: "ร้าน BB.Toner", qty: 54, amount: 2160 },
  { date: "2023-02-20", buyer: "นางสาวพิริญาพร บัวลา", qty: 27, amount: 490 },
  { date: "2023-03-08", buyer: "นางสาวพิริญาพร บัวลา", qty: 68, amount: 1270 },
  { date: "2023-03-08", buyer: "ศรีหมึกโทนเนอร์", qty: 55, amount: 1150 },
  { date: "2023-03-23", buyer: "นายสมาน เจริญธรรม", qty: 29, amount: 1160 },
  { date: "2023-03-31", buyer: "TJ โทนเนอร์", qty: 42, amount: 690 },
  { date: "2023-04-19", buyer: "สมาน เจริญธรรม", qty: 30, amount: 1200 },
  { date: "2023-05-12", buyer: "ร้านเลิศมงคงครุภัณฑ์", qty: 60, amount: 2180 },
  { date: "2023-06-07", buyer: "นางสาวพิริญาพร บัวลา", qty: 44, amount: 820 },
  { date: "2023-07-14", buyer: "ร้าน BB.Toner", qty: 125, amount: 4720 },
  { date: "2023-08-04", buyer: "TJ โทนเนอร์", qty: 71, amount: 1300 },
  { date: "2023-09-14", buyer: "TJ โทนเนอร์", qty: 40, amount: 660 },
  { date: "2023-09-15", buyer: "ร้าน BB.Toner", qty: 61, amount: 2420 },
  { date: "2023-10-25", buyer: "TJ โทนเนอร์", qty: 58, amount: 1100 },
  { date: "2023-10-26", buyer: "ร้าน BB.Toner", qty: 62, amount: 2080 },
  { date: "2023-12-01", buyer: "TJ โทนเนอร์", qty: 74, amount: 1445 },
  { date: "2023-12-21", buyer: "ร้าน BB.Toner", qty: 65, amount: 2250 },
  { date: "2024-01-26", buyer: "ร้าน BB.Toner", qty: 63, amount: 2020 },
  { date: "2024-01-30", buyer: "TJ โทนเนอร์", qty: 76, amount: 1305 },
  { date: "2024-03-15", buyer: "TJ โทนเนอร์", qty: 58, amount: 1090 },
  { date: "2024-04-04", buyer: "ร้าน BB.Toner", qty: 87, amount: 3130 },
  { date: "2024-05-03", buyer: "TJ โทนเนอร์", qty: 33, amount: 1060 },
  { date: "2024-06-07", buyer: "ทองพูล หาวัน", qty: 99, amount: 3690 },
  { date: "2024-06-19", buyer: "TJ โทนเนอร์", qty: 56, amount: 1990 },
  { date: "2024-07-15", buyer: "นายทองพูล หาวัน", qty: 57, amount: 2410 },
  { date: "2024-07-31", buyer: "ร้าน BB Toner", qty: 118, amount: 2565 },
  { date: "2024-08-21", buyer: "TJ โทนเนอร์", qty: 81, amount: 3630 },
  { date: "2024-08-24", buyer: "ร้าน BB Toner", qty: 49, amount: 4583 },
  { date: "2024-10-02", buyer: "TJ โทนเนอร์", qty: 96, amount: 4160 },
  { date: "2024-10-15", buyer: "ร้าน BB Toner", qty: 218, amount: 4795 },
  { date: "2024-11-13", buyer: "TJ โทนเนอร์", qty: 85, amount: 4040 },
  { date: "2025-01-30", buyer: "TJ โทนเนอร์", qty: 89, amount: 3120 },
  { date: "2025-02-07", buyer: "นายทองพูล หาวัน", qty: 23, amount: 750 },
  { date: "2025-03-06", buyer: "ร้าน BB Toner", qty: 178, amount: 5105 },
  { date: "2025-03-14", buyer: "นายทองพูล หาวัน", qty: 5, amount: 200 },
  { date: "2025-03-26", buyer: "TJ โทนเนอร์", qty: 63, amount: 1890 },
  { date: "2025-05-08", buyer: "ร้าน BB Toner", qty: 102, amount: 2700 },
  { date: "2025-06-13", buyer: "นายทองพูล หาวัน", qty: 111, amount: 3010 },
  { date: "2025-07-27", buyer: "นายทองพูล หาวัน", qty: 47, amount: 1430 },
  { date: "2025-08-29", buyer: "นายทองพูล หาวัน", qty: 76, amount: 2990 },
  { date: "2025-11-21", buyer: "ร้าน BB Toner", qty: 20, amount: 200 },
  { date: "2025-11-21", buyer: "นายทศพร บุญกลาง", qty: 369, amount: 21815 },
  { date: "2026-02-06", buyer: "ร้าน BB Toner", qty: 49, amount: 880 },
  { date: "2026-07-10", buyer: "นางสาวกฤษณา โกสิน", qty: 211, amount: 7730 },
  { date: "2026-07-21", buyer: "นายอาซา", qty: 21, amount: 1150 },
  { date: "2026-09-03", buyer: "นายอาซา", qty: 162, amount: 3240 }
];

/**
 * นำเข้าข้อมูลย้อนหลังการจำหน่ายตลับหมึกครั้งเดียว (รันจาก Apps Script editor: เลือก importCartridgeHistoricalSales > Run)
 * เขียนแถวเข้า CartridgeSales — รันซ้ำได้อย่างปลอดภัย (ตรวจรายการซ้ำก่อนเขียนทุกครั้ง)
 */
function importCartridgeHistoricalSales() {
  const sh = getSheet_(SHEET_CTG_SALES);
  const existing = sheetToObjects_(SHEET_CTG_SALES);
  const now = new Date();
  let written = 0, skipped = 0;

  CARTRIDGE_SALES_IMPORT_DATA.forEach(r => {
    const dup = existing.some(e =>
      toIsoDate_(e.SaleDate) === r.date && String(e.BuyerName || '').trim() === r.buyer &&
      Number(e.Quantity) === r.qty && Number(e.Amount) === r.amount);
    if (dup) { skipped++; return; }
    const saleId = newId_('CTGS');
    sh.appendRow([saleId, new Date(r.date), r.buyer, r.qty, r.amount,
      'นำเข้าข้อมูลย้อนหลังจากรายงานการจำหน่ายตลับหมึก', 'ระบบ (นำเข้าข้อมูล)', now, now]);
    // เพิ่มรายการ "ส่งเข้าโครงการ" สถานะจำหน่ายแล้วให้ตรงกับรายการย้อนหลังนี้ด้วย — ทำให้ยอดสะสมส่งเข้าโครงการ (จำหน่ายแล้ว) เท่ากับยอดจำหน่ายสะสม
    syncContributionForSale_(saleId, new Date(r.date), r.qty);
    written++;
  });

  // เผื่อกรณีเคยนำเข้ารายการจำหน่ายไปแล้วก่อนอัปเดตนี้ (ตอนนั้นยังไม่มีการสร้างรายการ "ส่งเข้าโครงการ" คู่กัน)
  // ไล่ตรวจรายการจำหน่ายทุกแถวที่มีอยู่จริงในชีตอีกครั้ง เพื่อเติมรายการ "ส่งเข้าโครงการ (จำหน่ายแล้ว)" ให้ครบทุกแถว — ทำงานซ้ำได้อย่างปลอดภัย
  const allSales = sheetToObjects_(SHEET_CTG_SALES);
  allSales.forEach(s => {
    if (!s.SaleID || !s.SaleDate) return;
    syncContributionForSale_(s.SaleID, new Date(s.SaleDate), Number(s.Quantity) || 0);
  });

  Logger.log('นำเข้ารายการจำหน่ายตลับหมึกย้อนหลัง: เขียนใหม่ ' + written + ' แถว, ข้ามเพราะมีอยู่แล้ว ' + skipped + ' แถว');
  Logger.log('ตรวจสอบและเติมรายการ "ส่งเข้าโครงการ (จำหน่ายแล้ว)" ให้ตรงกับรายการจำหน่ายทั้งหมด: ' + allSales.length + ' แถว');
  syncCartridgeLinkedProject_();
}

// ============================= โครงการ "คัดแยก แบ่งปัน" =============================
// แนวคิด: หน่วยงานคัดแยกเครื่องใช้สำนักงาน/วัสดุที่เกินจำเป็นหรือหมดความจำเป็น ส่งเข้าโครงการ (ผ่านลิงก์ ไม่ต้องมี PIN)
//         ผู้ดูแลระบบตรวจสอบแล้วรับเข้า "คลังรอแบ่งปัน" → หน่วยงานอื่นที่ต้องใช้ขอรับแบ่งปันแทนการจัดซื้อใหม่
//         ทุกการเบิกจะตัดยอดคงเหลืออัตโนมัติ และนับเป็นมูลค่าที่ประหยัดงบจัดซื้อ
//
// ชีตที่เกี่ยวข้อง:
//   - ShareItems       : รายการคงเหลือรอแบ่งปัน (ยอดคงเหลือปัจจุบัน)
//   - ShareSubmissions : รายการที่หน่วยงานแจ้งส่งเข้าโครงการ / ผลสำรวจ (สาธารณะ ไม่ต้องใช้ PIN)
//   - ShareWithdrawals : รายการเบิก/รับแบ่งปัน (ผู้ดูแลระบบบันทึก)
//   - ShareDepartments : รายชื่อหน่วยงานสำหรับฟอร์ม (พิมพ์ชื่อใหม่เองในฟอร์มได้ด้วย)
//   - ShareStockLog    : ประวัติการเปลี่ยนแปลงยอดคงเหลือ (รับเข้า/เบิก/ตรวจนับ/แก้ไข) ไว้ตรวจสอบย้อนหลัง

const SHEET_SHARE_ITEMS = 'ShareItems';
const SHEET_SHARE_SUBS  = 'ShareSubmissions';
const SHEET_SHARE_WD    = 'ShareWithdrawals';
const SHEET_SHARE_DEPTS = 'ShareDepartments';
const SHEET_SHARE_LOG   = 'ShareStockLog';

const LINKED_PROJECT_SHARE_ID = 'PRJ-SHARE-LINK';

const SHARE_STATUS_PENDING  = 'รอตรวจสอบ';
const SHARE_STATUS_ACCEPTED = 'รับเข้าคลังแล้ว';
const SHARE_STATUS_REJECTED = 'ไม่รับเข้า';
const SHARE_STATUS_SURVEY   = 'ผลสำรวจเดิม';
const SHARE_STATUS_NONE     = 'ไม่มีรายการ';

const SHARE_CATEGORIES = ['หมึกพิมพ์/ตลับหมึก', 'ลวดเย็บ/เสียบกระดาษ', 'กระดาษ', 'ปากกา', 'อื่นๆ'];
const SHARE_USAGE_OPTIONS = ['หมดความจำเป็นในการใช้งาน', 'มีรายการเกินจำเป็น (มีปริมาณมากใช้งานไม่ทัน)'];
const SHARE_NO_ITEM_TEXT = 'ไม่มีรายการที่เกินความจำเป็น';

// กระดาษ 1 รีม = กี่แผ่น (ใช้คำนวณผลกระทบสิ่งแวดล้อมของการแบ่งปันกระดาษ) — ปรับได้ที่ชีต EnvironmentalSettings คีย์ SharePaperSheetsPerReam
const ENV_SHARE_SHEETS_PER_REAM = 500;
// หน่วยของตลับหมึก/กล่องโทนเนอร์ที่นับเป็น "1 ตลับ" เมื่อคำนวณน้ำหนัก (ขวดหมึกเติมไม่นับ)
const SHARE_CARTRIDGE_UNITS = ['ตลับ', 'กล่อง'];

/** รายชื่อหน่วยงานที่พบในไฟล์ข้อมูลเดิม (เพิ่มจากรายชื่อตลับหมึก) — ใช้ seed ชีต ShareDepartments */
const SHARE_EXTRA_DEPARTMENT_NAMES = [
 "คลินิกปฐมภูมิ",
 "ห้องคลอด",
 "คลินิกศัลยกรรม",
 "คลินิกจิตเวช",
 "คลินิกตรวจสุขภาพ",
 "18A",
 "อาคารพักเจ้าหน้าที่",
 "คลินิกเวชปฏิบัติ",
 "ฝ่ายเวชสารสนเทศ",
 "หน่วยผลิตยา",
 "จำหน่าย",
 "ห้องศูนย์หัวใจ",
 "ฝ่ายบริหารงานพัสดุฯ",
 "ห้องตรวจสุขภาพ",
 "ฝ่ายบริหารอาคารฯ",
 "20B",
 "OPD",
 "ศูนย์ประกันสุขภาพ สาขา",
 "ฝ่ายบริหารการเบิก-จ่าย"
];

// ----------------------------- helpers -----------------------------
function shareNum_(v) {
  if (v === '' || v === null || v === undefined) return 0;
  const n = Number(v);
  return isNaN(n) ? 0 : n;
}
function shareHasNum_(v) {
  return !(v === '' || v === null || v === undefined || isNaN(Number(v)));
}
function shareText_(s, max) {
  const t = String(s === null || s === undefined ? '' : s).replace(/\s+/g, ' ').trim();
  return max ? t.slice(0, max) : t;
}
function shareKey_(s) { return shareText_(s).toLowerCase(); }
function shareActive_(v) {
  return !(v === false || String(v).trim().toUpperCase() === 'FALSE');
}

/** เดาหมวดหมู่จากชื่อรายการ (ใช้ตอนนำเข้าข้อมูลที่ไม่มีหมวด และเป็นค่าเริ่มต้นในฟอร์ม) */
function guessShareCategory_(name) {
  const n = String(name || '');
  if (/ลวด/.test(n)) return 'ลวดเย็บ/เสียบกระดาษ';
  if (/หมึก|toner|ตลับ/i.test(n)) return 'หมึกพิมพ์/ตลับหมึก';
  if (/ปากกา/.test(n)) return 'ปากกา';
  if (/กระดาษ/.test(n) && !/สติ๊กเกอร์|สติกเกอร์/.test(n)) return 'กระดาษ';
  return 'อื่นๆ';
}
function normalizeShareCategory_(cat, itemName) {
  const c = shareText_(cat);
  return SHARE_CATEGORIES.indexOf(c) !== -1 ? c : guessShareCategory_(itemName);
}

/** สถานะของรายการคงเหลือ คำนวณจากข้อมูลจริง ไม่ต้องกรอกเอง */
function shareStockStatus_(quantity, unitValue) {
  if (!shareHasNum_(quantity)) return 'รอตรวจสอบจำนวน';
  if (Number(quantity) <= 0) return 'หมดแล้ว';
  if (!shareHasNum_(unitValue) || Number(unitValue) <= 0) return 'รอตรวจสอบราคา';
  return 'ระบุราคาแล้ว';
}

/** อ่านชีตแบบเข้มงวด: ถ้ายังไม่มีชีต ให้แจ้งวิธีแก้ที่ชัดเจน */
function shareRows_(sheetName) {
  try {
    return sheetToObjects_(sheetName);
  } catch (e) {
    throw new Error('ยังไม่พบชีต "' + sheetName + '" — กรุณารัน setupSheets() และ importShareProjectData() ใน Apps Script editor ก่อน (' + (e.message || e) + ')');
  }
}

/** เพิ่มแถวโดยจับคู่ตามชื่อหัวคอลัมน์ (ไม่พึ่งลำดับคอลัมน์) */
function shareAppend_(sheetName, obj) {
  const sh = getSheet_(sheetName);
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  sh.appendRow(headers.map(h => (h in obj ? obj[h] : '')));
}

/** แก้ไขหลายฟิลด์ของแถวที่ค้นพบแล้ว (found จาก findRowById_) */
function shareUpdate_(sheetName, found, fields) {
  const sh = getSheet_(sheetName);
  Object.keys(fields).forEach(key => {
    const col = found.headers.indexOf(key);
    if (col !== -1) sh.getRange(found.rowIndex, col + 1).setValue(fields[key]);
  });
}

/** ล็อกกันการแก้ยอดคงเหลือพร้อมกันหลายคน (read-modify-write) */
function withShareLock_(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try { return fn(); } finally { lock.releaseLock(); }
}

function shareIsAdmin_(departmentId, pin) {
  if (!departmentId || !pin) return false;
  try { return !!verifyPin(departmentId, pin).ok; } catch (e) { return false; }
}

function parseShareDate_(v) {
  if (!v) return new Date();
  const d = new Date(v);
  return isNaN(d) ? new Date() : d;
}

function logShareStock_(itemId, itemName, action, before, after, by, note) {
  try {
    shareAppend_(SHEET_SHARE_LOG, {
      LogID: newId_('SHL'), Timestamp: new Date(), ItemID: itemId, ItemName: itemName, Action: action,
      QtyBefore: before === null || before === undefined ? '' : before,
      QtyAfter: after === null || after === undefined ? '' : after,
      By: by || '', Note: note || ''
    });
  } catch (e) { /* ไม่มีชีตประวัติ ก็ไม่ให้ขวางการทำงานหลัก */ }
}

// ----------------------------- setup / seed -----------------------------
/** เติมรายชื่อหน่วยงานในชีต ShareDepartments (เฉพาะชื่อที่ยังไม่มี) — รันซ้ำได้อย่างปลอดภัย */
function ensureShareDepartmentsSeeded_() {
  const existing = sheetToObjects_(SHEET_SHARE_DEPTS);
  const names = existing.map(d => d.DepartmentName);
  let nextSort = existing.length ? Math.max.apply(null, existing.map(d => Number(d.SortOrder) || 0)) + 1 : 1;
  const all = CARTRIDGE_DEPARTMENT_NAMES.concat(SHARE_EXTRA_DEPARTMENT_NAMES);
  let added = 0;
  all.forEach(name => {
    if (names.indexOf(name) !== -1) return;
    names.push(name);
    shareAppend_(SHEET_SHARE_DEPTS, { DepartmentID: newId_('SHD'), DepartmentName: name, SortOrder: nextSort++, IsActive: true });
    added++;
  });
  if (added) Logger.log('เพิ่มหน่วยงานในชีต ShareDepartments: ' + added + ' หน่วยงาน');
}

function shareDepartmentNames_() {
  try {
    return sheetToObjects_(SHEET_SHARE_DEPTS)
      .filter(d => shareActive_(d.IsActive))
      .sort((a, b) => (Number(a.SortOrder) || 0) - (Number(b.SortOrder) || 0))
      .map(d => d.DepartmentName);
  } catch (e) { return []; }
}

// ----------------------------- ผลกระทบสิ่งแวดล้อมของการแบ่งปัน -----------------------------
/**
 * ประมาณการผลกระทบจากสิ่งที่ "แบ่งปันไปใช้แทนการซื้อใหม่" (นับจากรายการเบิกจริง)
 *  - กระดาษ: นับเฉพาะหน่วย "รีม" × แผ่นต่อรีม → ต้นไม้ที่รักษาได้ / CO2 ที่ไม่ต้องผลิตใหม่
 *  - ตลับหมึก: นับเฉพาะหมวดหมึกพิมพ์ที่หน่วยเป็น ตลับ/กล่อง × น้ำหนักต่อตลับ × CO2 ต่อ กก.
 *  เป็นค่าประมาณเชิงภาพรวมจากตัวแปรในชีต EnvironmentalSettings / CartridgeSettings
 */
function getShareImpact_(wds) {
  const env = getEnvironmentalSettings_();
  const ctg = getCartridgeSettings_();
  wds = wds || shareRows_(SHEET_SHARE_WD);
  let reams = 0, cartridges = 0;
  wds.forEach(w => {
    const qty = shareNum_(w.Quantity);
    const unit = shareText_(w.Unit);
    const cat = shareText_(w.Category);
    if (cat === 'กระดาษ' && unit === 'รีม') reams += qty;
    if (cat === 'หมึกพิมพ์/ตลับหมึก' && SHARE_CARTRIDGE_UNITS.indexOf(unit) !== -1) cartridges += qty;
  });
  const sheets = reams * env.sharePaperSheetsPerReam;
  const treesSaved = sheets / env.paperSheetsPerTree;
  const paperCo2Kg = sheets * env.paperCo2PerSheetKg;
  const cartridgeWeightKg = cartridges * ctg.weightPerUnitKg;
  const cartridgeCo2Kg = cartridgeWeightKg * env.cartridgeCo2PerKg;
  return {
    paperReams: round2_(reams),
    paperSheets: Math.round(sheets),
    treesSaved: round2_(treesSaved),
    paperCo2Kg: round2_(paperCo2Kg),
    cartridges: round2_(cartridges),
    cartridgeWeightKg: round2_(cartridgeWeightKg),
    cartridgeCo2Kg: round2_(cartridgeCo2Kg),
    totalCo2Kg: round2_(paperCo2Kg + cartridgeCo2Kg),
    sheetsPerReam: env.sharePaperSheetsPerReam,
    weightPerCartridgeKg: ctg.weightPerUnitKg
  };
}

// ----------------------------- READ: ภาพรวมทั้งหมดของโครงการ -----------------------------
/**
 * ข้อมูลทั้งหมดสำหรับแท็บ "คัดแยก แบ่งปัน" (เรียกครั้งเดียว)
 * ถ้าส่ง departmentId + pin ของผู้ดูแลระบบที่ถูกต้อง จะเห็นชื่อ-ตำแหน่งผู้แจ้งด้วย; ถ้าไม่ส่ง/ไม่ถูกต้อง จะไม่ส่งข้อมูลส่วนบุคคลออกไปเลย
 */
function getShareOverview(adminDepartmentId, adminPin) {
  const isAdmin = shareIsAdmin_(adminDepartmentId, adminPin);

  // ---- รายการเบิก/แบ่งปัน ----
  const wds = shareRows_(SHEET_SHARE_WD).filter(w => w.WithdrawDate).map(w => {
    const qty = shareNum_(w.Quantity);
    const hasPrice = shareHasNum_(w.UnitPrice);
    const price = hasPrice ? Number(w.UnitPrice) : null;
    return {
      WithdrawalID: w.WithdrawalID,
      WithdrawDate: toIsoDate_(w.WithdrawDate),
      Category: w.Category || 'อื่นๆ',
      ItemName: w.ItemName,
      Quantity: qty,
      Unit: w.Unit || '',
      UnitPrice: price,
      LineTotal: hasPrice ? round2_(qty * price) : 0,
      DepartmentName: w.DepartmentName || '(ไม่ระบุหน่วยงาน)',
      Note: w.Note || ''
    };
  }).sort((a, b) => a.WithdrawDate < b.WithdrawDate ? 1 : (a.WithdrawDate > b.WithdrawDate ? -1 : 0));

  const totalValue = wds.reduce((s, w) => s + w.LineTotal, 0);

  const catMap = {}, deptMap = {}, itemMap = {}, monthMap = {};
  wds.forEach(w => {
    (catMap[w.Category] = catMap[w.Category] || { count: 0, value: 0 }).count++;
    catMap[w.Category].value += w.LineTotal;
    (deptMap[w.DepartmentName] = deptMap[w.DepartmentName] || { count: 0, value: 0 }).count++;
    deptMap[w.DepartmentName].value += w.LineTotal;
    const ik = shareKey_(w.ItemName);
    (itemMap[ik] = itemMap[ik] || { ItemName: w.ItemName, Unit: w.Unit, count: 0, quantity: 0, value: 0 }).count++;
    itemMap[ik].quantity += w.Quantity;
    itemMap[ik].value += w.LineTotal;
    const ym = w.WithdrawDate.slice(0, 7);
    (monthMap[ym] = monthMap[ym] || { count: 0, value: 0 }).count++;
    monthMap[ym].value += w.LineTotal;
  });

  const byCategory = Object.keys(catMap).map(k => ({
    Category: k, count: catMap[k].count, value: round2_(catMap[k].value),
    percent: totalValue ? Math.round((catMap[k].value / totalValue) * 100) : 0
  })).sort((a, b) => b.value - a.value);

  const byDepartment = Object.keys(deptMap).map(k => ({
    DepartmentName: k, count: deptMap[k].count, value: round2_(deptMap[k].value)
  })).sort((a, b) => b.value - a.value).map((r, i) => Object.assign(r, { rank: i + 1 }));

  const topItems = Object.keys(itemMap).map(k => ({
    ItemName: itemMap[k].ItemName, Unit: itemMap[k].Unit, count: itemMap[k].count,
    quantity: round2_(itemMap[k].quantity), value: round2_(itemMap[k].value)
  })).sort((a, b) => b.value - a.value).slice(0, 8);

  // รายเดือนต่อเนื่องตั้งแต่เดือนแรกถึงเดือนล่าสุด (เดือนที่ไม่มีการเบิกแสดง 0)
  const monthly = [];
  const yms = Object.keys(monthMap).sort();
  if (yms.length) {
    let y = Number(yms[0].slice(0, 4)), m = Number(yms[0].slice(5, 7));
    const endY = Number(yms[yms.length - 1].slice(0, 4)), endM = Number(yms[yms.length - 1].slice(5, 7));
    while (y < endY || (y === endY && m <= endM)) {
      const key = y + '-' + ('0' + m).slice(-2);
      const v = monthMap[key] || { count: 0, value: 0 };
      monthly.push({ ym: key, year: y, month: m, count: v.count, value: round2_(v.value) });
      m++; if (m > 12) { m = 1; y++; }
    }
  }

  // ---- รายการคงเหลือ ----
  const stock = shareRows_(SHEET_SHARE_ITEMS).filter(i => shareActive_(i.IsActive) && i.ItemName).map(i => {
    const hasQty = shareHasNum_(i.Quantity);
    const hasVal = shareHasNum_(i.UnitValue) && Number(i.UnitValue) > 0;
    const qty = hasQty ? Number(i.Quantity) : null;
    const val = hasVal ? Number(i.UnitValue) : null;
    return {
      ItemID: i.ItemID, ItemName: i.ItemName, Category: i.Category || 'อื่นๆ', Unit: i.Unit || '',
      Quantity: qty, UnitValue: val,
      LineValue: (qty !== null && val !== null && qty > 0) ? round2_(qty * val) : null,
      Status: shareStockStatus_(i.Quantity, i.UnitValue),
      Note: i.Note || '',
      LastCheckedAt: i.LastCheckedAt ? toIsoDate_(i.LastCheckedAt) : null,
      LastCheckedBy: i.LastCheckedBy || '',
      UpdatedAt: i.UpdatedAt ? toIsoDate_(i.UpdatedAt) : null
    };
  }).sort((a, b) => String(a.Category).localeCompare(String(b.Category), 'th') || String(a.ItemName).localeCompare(String(b.ItemName), 'th'));

  const available = stock.filter(i => i.Quantity !== null && i.Quantity > 0);
  const stockValue = available.reduce((s, i) => s + (i.LineValue || 0), 0);

  // ---- รายการที่ส่งเข้าโครงการ / ผลสำรวจ ----
  const subsAll = shareRows_(SHEET_SHARE_SUBS).filter(s => s.SubmissionID);
  const orderWeight = { [SHARE_STATUS_PENDING]: 0 };
  const submissions = subsAll.map(s => {
    const row = {
      SubmissionID: s.SubmissionID, SubmitDate: toIsoDate_(s.SubmitDate), DepartmentName: s.DepartmentName || '',
      UsageStatus: s.UsageStatus || '', ItemName: s.ItemName || '', Category: s.Category || 'อื่นๆ',
      Quantity: shareHasNum_(s.Quantity) && Number(s.Quantity) > 0 ? Number(s.Quantity) : null,
      Unit: s.Unit || '', Note: s.Note || '', Status: s.Status || SHARE_STATUS_PENDING
    };
    if (isAdmin) { row.SubmitterName = s.SubmitterName || ''; row.Position = s.Position || ''; }
    return row;
  }).sort((a, b) => {
    const wa = orderWeight[a.Status] === 0 ? 0 : 1, wb = orderWeight[b.Status] === 0 ? 0 : 1;
    if (wa !== wb) return wa - wb;
    return a.SubmitDate < b.SubmitDate ? 1 : (a.SubmitDate > b.SubmitDate ? -1 : 0);
  }).slice(0, 500);

  const respondedDepts = {};
  subsAll.forEach(s => { if (s.DepartmentName) respondedDepts[shareKey_(s.DepartmentName)] = 1; });
  const reportedItems = subsAll.filter(s => (s.Status || SHARE_STATUS_PENDING) !== SHARE_STATUS_NONE);
  const subCounts = {};
  subsAll.forEach(s => { const st = s.Status || SHARE_STATUS_PENDING; subCounts[st] = (subCounts[st] || 0) + 1; });

  const impact = getShareImpact_(shareRows_(SHEET_SHARE_WD));

  return {
    isAdmin: isAdmin,
    categories: SHARE_CATEGORIES,
    usageOptions: SHARE_USAGE_OPTIONS,
    monthNames: PL_MONTH_NAMES,
    departments: shareDepartmentNames_(),
    kpi: {
      withdrawalCount: wds.length,
      withdrawalValue: round2_(totalValue),
      withdrawalDeptCount: Object.keys(deptMap).length,
      unpricedWithdrawals: wds.filter(w => w.UnitPrice === null).length,
      firstDate: wds.length ? wds[wds.length - 1].WithdrawDate : null,
      lastDate: wds.length ? wds[0].WithdrawDate : null,
      stockLines: available.length,
      stockValue: round2_(stockValue),
      stockPricePending: available.filter(i => i.LineValue === null).length,
      stockCountPending: stock.filter(i => i.Quantity === null).length,
      submissionsPending: subCounts[SHARE_STATUS_PENDING] || 0,
      submissionsAccepted: subCounts[SHARE_STATUS_ACCEPTED] || 0,
      reportedItemCount: reportedItems.length,
      respondedDeptCount: Object.keys(respondedDepts).length
    },
    monthly: monthly,
    byCategory: byCategory,
    byDepartment: byDepartment,
    topItems: topItems,
    stock: stock,
    submissions: submissions,
    submissionCounts: subCounts,
    withdrawals: wds,
    impact: impact
  };
}

function getShareSubmitFormLink() {
  return (ScriptApp.getService().getUrl() || '') + '?page=share&action=form';
}
function getShareStockLink() {
  return (ScriptApp.getService().getUrl() || '') + '?page=share&action=stock';
}

// ----------------------------- WRITE: ส่งเข้าโครงการ (สาธารณะ ไม่ต้องใช้ PIN) -----------------------------
/**
 * หน่วยงานแจ้งรายการที่เกินจำเป็น/หมดความจำเป็น เพื่อส่งเข้าโครงการ — ใครก็บันทึกได้ผ่านลิงก์
 * data: {DepartmentName, SubmitterName, Position, NoItems, ItemName, Category, Quantity, Unit, UsageStatus, Note}
 * รายการจะอยู่สถานะ "รอตรวจสอบ" จนกว่าผู้ดูแลระบบจะรับเข้าคลัง/ไม่รับ — ไม่กระทบยอดคงเหลือทันที
 */
function submitShareItem(data) {
  data = data || {};
  let dept = shareText_(data.DepartmentName, 120);
  if (!dept && !isDeptMode_('share')) dept = NO_DEPT_TEXT;
  if (!dept) throw new Error('กรุณาระบุหน่วยงาน');
  const submitter = shareText_(data.SubmitterName, 120);
  if (!submitter) throw new Error('กรุณากรอกชื่อ-สกุลผู้แจ้ง');
  const position = shareText_(data.Position, 120);
  const noItems = data.NoItems === true || data.NoItems === 'true';
  const now = new Date();
  const id = newId_('SHS');

  if (noItems) {
    shareAppend_(SHEET_SHARE_SUBS, {
      SubmissionID: id, SubmitDate: now, DepartmentName: dept, SubmitterName: submitter, Position: position,
      UsageStatus: '-', ItemName: SHARE_NO_ITEM_TEXT, Category: '', Quantity: 0, Unit: '', Note: shareText_(data.Note, 300),
      Status: SHARE_STATUS_NONE, CreatedAt: now
    });
    syncShareLinkedProject_();
    return { ok: true, submissionId: id, noItems: true };
  }

  const itemName = shareText_(data.ItemName, 200);
  if (!itemName) throw new Error('กรุณาระบุชื่อรายการ');
  const qty = Number(data.Quantity);
  if (!qty || qty <= 0) throw new Error('กรุณาระบุจำนวนให้ถูกต้อง');
  const unit = shareText_(data.Unit, 30);
  if (!unit) throw new Error('กรุณาระบุหน่วยนับ เช่น กล่อง ตลับ รีม');
  const usage = SHARE_USAGE_OPTIONS.indexOf(data.UsageStatus) !== -1 ? data.UsageStatus : SHARE_USAGE_OPTIONS[0];

  shareAppend_(SHEET_SHARE_SUBS, {
    SubmissionID: id, SubmitDate: parseShareDate_(data.SubmitDate), DepartmentName: dept, SubmitterName: submitter,
    Position: position, UsageStatus: usage, ItemName: itemName, Category: normalizeShareCategory_(data.Category, itemName),
    Quantity: qty, Unit: unit, Note: shareText_(data.Note, 300), Status: SHARE_STATUS_PENDING, CreatedAt: now
  });
  syncShareLinkedProject_();
  return { ok: true, submissionId: id };
}

// ----------------------------- WRITE: ตรวจสอบรายการที่ส่งเข้า (ผู้ดูแลระบบ) -----------------------------
/**
 * ตรวจสอบรายการที่ส่งเข้าโครงการ — decision: 'accept' (รับเข้าคลัง) | 'reject' (ไม่รับ) | 'reopen' (กลับเป็นรอตรวจสอบ)
 * accept: data = {ItemName, Category, Quantity, Unit, UnitValue} (ผู้ตรวจแก้ไขชื่อ/จำนวน/ราคาได้ก่อนรับเข้า)
 */
function reviewShareSubmission(departmentId, pin, submissionId, decision, data) {
  const auth = assertAuthorized_(departmentId, pin);
  data = data || {};
  return withShareLock_(() => {
    const found = findRowById_(SHEET_SHARE_SUBS, 'SubmissionID', submissionId);
    if (!found) throw new Error('ไม่พบรายการที่ส่งเข้านี้');
    const cur = found.values[found.headers.indexOf('Status')] || SHARE_STATUS_PENDING;
    const now = new Date();

    if (decision === 'reject') {
      if (cur === SHARE_STATUS_ACCEPTED) throw new Error('รายการนี้รับเข้าคลังไปแล้ว ไม่สามารถปฏิเสธได้');
      shareUpdate_(SHEET_SHARE_SUBS, found, { Status: SHARE_STATUS_REJECTED, ReviewedBy: auth.departmentName, ReviewedAt: now });
    } else if (decision === 'reopen') {
      if (cur === SHARE_STATUS_ACCEPTED) throw new Error('รายการนี้รับเข้าคลังไปแล้ว ไม่สามารถย้อนสถานะได้');
      shareUpdate_(SHEET_SHARE_SUBS, found, { Status: SHARE_STATUS_PENDING, ReviewedBy: '', ReviewedAt: '' });
    } else if (decision === 'accept') {
      if (cur === SHARE_STATUS_ACCEPTED) throw new Error('รายการนี้รับเข้าคลังไปแล้ว');
      if (cur === SHARE_STATUS_NONE) throw new Error('รายการนี้เป็นการแจ้งว่าไม่มีรายการ ไม่สามารถรับเข้าคลังได้');
      const get = h => found.values[found.headers.indexOf(h)];
      const itemName = shareText_(data.ItemName !== undefined ? data.ItemName : get('ItemName'), 200);
      if (!itemName) throw new Error('กรุณาระบุชื่อรายการ');
      const qty = Number(data.Quantity !== undefined ? data.Quantity : get('Quantity'));
      if (!qty || qty <= 0) throw new Error('กรุณาระบุจำนวนที่รับเข้าให้ถูกต้อง');
      const unit = shareText_(data.Unit !== undefined ? data.Unit : get('Unit'), 30);
      if (!unit) throw new Error('กรุณาระบุหน่วยนับ');
      const cat = normalizeShareCategory_(data.Category !== undefined ? data.Category : get('Category'), itemName);
      const unitValue = shareHasNum_(data.UnitValue) && Number(data.UnitValue) > 0 ? Number(data.UnitValue) : null;
      const itemId = addToShareStock_(itemName, cat, unit, qty, unitValue, auth.departmentName,
        'รับเข้าจากหน่วยงาน ' + get('DepartmentName'), 'รับเข้าคลัง');
      shareUpdate_(SHEET_SHARE_SUBS, found, {
        Status: SHARE_STATUS_ACCEPTED, LinkedItemID: itemId, ReviewedBy: auth.departmentName, ReviewedAt: now
      });
    } else {
      throw new Error('คำสั่งไม่ถูกต้อง');
    }
    syncShareLinkedProject_();
    return { ok: true };
  });
}

/** เพิ่มจำนวนเข้ารายการคงเหลือ (ถ้ามีชื่อ+หน่วยเดียวกันอยู่แล้วจะรวมยอดให้) คืนค่า ItemID — เรียกภายใน lock เท่านั้น */
function addToShareStock_(name, category, unit, qty, unitValue, by, note, action) {
  const items = sheetToObjects_(SHEET_SHARE_ITEMS);
  const match = items.find(i => shareActive_(i.IsActive) &&
    shareKey_(i.ItemName) === shareKey_(name) && shareKey_(i.Unit) === shareKey_(unit));
  const now = new Date();
  if (match) {
    const before = shareNum_(match.Quantity);
    const found = findRowById_(SHEET_SHARE_ITEMS, 'ItemID', match.ItemID);
    const fields = { Quantity: before + qty, UpdatedAt: now };
    if (unitValue !== null && !(shareHasNum_(match.UnitValue) && Number(match.UnitValue) > 0)) fields.UnitValue = unitValue;
    shareUpdate_(SHEET_SHARE_ITEMS, found, fields);
    logShareStock_(match.ItemID, match.ItemName, action || 'รับเข้าคลัง', before, before + qty, by, note);
    return match.ItemID;
  }
  const itemId = newId_('SHI');
  shareAppend_(SHEET_SHARE_ITEMS, {
    ItemID: itemId, ItemName: name, Category: category, Unit: unit, Quantity: qty,
    UnitValue: unitValue === null ? '' : unitValue, Note: '', IsActive: true, CreatedAt: now, UpdatedAt: now
  });
  logShareStock_(itemId, name, action || 'รับเข้าคลัง', 0, qty, by, note);
  return itemId;
}

// ----------------------------- WRITE: จัดการรายการคงเหลือ (ผู้ดูแลระบบ) -----------------------------
/** เพิ่มรายการคงเหลือใหม่ หรือแก้ไขรายการเดิม (ถ้าส่ง ItemID) — data: {ItemID?, ItemName, Category, Unit, Quantity, UnitValue, Note} */
function saveShareItem(departmentId, pin, data) {
  const auth = assertAuthorized_(departmentId, pin);
  data = data || {};
  return withShareLock_(() => {
    const name = shareText_(data.ItemName, 200);
    if (!name) throw new Error('กรุณาระบุชื่อรายการ');
    const unit = shareText_(data.Unit, 30);
    if (!unit) throw new Error('กรุณาระบุหน่วยนับ');
    if (!shareHasNum_(data.Quantity) || Number(data.Quantity) < 0) throw new Error('กรุณาระบุจำนวนให้ถูกต้อง (0 ขึ้นไป)');
    const qty = Number(data.Quantity);
    if (shareHasNum_(data.UnitValue) && Number(data.UnitValue) < 0) throw new Error('มูลค่าต่อหน่วยต้องไม่ติดลบ');
    const unitValue = shareHasNum_(data.UnitValue) && Number(data.UnitValue) > 0 ? Number(data.UnitValue) : '';
    const cat = normalizeShareCategory_(data.Category, name);
    const note = shareText_(data.Note, 300);

    if (data.ItemID) {
      const found = findRowById_(SHEET_SHARE_ITEMS, 'ItemID', data.ItemID);
      if (!found) throw new Error('ไม่พบรายการคงเหลือนี้');
      const before = found.values[found.headers.indexOf('Quantity')];
      shareUpdate_(SHEET_SHARE_ITEMS, found, {
        ItemName: name, Category: cat, Unit: unit, Quantity: qty, UnitValue: unitValue, Note: note, UpdatedAt: new Date()
      });
      logShareStock_(data.ItemID, name, 'แก้ไขรายการ', shareHasNum_(before) ? Number(before) : '', qty, auth.departmentName, '');
      syncShareLinkedProject_();
      return { ok: true, itemId: data.ItemID };
    }
    const itemId = addToShareStock_(name, cat, unit, qty, unitValue === '' ? null : unitValue, auth.departmentName,
      note || 'เพิ่มรายการโดยผู้ดูแลระบบ', 'เพิ่มรายการ');
    syncShareLinkedProject_();
    return { ok: true, itemId: itemId };
  });
}

/** ตรวจนับ/ยืนยันยอดคงเหลือ: บันทึกวันที่ตรวจสอบล่าสุด และปรับยอดให้ตรงกับจำนวนที่นับได้จริง (ถ้าระบุ) */
function checkShareItem(departmentId, pin, itemId, actualQuantity) {
  const auth = assertAuthorized_(departmentId, pin);
  return withShareLock_(() => {
    const found = findRowById_(SHEET_SHARE_ITEMS, 'ItemID', itemId);
    if (!found) throw new Error('ไม่พบรายการคงเหลือนี้');
    const now = new Date();
    const before = found.values[found.headers.indexOf('Quantity')];
    const fields = { LastCheckedAt: now, LastCheckedBy: auth.departmentName };
    let after = shareHasNum_(before) ? Number(before) : '';
    if (actualQuantity !== null && actualQuantity !== undefined && actualQuantity !== '') {
      const q = Number(actualQuantity);
      if (isNaN(q) || q < 0) throw new Error('จำนวนที่นับได้ต้องเป็นตัวเลข 0 ขึ้นไป');
      fields.Quantity = q; fields.UpdatedAt = now; after = q;
    }
    shareUpdate_(SHEET_SHARE_ITEMS, found, fields);
    logShareStock_(itemId, found.values[found.headers.indexOf('ItemName')], 'ตรวจนับ',
      shareHasNum_(before) ? Number(before) : '', after, auth.departmentName,
      before !== after ? 'ปรับยอดตามการตรวจนับ' : 'ยอดตรงกับระบบ');
    syncShareLinkedProject_();
    return { ok: true };
  });
}

/** ลบรายการคงเหลือ (ซ่อนจากระบบ ไม่ลบแถวจริง เพื่อคงประวัติ) */
function deleteShareItem(departmentId, pin, itemId) {
  const auth = assertAuthorized_(departmentId, pin);
  return withShareLock_(() => {
    const found = findRowById_(SHEET_SHARE_ITEMS, 'ItemID', itemId);
    if (!found) throw new Error('ไม่พบรายการคงเหลือนี้');
    shareUpdate_(SHEET_SHARE_ITEMS, found, { IsActive: false, UpdatedAt: new Date() });
    logShareStock_(itemId, found.values[found.headers.indexOf('ItemName')], 'ลบรายการ',
      shareNum_(found.values[found.headers.indexOf('Quantity')]), '', auth.departmentName, '');
    syncShareLinkedProject_();
    return { ok: true };
  });
}

// ----------------------------- WRITE: เบิก/รับแบ่งปัน (ผู้ดูแลระบบ) -----------------------------
/**
 * บันทึกการเบิก/จ่ายแบ่งปัน — data: {WithdrawDate, ItemID?, ItemName, Category, Quantity, Unit, UnitPrice, DepartmentName, Note}
 * ถ้าเลือก ItemID จากรายการคงเหลือ ระบบจะตัดยอดคงเหลืออัตโนมัติและไม่ให้เบิกเกินคงเหลือ
 */
function recordShareWithdrawal(departmentId, pin, data) {
  const auth = assertAuthorized_(departmentId, pin);
  data = data || {};
  return withShareLock_(() => {
    let receiver = shareText_(data.DepartmentName, 120);
    if (!receiver && !isDeptMode_('share')) receiver = NO_DEPT_TEXT;
    if (!receiver) throw new Error('กรุณาระบุหน่วยงานผู้เบิก');
    const qty = Number(data.Quantity);
    if (!qty || qty <= 0) throw new Error('กรุณาระบุจำนวนให้ถูกต้อง');
    const date = parseShareDate_(data.WithdrawDate);

    let name, unit, cat, price = null, itemId = '';
    if (data.ItemID) {
      const found = findRowById_(SHEET_SHARE_ITEMS, 'ItemID', data.ItemID);
      if (!found || !shareActive_(found.values[found.headers.indexOf('IsActive')])) throw new Error('ไม่พบรายการคงเหลือนี้');
      const g = h => found.values[found.headers.indexOf(h)];
      const stockQty = shareNum_(g('Quantity'));
      if (qty > stockQty) throw new Error('คงเหลือไม่พอ: ' + g('ItemName') + ' เหลือ ' + stockQty + ' ' + g('Unit'));
      name = g('ItemName'); unit = g('Unit'); cat = g('Category'); itemId = data.ItemID;
      if (shareHasNum_(data.UnitPrice)) price = Number(data.UnitPrice);
      else if (shareHasNum_(g('UnitValue')) && Number(g('UnitValue')) > 0) price = Number(g('UnitValue'));
      shareUpdate_(SHEET_SHARE_ITEMS, found, { Quantity: stockQty - qty, UpdatedAt: new Date() });
      logShareStock_(itemId, name, 'เบิก/แบ่งปัน', stockQty, stockQty - qty, auth.departmentName, 'ให้ ' + receiver);
    } else {
      name = shareText_(data.ItemName, 200);
      if (!name) throw new Error('กรุณาระบุชื่อรายการ');
      unit = shareText_(data.Unit, 30);
      if (!unit) throw new Error('กรุณาระบุหน่วยนับ');
      cat = normalizeShareCategory_(data.Category, name);
      if (shareHasNum_(data.UnitPrice)) price = Number(data.UnitPrice);
    }
    if (price !== null && price < 0) throw new Error('ราคาต่อหน่วยต้องไม่ติดลบ');

    const id = newId_('SHW');
    shareAppend_(SHEET_SHARE_WD, {
      WithdrawalID: id, WithdrawDate: date, Category: cat, ItemName: name, Quantity: qty, Unit: unit,
      UnitPrice: price === null ? '' : price, DepartmentName: receiver, ItemID: itemId,
      Note: shareText_(data.Note, 300), RecordedByName: auth.departmentName, CreatedAt: new Date()
    });
    syncShareLinkedProject_();
    return { ok: true, withdrawalId: id };
  });
}

/** ลบรายการเบิก — ถ้าผูกกับรายการคงเหลือ จะคืนยอดกลับเข้าคลังให้ */
function deleteShareWithdrawal(departmentId, pin, withdrawalId) {
  const auth = assertAuthorized_(departmentId, pin);
  return withShareLock_(() => {
    const found = findRowById_(SHEET_SHARE_WD, 'WithdrawalID', withdrawalId);
    if (!found) throw new Error('ไม่พบรายการเบิกนี้');
    const g = h => found.values[found.headers.indexOf(h)];
    const itemId = g('ItemID');
    if (itemId) {
      const item = findRowById_(SHEET_SHARE_ITEMS, 'ItemID', itemId);
      if (item) {
        const before = shareNum_(item.values[item.headers.indexOf('Quantity')]);
        const after = before + shareNum_(g('Quantity'));
        shareUpdate_(SHEET_SHARE_ITEMS, item, { Quantity: after, UpdatedAt: new Date() });
        logShareStock_(itemId, g('ItemName'), 'ยกเลิกการเบิก', before, after, auth.departmentName, 'คืนยอดจากรายการเบิกที่ถูกลบ');
      }
    }
    getSheet_(SHEET_SHARE_WD).deleteRow(found.rowIndex);
    syncShareLinkedProject_();
    return { ok: true };
  });
}

// ----------------------------- โครงการเชื่อมโยง (แถวตัวแทนในชีต Projects) -----------------------------
function syncShareLinkedProject_() {
  try {
    const ov = getShareOverview();
    const k = ov.kpi;
    updateLinkedProjectFields_(LINKED_PROJECT_SHARE_ID, {
      CurrentStatus: `แบ่งปันแล้ว ${k.withdrawalCount.toLocaleString()} รายการ มูลค่า ${k.withdrawalValue.toLocaleString()} บาท • ` +
        `คงเหลือรอแบ่งปัน ${k.stockLines.toLocaleString()} รายการ • รอตรวจสอบ ${k.submissionsPending.toLocaleString()} รายการ`,
      Progress: yearProgressPct_()
    });
  } catch (e) { /* ยังไม่มีชีต/ข้อมูล ก็ข้ามไปเงียบๆ */ }
}

// ----------------------------- นำเข้าข้อมูลเริ่มต้นจากไฟล์ Excel -----------------------------
/** ข้อมูลจากไฟล์ Excel "รายการเบิก / รายการคงเหลือ (รอแบ่งปัน) / ผลสำรวจ" — ใช้ครั้งเดียวกับ importShareProjectData() */
const SHARE_IMPORT_DATA = {
 "stockAsOf": "19 มิ.ย. 2569",
 "withdrawals": [
  {
   "date": "2025-12-16",
   "cat": "ลวดเย็บ/เสียบกระดาษ",
   "item": "ลวดเสียบกระดาษ",
   "qty": 6,
   "unit": "กล่อง",
   "price": 112,
   "dept": "ฝ่ายวิสัญญีวิทยา"
  },
  {
   "date": "2025-12-16",
   "cat": "อื่นๆ",
   "item": "เทปใสแกน 3 นิ้ว 1/2",
   "qty": 6,
   "unit": "ม้วน",
   "price": 36.5,
   "dept": "ฝ่ายเวชศาสตร์ชันสูตร"
  },
  {
   "date": "2025-12-18",
   "cat": "หมึกพิมพ์/ตลับหมึก",
   "item": "หมึก HP 83A",
   "qty": 4,
   "unit": "ตลับ",
   "price": 2450.3,
   "dept": "ฝ่ายยาและเวชภัณฑ์"
  },
  {
   "date": "2025-12-18",
   "cat": "ลวดเย็บ/เสียบกระดาษ",
   "item": "ลวดเสียบกระดาษ",
   "qty": 3,
   "unit": "กล่อง",
   "price": 112,
   "dept": "ฝ่ายยาและเวชภัณฑ์"
  },
  {
   "date": "2025-12-18",
   "cat": "อื่นๆ",
   "item": "เทปใสแกน 3 นิ้ว 1/2",
   "qty": 4,
   "unit": "ม้วน",
   "price": 36.5,
   "dept": "ฝ่ายยาและเวชภัณฑ์"
  },
  {
   "date": "2025-12-18",
   "cat": "ลวดเย็บ/เสียบกระดาษ",
   "item": "ลวดเสียบกระดาษ",
   "qty": 20,
   "unit": "กล่อง",
   "price": 112,
   "dept": "คลินิกปฐมภูมิ"
  },
  {
   "date": "2025-12-22",
   "cat": "หมึกพิมพ์/ตลับหมึก",
   "item": "หมึก HP 05A",
   "qty": 1,
   "unit": "กล่อง",
   "price": 4650,
   "dept": "ห้องคลอด"
  },
  {
   "date": "2025-12-22",
   "cat": "ลวดเย็บ/เสียบกระดาษ",
   "item": "ลวดเสียบกระดาษ",
   "qty": 20,
   "unit": "กล่อง",
   "price": 112,
   "dept": "คลินิกศัลยกรรม"
  },
  {
   "date": "2025-12-22",
   "cat": "กระดาษ",
   "item": "กระดาษ F4",
   "qty": 3,
   "unit": "รีม",
   "price": 139,
   "dept": "คลินิกจิตเวช"
  },
  {
   "date": "2025-12-22",
   "cat": "กระดาษ",
   "item": "กระดาษ F4",
   "qty": 2,
   "unit": "รีม",
   "price": 139,
   "dept": "คลินิกตรวจสุขภาพ"
  },
  {
   "date": "2026-01-06",
   "cat": "หมึกพิมพ์/ตลับหมึก",
   "item": "หมึก HP 78A",
   "qty": 2,
   "unit": "ตลับ",
   "price": 2803.8,
   "dept": "18A"
  },
  {
   "date": "2026-01-09",
   "cat": "ลวดเย็บ/เสียบกระดาษ",
   "item": "ลวดเย็บกระดาษ No.10-1M",
   "qty": 24,
   "unit": "กล่อง",
   "price": 10.62,
   "dept": "คลินิกศัลยกรรม"
  },
  {
   "date": "2026-01-09",
   "cat": "อื่นๆ",
   "item": "เทปใสแกน 3 นิ้ว 1/2",
   "qty": 2,
   "unit": "ม้วน",
   "price": 36.5,
   "dept": "คลินิกศัลยกรรม"
  },
  {
   "date": "2026-01-20",
   "cat": "อื่นๆ",
   "item": "ถังขยะมีฝาปิดขนาด 10 ลิตร",
   "qty": 1,
   "unit": "ถัง",
   "price": 320,
   "dept": "อาคารพักเจ้าหน้าที่"
  },
  {
   "date": "2026-01-23",
   "cat": "กระดาษ",
   "item": "กระดาษ F4",
   "qty": 4,
   "unit": "รีม",
   "price": 139,
   "dept": "คลินิกเวชปฏิบัติ"
  },
  {
   "date": "2026-01-23",
   "cat": "ลวดเย็บ/เสียบกระดาษ",
   "item": "ลวดเย็บกระดาษ No.10-1M",
   "qty": 24,
   "unit": "กล่อง",
   "price": 10.62,
   "dept": "ฝ่ายเวชสารสนเทศ"
  },
  {
   "date": "2026-01-23",
   "cat": "ปากกา",
   "item": "ปากกาเคมี 2 หัว สีดำ",
   "qty": 5,
   "unit": "ด้าม",
   "price": 45,
   "dept": "ฝ่ายเวชสารสนเทศ"
  },
  {
   "date": "2026-01-23",
   "cat": "ปากกา",
   "item": "ปากกาเคมี 2 หัว สีแดง",
   "qty": 5,
   "unit": "ด้าม",
   "price": 45,
   "dept": "ฝ่ายเวชสารสนเทศ"
  },
  {
   "date": "2026-01-28",
   "cat": "หมึกพิมพ์/ตลับหมึก",
   "item": "หมึก HP 78A",
   "qty": 2,
   "unit": "ตลับ",
   "price": 2803.8,
   "dept": "หน่วยผลิตยา"
  },
  {
   "date": "2026-01-28",
   "cat": "ปากกา",
   "item": "ปากกาเคมี 2 หัว สีน้ำเงิน",
   "qty": 5,
   "unit": "ด้าม",
   "price": 45,
   "dept": "หน่วยผลิตยา"
  },
  {
   "date": "2026-01-28",
   "cat": "หมึกพิมพ์/ตลับหมึก",
   "item": "หมึก HP 36A",
   "qty": 1,
   "unit": "ตลับ",
   "price": 3950,
   "dept": "หน่วยผลิตยา"
  },
  {
   "date": "2026-01-28",
   "cat": "ลวดเย็บ/เสียบกระดาษ",
   "item": "ลวดเสียบกระดาษ",
   "qty": 4,
   "unit": "กล่อง",
   "price": 112,
   "dept": "หน่วยผลิตยา"
  },
  {
   "date": "2026-01-30",
   "cat": "หมึกพิมพ์/ตลับหมึก",
   "item": "ตลับหมึก Canon 045 สีดำ",
   "qty": 1,
   "unit": "กล่อง",
   "price": 800,
   "dept": "จำหน่าย"
  },
  {
   "date": "2026-01-30",
   "cat": "หมึกพิมพ์/ตลับหมึก",
   "item": "ตลับหมึก Canon 045 สีน้ำเงิน",
   "qty": 2,
   "unit": "กล่อง",
   "price": 1000,
   "dept": "จำหน่าย"
  },
  {
   "date": "2026-01-30",
   "cat": "หมึกพิมพ์/ตลับหมึก",
   "item": "ตลับหมึก Canon 045 สีแดง",
   "qty": 3,
   "unit": "กล่อง",
   "price": 1000,
   "dept": "จำหน่าย"
  },
  {
   "date": "2026-01-30",
   "cat": "หมึกพิมพ์/ตลับหมึก",
   "item": "ตลับหมึก Canon 045 สีเหลือง",
   "qty": 3,
   "unit": "กล่อง",
   "price": 1000,
   "dept": "จำหน่าย"
  },
  {
   "date": "2026-01-30",
   "cat": "หมึกพิมพ์/ตลับหมึก",
   "item": "หมึก HP 202A",
   "qty": 5,
   "unit": "ตลับ",
   "price": 240,
   "dept": "จำหน่าย"
  },
  {
   "date": "2026-03-11",
   "cat": "ลวดเย็บ/เสียบกระดาษ",
   "item": "ลวดเสียบกระดาษ",
   "qty": 12,
   "unit": "กล่อง",
   "price": 112,
   "dept": "ห้องศูนย์หัวใจ"
  },
  {
   "date": "2026-03-11",
   "cat": "อื่นๆ",
   "item": "กาวแท่ง",
   "qty": 5,
   "unit": "แท่ง",
   "price": 38,
   "dept": "ห้องศูนย์หัวใจ"
  },
  {
   "date": "2026-03-11",
   "cat": "อื่นๆ",
   "item": "ใบสั่งซื้อหัวกาชาด",
   "qty": 2,
   "unit": "รีม",
   "price": null,
   "dept": "ฝ่ายบริหารงานพัสดุฯ"
  },
  {
   "date": "2026-03-19",
   "cat": "ลวดเย็บ/เสียบกระดาษ",
   "item": "ลวดเย็บกระดาษ No.10-1M",
   "qty": 24,
   "unit": "กล่อง",
   "price": 10.62,
   "dept": "ห้องตรวจสุขภาพ"
  },
  {
   "date": "2026-03-19",
   "cat": "หมึกพิมพ์/ตลับหมึก",
   "item": "หมึก HP 83A",
   "qty": 2,
   "unit": "ตลับ",
   "price": 2450.3,
   "dept": "ห้องตรวจสุขภาพ"
  },
  {
   "date": "2026-03-19",
   "cat": "ลวดเย็บ/เสียบกระดาษ",
   "item": "ลวดเย็บกระดาษ No.10-1M",
   "qty": 24,
   "unit": "กล่อง",
   "price": 10.62,
   "dept": "ฝ่ายบริหารอาคารฯ"
  },
  {
   "date": "2026-04-02",
   "cat": "หมึกพิมพ์/ตลับหมึก",
   "item": "หมึก HP 83A",
   "qty": 1,
   "unit": "ตลับ",
   "price": 2450.3,
   "dept": "ฝ่ายบริหารงานทั่วไป"
  },
  {
   "date": "2026-04-02",
   "cat": "หมึกพิมพ์/ตลับหมึก",
   "item": "หมึก HP 78A",
   "qty": 1,
   "unit": "ตลับ",
   "price": 2803.8,
   "dept": "ฝ่ายบริหารงานทั่วไป"
  },
  {
   "date": "2026-04-03",
   "cat": "ปากกา",
   "item": "ปากกาเคมี 2 หัว สีน้ำเงิน",
   "qty": 7,
   "unit": "ด้าม",
   "price": 45,
   "dept": "20B"
  },
  {
   "date": "2026-04-03",
   "cat": "ปากกา",
   "item": "ปากกาเคมี 2 หัว สีน้ำเงิน",
   "qty": 5,
   "unit": "ด้าม",
   "price": 45,
   "dept": "ฝ่ายบริหารอาคารฯ"
  },
  {
   "date": "2026-04-23",
   "cat": "กระดาษ",
   "item": "กระดาษ A3",
   "qty": 3,
   "unit": "รีม",
   "price": 255,
   "dept": "OPD"
  },
  {
   "date": "2026-05-16",
   "cat": "กระดาษ",
   "item": "กระดาษ F4",
   "qty": 4,
   "unit": "รีม",
   "price": 139,
   "dept": "ฝ่ายบริหารงานพัสดุฯ"
  },
  {
   "date": "2026-06-18",
   "cat": "หมึกพิมพ์/ตลับหมึก",
   "item": "หมึก HP 05A",
   "qty": 1,
   "unit": "กล่อง",
   "price": 4650,
   "dept": "ห้องคลอด"
  },
  {
   "date": "2026-06-19",
   "cat": "หมึกพิมพ์/ตลับหมึก",
   "item": "หมึก HP 83A",
   "qty": 2,
   "unit": "ตลับ",
   "price": 2450.3,
   "dept": "ฝ่ายยาและเวชภัณฑ์"
  },
  {
   "date": "2026-06-19",
   "cat": "หมึกพิมพ์/ตลับหมึก",
   "item": "หมึก HP 78A",
   "qty": 1,
   "unit": "ตลับ",
   "price": 2803.8,
   "dept": "ฝ่ายยาและเวชภัณฑ์"
  },
  {
   "date": "2026-06-19",
   "cat": "อื่นๆ",
   "item": "เทปใสแกน 3 นิ้ว 1/2",
   "qty": 2,
   "unit": "ม้วน",
   "price": 36.5,
   "dept": "OPD"
  },
  {
   "date": "2026-06-19",
   "cat": "กระดาษ",
   "item": "กระดาษสีชมพู A4 ขนาด 120 แกรม",
   "qty": 1,
   "unit": "แพ็ค",
   "price": 210,
   "dept": "OPD"
  },
  {
   "date": "2026-06-19",
   "cat": "กระดาษ",
   "item": "กระดาษ A3",
   "qty": 5,
   "unit": "รีม",
   "price": 285,
   "dept": "OPD"
  },
  {
   "date": "2026-06-19",
   "cat": "กระดาษ",
   "item": "กระดาษ A3",
   "qty": 8,
   "unit": "รีม",
   "price": 285,
   "dept": "คลินิกจิตเวช"
  }
 ],
 "stock": [
  {
   "item": "หมึก HP 80A",
   "qty": 1,
   "unit": "กล่อง",
   "val": 5170
  },
  {
   "item": "หมึก Brother",
   "qty": 1,
   "unit": "ตลับ",
   "val": null
  },
  {
   "item": "หมึก Canon pixma รวมสี",
   "qty": 1,
   "unit": "ตลับ",
   "val": null
  },
  {
   "item": "หมึก Canon pixma สีดำ",
   "qty": 1,
   "unit": "ตลับ",
   "val": null
  },
  {
   "item": "ลวดเย็บกระดาษ M8-5M",
   "qty": 5,
   "unit": "กล่อง",
   "val": null
  },
  {
   "item": "สติ๊กเกอร์ต่อเนื่อง,สติ๊กเกอร์ไดเรคเทอร์มอล ไม่พิมพ์ข้อความ",
   "qty": 6,
   "unit": "กล่อง",
   "val": null
  },
  {
   "item": "เทปใสแกน 3 นิ้ว 1/2",
   "qty": 10,
   "unit": "ม้วน",
   "val": 36.5
  },
  {
   "item": "ปากกาเคมี 2 หัว สีดำ",
   "qty": 6,
   "unit": "ด้าม",
   "val": 45
  },
  {
   "item": "กระดาษสต๊อคการ์ด",
   "qty": null,
   "unit": null,
   "val": null
  },
  {
   "item": "ถังขยะมีฝาปิดขนาด 10 ลิตร",
   "qty": 4,
   "unit": "ถัง",
   "val": 320
  },
  {
   "item": "สติ๊กเกอร์ต่อเนื่อง",
   "qty": 12,
   "unit": "แพ็ค",
   "val": null
  },
  {
   "item": "COMPUTE หมึกเติม รุ่น HP GT-52 สีเหลือง ขนาด 120 cc.",
   "qty": 2,
   "unit": "ขวด",
   "val": 145
  },
  {
   "item": "COMPUTE หมึกเติม รุ่น HP GT-52 สีแดง ขนาด 120 cc.",
   "qty": 3,
   "unit": "ขวด",
   "val": 145
  },
  {
   "item": "COMPUTE หมึกเติม รุ่น HP GT-52 สีน้ำเงิน ขนาด 120 cc.",
   "qty": 1,
   "unit": "ขวด",
   "val": 145
  }
 ],
 "survey": [
  {
   "date": "2025-11-26",
   "name": "นางสาวยุพาวรรณ หามนตรี",
   "pos": "เจ้าหน้าที่พัสดุ (ฝ่ายผู้ป่วยนอก)",
   "dept": "ศูนย์ประกันสุขภาพ สาขา",
   "status": "หมดความจำเป็นในการใช้งาน",
   "item": "สติ๊กเกอร์ต่อเนื่อง,สติ๊กเกอร์ไดเรคเทอร์มอล ไม่พิมพ์ข้อความ",
   "note": null
  },
  {
   "date": "2025-11-26",
   "name": "นางสาวนิรินดา ทองขำ",
   "pos": "เจ้าหน้าที่สถิติ",
   "dept": "ฝ่ายบริหารการเบิก-จ่าย",
   "status": "หมดความจำเป็นในการใช้งาน",
   "item": "กระดาษ A5",
   "note": null
  },
  {
   "date": "2025-11-26",
   "name": "นางสาวสุชาดา รัตนบัวแก้ว",
   "pos": "เจ้าหน้าที่บริหารงานทั่วไป",
   "dept": "หน่วยซ่อมบำรุง ฝ่ายอาคารฯ",
   "status": "มีรายการเกินจำเป็น (มีปริมาณมากใช้งานไม่ทัน)",
   "item": "ลวดเสียบกระดาษ",
   "note": null
  },
  {
   "date": "2025-11-26",
   "name": "นางสาวเยาวภา ธิติธัญญาภัค",
   "pos": "เจ้าหน้าที่ธุรการ",
   "dept": "ฝ่ายวิชาการและงานวิจัย",
   "status": "หมดความจำเป็นในการใช้งาน",
   "item": "แฟ้มสันกว้าง",
   "note": "(งานกิจกรรมวิชาการและการจัดการความรู้ และ งานเวชนิทัศน์และโสตทัศนศึกษา)"
  },
  {
   "date": "2025-11-27",
   "name": "นางสาวทิพยวรรณ อ้วนชัย",
   "pos": "เจ้าหน้าที่บริหารงานทั่วไป",
   "dept": "ตึกสว่างวัฒนา",
   "status": "-",
   "item": "ไม่มีรายการที่เกินความจำเป็น",
   "note": null
  },
  {
   "date": "2025-12-01",
   "name": "นางสาวจิตสุภา ส่งเสริม",
   "pos": "เจ้าหน้าที่บริหารงานทั่วไป",
   "dept": "ฝ่ายเวชศาสตร์ชันสูตร",
   "status": "หมดความจำเป็นในการใช้งาน",
   "item": "ลวดเย็บกระดาษ No.3-1M",
   "note": null
  },
  {
   "date": "2025-12-01",
   "name": "นางอรรฆรชา มัคสุวรรณ",
   "pos": "เจ้าหน้าที่ธุรการ",
   "dept": "ฝ่ายทันตกรรม",
   "status": "-",
   "item": "ทันตกรรมไม่มีรายการส่งค่ะ ",
   "note": null
  },
  {
   "date": "2025-12-04",
   "name": "นางสาวสุชาดา รัตนบัวแก้ว",
   "pos": "เจ้าหน้าที่บริหารงานทั่วไป",
   "dept": "หน่วยธุรการทั่วไป ฝ่ายอาคารฯ",
   "status": "มีรายการเกินจำเป็น (มีปริมาณมากใช้งานไม่ทัน)",
   "item": "เทปใสทุกขนาด",
   "note": null
  },
  {
   "date": "2025-12-04",
   "name": "นางสาวสุชาดา รัตนบัวแก้ว",
   "pos": "เจ้าหน้าที่บริหารงานทั่วไป",
   "dept": "หน่วยธุรการทั่วไป ฝ่ายอาคารฯ",
   "status": "หมดความจำเป็นในการใช้งาน",
   "item": "ตลับหมึกใหม่ทุกชนิด",
   "note": "83A"
  },
  {
   "date": "2025-12-04",
   "name": "นางสาวสุชาดา รัตนบัวแก้ว",
   "pos": "เจ้าหน้าที่บริหารงานทั่วไป",
   "dept": "หน่วยซ่อมบำรุง ฝ่ายอาคารฯ",
   "status": "หมดความจำเป็นในการใช้งาน",
   "item": "ใบสั่งซื้อหัวกาชาด",
   "note": null
  },
  {
   "date": "2025-12-04",
   "name": "นางสาวสุชาดา รัตนบัวแก้ว",
   "pos": "เจ้าหน้าที่บริหารงานทั่วไป",
   "dept": "หน่วยซ่อมบำรุง ฝ่ายอาคารฯ",
   "status": "หมดความจำเป็นในการใช้งาน",
   "item": "กระดาษ F4",
   "note": null
  }
 ]
};

/**
 * นำเข้าข้อมูลเริ่มต้นของโครงการ (รันจาก Apps Script editor: เลือก importShareProjectData > Run)
 * รันซ้ำได้อย่างปลอดภัย — ตรวจรายการซ้ำก่อนเขียนทุกครั้ง
 */
function importShareProjectData() {
  ensureShareDepartmentsSeeded_();
  const now = new Date();
  const by = 'ระบบ (นำเข้าข้อมูล)';
  let w = 0, s = 0, v = 0, skipped = 0;

  const wdExisting = sheetToObjects_(SHEET_SHARE_WD);
  SHARE_IMPORT_DATA.withdrawals.forEach(r => {
    const dup = wdExisting.some(e => toIsoDate_(e.WithdrawDate) === r.date && shareKey_(e.ItemName) === shareKey_(r.item) &&
      shareKey_(e.DepartmentName) === shareKey_(r.dept) && Number(e.Quantity) === r.qty &&
      (shareHasNum_(e.UnitPrice) ? Number(e.UnitPrice) : null) === r.price);
    if (dup) { skipped++; return; }
    shareAppend_(SHEET_SHARE_WD, {
      WithdrawalID: newId_('SHW'), WithdrawDate: new Date(r.date), Category: r.cat || guessShareCategory_(r.item),
      ItemName: r.item, Quantity: r.qty, Unit: r.unit, UnitPrice: r.price === null ? '' : r.price,
      DepartmentName: r.dept, ItemID: '', Note: 'นำเข้าข้อมูลย้อนหลังจากไฟล์ Excel', RecordedByName: by, CreatedAt: now
    });
    w++;
  });

  const stockExisting = sheetToObjects_(SHEET_SHARE_ITEMS);
  SHARE_IMPORT_DATA.stock.forEach(r => {
    const dup = stockExisting.some(e => shareActive_(e.IsActive) && shareKey_(e.ItemName) === shareKey_(r.item));
    if (dup) { skipped++; return; }
    const id = newId_('SHI');
    shareAppend_(SHEET_SHARE_ITEMS, {
      ItemID: id, ItemName: r.item, Category: guessShareCategory_(r.item), Unit: r.unit || '',
      Quantity: r.qty === null ? '' : r.qty, UnitValue: r.val === null ? '' : r.val,
      Note: 'ข้อมูล ณ ' + SHARE_IMPORT_DATA.stockAsOf, IsActive: true, CreatedAt: now, UpdatedAt: now
    });
    logShareStock_(id, r.item, 'นำเข้า', 0, r.qty === null ? '' : r.qty, by, 'ยอดคงเหลือ ณ ' + SHARE_IMPORT_DATA.stockAsOf);
    s++;
  });

  const subExisting = sheetToObjects_(SHEET_SHARE_SUBS);
  SHARE_IMPORT_DATA.survey.forEach(r => {
    const dup = subExisting.some(e => toIsoDate_(e.SubmitDate) === r.date && shareKey_(e.SubmitterName) === shareKey_(r.name) &&
      shareKey_(e.DepartmentName) === shareKey_(r.dept) && shareKey_(e.ItemName) === shareKey_(r.item));
    if (dup) { skipped++; return; }
    const none = r.status === '-' || /^ไม่มีรายการ/.test(r.item) || /ไม่มีรายการ/.test(r.item);
    shareAppend_(SHEET_SHARE_SUBS, {
      SubmissionID: newId_('SHS'), SubmitDate: new Date(r.date), DepartmentName: r.dept, SubmitterName: r.name, Position: r.pos,
      UsageStatus: r.status, ItemName: r.item, Category: none ? '' : guessShareCategory_(r.item), Quantity: '', Unit: '',
      Note: r.note || '', Status: none ? SHARE_STATUS_NONE : SHARE_STATUS_SURVEY, CreatedAt: now
    });
    v++;
  });

  Logger.log('นำเข้าโครงการคัดแยก แบ่งปัน: รายการเบิก ' + w + ' • รายการคงเหลือ ' + s + ' • ผลสำรวจ ' + v + ' • ข้าม(มีอยู่แล้ว) ' + skipped);
  syncShareLinkedProject_();
}

// ============================= สอบถามเกี่ยวกับโครงการ (ถาม-ตอบ) =============================
// ใครก็ถามได้โดยไม่ต้องมี PIN — ถ้าตอบอัตโนมัติได้จะตอบทันที ไม่ได้จะส่งอีเมลแจ้งหน่วยงานที่รับผิดชอบ
// และรอผู้ดูแลระบบ (ต้องมีสิทธิ์ IsAdmin) เข้ามาตอบในระบบ

/** หาอีเมลของหน่วยงานจากชื่อ (ใช้ DepartmentEmail ในชีต Departments) */
function findDepartmentEmailByName_(name) {
  if (!name) return '';
  const dept = sheetToObjects_(SHEET_DEPARTMENTS).find(d => d.DepartmentName === name);
  return (dept && dept.DepartmentEmail) ? String(dept.DepartmentEmail).trim() : '';
}

/**
 * ลองตอบคำถามอัตโนมัติจากข้อมูลโครงการที่มีอยู่แล้ว (กฎอย่างง่ายจับคำสำคัญ)
 * คืนค่าข้อความคำตอบถ้าตอบได้ หรือ null ถ้าตอบไม่ได้ (ต้องส่งต่อให้ผู้ดูแลระบบ)
 */
function tryAutoAnswerProjectQuestion_(project, questionText) {
  const q = String(questionText || '');

  if (/คืบหน้า|สถานะ|ไปถึงไหน|progress|status/i.test(q)) {
    return `สถานะปัจจุบันของ "${project.ProjectName}": ${project.CurrentStatus || 'ยังไม่มีข้อมูลสถานะ'} ` +
      `(ความคืบหน้า ${project.Progress || 0}%, หมวดสถานะ: ${project.StatusCategory || '-'})`;
  }
  if (/ผู้รับผิดชอบ|ใครดูแล|ติดต่อใคร|ติดต่อหน่วยงาน|contact|responsible/i.test(q)) {
    return `โครงการ "${project.ProjectName}" ดูแลโดย ${project.ResponsiblePerson || 'ไม่ระบุผู้รับผิดชอบ'} ` +
      `(${project.ResponsibleUnit || project.Department || 'ไม่ระบุหน่วยงาน'})`;
  }
  if (/เมื่อไหร่|กำหนด|วันที่|เริ่ม|สิ้นสุด|เสร็จ|deadline|timeline|date/i.test(q)) {
    return `โครงการ "${project.ProjectName}" ดำเนินการตั้งแต่ ${toIsoDate_(project.StartDate) || 'ไม่ระบุ'} ` +
      `ถึง ${toIsoDate_(project.TargetDate) || 'ไม่ระบุกำหนดสิ้นสุด'}`;
  }
  if (/รายละเอียด|เกี่ยวกับ|คืออะไร|ทำอะไร|about|detail/i.test(q) && project.Description) {
    return `รายละเอียดโครงการ "${project.ProjectName}": ${project.Description}`;
  }
  if (/ประเภท|หมวด|category/i.test(q)) {
    return `โครงการ "${project.ProjectName}" อยู่ในประเภท: ${project.Category || 'ไม่ระบุ'}`;
  }
  return null;
}

/**
 * ส่งคำถามเกี่ยวกับโครงการ — ไม่ต้องมี PIN (เปิดให้ทุกคนถามได้)
 * ถ้าตอบอัตโนมัติได้ จะบันทึกคำตอบและคืนค่าทันที ถ้าตอบไม่ได้ จะส่งอีเมลแจ้งหน่วยงานที่รับผิดชอบ (ถ้ามีอีเมลในระบบ)
 * แล้วรอผู้ดูแลระบบเข้ามาตอบ
 */
function submitProjectQuestion(projectId, askerName, askerContact, questionText) {
  const qText = String(questionText || '').trim();
  if (!qText) throw new Error('กรุณากรอกคำถาม');

  const found = findRowById_(SHEET_PROJECTS, 'ProjectID', projectId);
  if (!found) throw new Error('ไม่พบโครงการ');
  const project = {};
  found.headers.forEach((h, i) => project[h] = found.values[i]);

  const qSh = getSheet_('ProjectQuestions');
  const headers = qSh.getRange(1, 1, 1, qSh.getLastColumn()).getValues()[0];
  const now = new Date();
  const questionId = newId_('Q');
  const askerNameVal = String(askerName || '').trim();
  const askerContactVal = String(askerContact || '').trim();

  const autoAnswer = tryAutoAnswerProjectQuestion_(project, qText);
  const status = autoAnswer ? 'answered-auto' : 'pending';

  const rowObj = {
    QuestionID: questionId, ProjectID: projectId, ProjectName: project.ProjectName || '',
    AskerName: askerNameVal, AskerContact: askerContactVal, QuestionText: qText,
    Status: status, AutoAnswer: autoAnswer || '', AdminAnswer: '', AnsweredByName: autoAnswer ? 'ระบบ (ตอบอัตโนมัติ)' : '',
    CreatedAt: now, AnsweredAt: autoAnswer ? now : ''
  };
  const row = headers.map(h => (h in rowObj ? rowObj[h] : ''));
  qSh.appendRow(row);

  if (!autoAnswer) {
    const email = findDepartmentEmailByName_(project.ResponsibleUnit) || findDepartmentEmailByName_(project.Department);
    if (email) {
      try {
        MailApp.sendEmail({
          to: email,
          subject: `[สอบถามโครงการ] ${project.ProjectName || projectId}`,
          body: `มีคำถามใหม่เกี่ยวกับโครงการ "${project.ProjectName || projectId}"\n\n` +
            `ผู้ถาม: ${askerNameVal || 'ไม่ระบุชื่อ'}${askerContactVal ? ' (' + askerContactVal + ')' : ''}\n\n` +
            `คำถาม:\n${qText}\n\n` +
            `กรุณาเข้าสู่ระบบผู้ดูแล แล้วเปิดหน้ารายละเอียดโครงการนี้เพื่อตอบคำถาม`
        });
      } catch (e) {
        Logger.log('ส่งอีเมลแจ้งคำถามไม่สำเร็จ: ' + e);
      }
    }
  }

  return { ok: true, questionId: questionId, autoAnswered: !!autoAnswer, answer: autoAnswer || null };
}

/** รายการคำถาม-คำตอบทั้งหมดของโครงการหนึ่ง (ใหม่สุดก่อน) — แสดงได้กับทุกคน ไม่ต้องมี PIN */
function getProjectQuestions(projectId) {
  return sheetToObjects_('ProjectQuestions')
    .filter(q => q.ProjectID === projectId)
    .map(q => ({
      QuestionID: q.QuestionID, AskerName: q.AskerName, AskerContact: q.AskerContact,
      QuestionText: q.QuestionText, Status: q.Status, AutoAnswer: q.AutoAnswer, AdminAnswer: q.AdminAnswer,
      AnsweredByName: q.AnsweredByName, CreatedAt: toIsoDate_(q.CreatedAt), AnsweredAt: toIsoDate_(q.AnsweredAt)
    }))
    .sort((a, b) => new Date(b.CreatedAt) - new Date(a.CreatedAt));
}

/** ผู้ดูแลระบบ (ต้องมีสิทธิ์ IsAdmin) ตอบคำถามที่ระบบตอบอัตโนมัติไม่ได้ */
function answerProjectQuestion(departmentId, pin, questionId, answerText) {
  const auth = assertAuthorized_(departmentId, pin);
  const answer = String(answerText || '').trim();
  if (!answer) throw new Error('กรุณากรอกคำตอบ');

  const found = findRowById_('ProjectQuestions', 'QuestionID', questionId);
  if (!found) throw new Error('ไม่พบคำถามนี้');
  const headers = found.headers;
  const sh = getSheet_('ProjectQuestions');
  const now = new Date();
  const fields = { AdminAnswer: answer, Status: 'answered-admin', AnsweredByName: auth.departmentName, AnsweredAt: now };
  Object.keys(fields).forEach(key => {
    const col = headers.indexOf(key);
    if (col !== -1) sh.getRange(found.rowIndex, col + 1).setValue(fields[key]);
  });
  return { ok: true };
}

// ============================================================================================
// ทะเบียนผู้รับซื้อ + การส่งจำหน่ายรายรอบ (ทอดไม่ทิ้ง / จากครัวสู่ท้องฟ้า / ตลับหมึกเก่าเราขอ)
// และโครงการ "ปฏิทินเก่าเราขอ" (บริจาคปฏิทินเก่าทำอักษรเบรลล์)
// ============================================================================================
const SHEET_BUYERS   = 'Buyers';
const SHEET_FTF_SHIP = 'FryToFlyShipments';
const SHEET_KIT_SHIP = 'KitchenOilShipments';
const SHEET_CAL_DON  = 'CalendarDonations';

const BUYER_HEADERS = ['BuyerID', 'BuyerName', 'BuyerKey', 'Phone', 'Projects', 'Note', 'IsActive', 'CreatedAt'];
const SHIP_HEADERS  = ['ShipmentID', 'ShipmentDate', 'BuyerName', 'Quantity', 'Unit', 'Liters', 'UnitPrice', 'Amount', 'Note', 'RecordedByName', 'CreatedAt', 'UpdatedAt'];
const CAL_HEADERS   = ['DonationID', 'DonationDate', 'RoundName', 'DonorType', 'DonorName', 'WeightKg', 'Note', 'RecordedByName', 'CreatedAt', 'UpdatedAt'];

const KIT_UNITS = ['ลิตร', 'ปี๊บ'];
const CAL_DONOR_TYPES = ['หน่วยงาน', 'เจ้าหน้าที่', 'บริษัท/องค์กรภายนอก'];

// ค่าเริ่มต้นของตัวแปรคำนวณ (ปรับได้ที่ชีต EnvironmentalSettings โดยไม่ต้องแก้โค้ด)
const ENV_KITCHEN_LITERS_PER_CAN = 18;     // น้ำมันที่บรรจุในปี๊บ 1 ปี๊บ ≈ 18 ลิตร
const ENV_KITCHEN_CO2_PER_LITER_KG = 2.5;  // ใช้ค่าเดียวกับโครงการทอดไม่ทิ้ง (น้ำมันใช้แล้วนำไปแปรรูป แทนการทิ้ง)
const ENV_CALENDAR_TREES_PER_TON = 17;     // รีไซเคิลกระดาษ 1 ตัน ช่วยรักษาต้นไม้ได้ประมาณ 17 ต้น (ค่าประมาณทั่วไป)
const ENV_CALENDAR_CO2_PER_KG = 0.9;       // รีไซเคิลกระดาษ 1 กก. ลด CO2 ได้ประมาณ (กก.)

// ---------- helper ทั่วไป ----------
function numOr0_(v) { const n = Number(v); return isNaN(n) ? 0 : n; }

/** คืนชีตตามชื่อ ถ้ายังไม่มีให้สร้างพร้อมหัวตาราง (ทำให้ฟังก์ชันใหม่ใช้งานได้ทันทีแม้ยังไม่ได้รัน setupSheets ซ้ำ) */
function ensureSheet_(name, headers) {
  const ss = getSS();
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.getSheets().find(s => s.getName().trim().toLowerCase() === name.toLowerCase());
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
    sh.setFrozenRows(1);
    invalidateReadCache_();
  }
  return trackSheet_(sh);
}

/** อ่านชีตเป็น object ถ้าไม่มีชีตคืน [] (ไม่สร้างชีตใหม่) */
function readRowsIfSheet_(name) {
  try { return sheetToObjects_(name); } catch (e) { return []; }
}

/** เพิ่มแถวโดยจับคู่ตามชื่อหัวคอลัมน์ (ปลอดภัยแม้ลำดับคอลัมน์ในชีตต่างจากโค้ด) */
function appendByHeaders_(sheetName, obj) {
  const sh = getSheet_(sheetName);
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  sh.appendRow(headers.map(h => (h in obj ? obj[h] : '')));
}

/** เพิ่มคอลัมน์ UnitPrice ให้ CartridgeSales ที่มีอยู่แล้ว (ถ้ายังไม่มี) โดยไม่กระทบข้อมูลเดิม */
function ensureCartridgeSalesUnitPriceColumn_() {
  const sh = getSheet_(SHEET_CTG_SALES);
  const lastCol = sh.getLastColumn();
  const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  if (headers.indexOf('UnitPrice') !== -1) return;
  sh.getRange(1, lastCol + 1).setValue('UnitPrice').setFontWeight('bold');
  Logger.log('เพิ่มคอลัมน์ UnitPrice ในชีต CartridgeSales (แถวเดิมจะคำนวณราคาต่อตลับจาก จำนวนเงิน ÷ จำนวน ให้อัตโนมัติ)');
}

/** แปลง (จำนวน, ราคาต่อหน่วย, จำนวนเงินรวม) ให้ครบทั้งสามค่า — กรอกอย่างใดอย่างหนึ่งก็ได้ ระบบคำนวณอีกค่าให้ */
function shipResolvePrice_(qty, unitPrice, amount) {
  const has = v => v !== undefined && v !== null && String(v).trim() !== '' && !isNaN(Number(v));
  let up = has(unitPrice) ? Number(unitPrice) : null;
  let am = has(amount) ? Number(amount) : null;
  if ((up !== null && up < 0) || (am !== null && am < 0)) throw new Error('ราคา/จำนวนเงินต้องไม่ติดลบ');
  if (up === null && am === null) throw new Error('กรุณาระบุราคาต่อหน่วย หรือจำนวนเงินรวม (ถ้าไม่มีรายได้ให้ใส่ 0)');
  if (am === null) am = round2_(qty * up);
  if (up === null) up = qty > 0 ? round2_(am / qty) : 0;
  return { unitPrice: round2_(up), amount: round2_(am) };
}

// ---------- ทะเบียนผู้รับซื้อ ----------
/** คีย์เทียบชื่อผู้รับซื้อ: ตัดช่องว่าง/จุด/คำนำหน้า (ร้าน นาย นาง นางสาว บริษัท…) เพื่อให้ "ร้าน BB.Toner" = "ร้าน BB Toner" = "BB Toner" */
function buyerKey_(name) {
  const base = String(name || '').toLowerCase().replace(/[\s\.\-_,()]/g, '');
  const stripped = base.replace(/^(นางสาว|นส|นาง|นาย|คุณ|ร้าน|บริษัท|บจก|หจก)/, '');
  return stripped || base;
}

function isBuyerActive_(v) { return !(v === false || String(v).toUpperCase() === 'FALSE'); }

/** รายชื่อผู้รับซื้อทั้งหมดที่ระบบรู้จัก (ทะเบียน Buyers + ชื่อที่เคยบันทึกในรายการจำหน่าย/ส่งน้ำมันทุกโครงการ) รวมชื่อสะกดต่างกันให้เป็นรายเดียว */
function getBuyerList_() {
  const groups = {};
  function grp(key) {
    if (!groups[key]) groups[key] = { names: {}, registered: '', rounds: 0, projects: {}, last: null };
    return groups[key];
  }
  readRowsIfSheet_(SHEET_BUYERS).forEach(b => {
    const nm = String(b.BuyerName || '').trim();
    if (!nm || !isBuyerActive_(b.IsActive)) return;
    const x = grp(b.BuyerKey || buyerKey_(nm));
    if (!x.registered) x.registered = nm;
    String(b.Projects || '').split(',').forEach(t => { t = t.trim(); if (t) x.projects[t] = true; });
  });
  [[SHEET_CTG_SALES, 'SaleDate', 'ตลับหมึก'], [SHEET_FTF_SHIP, 'ShipmentDate', 'ทอดไม่ทิ้ง'], [SHEET_KIT_SHIP, 'ShipmentDate', 'จากครัวสู่ท้องฟ้า']].forEach(src => {
    readRowsIfSheet_(src[0]).forEach(r => {
      const nm = String(r.BuyerName || '').trim().replace(/\s+/g, ' ');
      if (!nm) return;
      const x = grp(buyerKey_(nm));
      x.names[nm] = (x.names[nm] || 0) + 1;
      x.rounds++;
      x.projects[src[2]] = true;
      const d = r[src[1]] ? new Date(r[src[1]]) : null;
      if (d && !isNaN(d.getTime()) && (!x.last || d > x.last)) x.last = d;
    });
  });
  return Object.keys(groups).map(key => {
    const x = groups[key];
    let name = x.registered;
    if (!name) {
      let best = -1;
      Object.keys(x.names).forEach(n => { if (x.names[n] > best) { best = x.names[n]; name = n; } });
    }
    return { BuyerName: name, BuyerKey: key, Rounds: x.rounds, Projects: Object.keys(x.projects).join(', '), LastDate: x.last ? toIsoDate_(x.last) : null };
  }).filter(b => b.BuyerName).sort((a, b) => b.Rounds - a.Rounds || a.BuyerName.localeCompare(b.BuyerName, 'th'));
}

function safeBuyerOptions_() {
  try { return getBuyerList_().map(b => ({ BuyerName: b.BuyerName, BuyerKey: b.BuyerKey, Rounds: b.Rounds })); } catch (e) { return []; }
}

/** ลงทะเบียนผู้รับซื้อลงชีต Buyers ถ้ายังไม่มี (หรือเพิ่มชื่อโครงการลงช่อง Projects) — คืน true ถ้าเพิ่มแถวใหม่ */
function registerBuyer_(name, key, tag) {
  ensureSheet_(SHEET_BUYERS, BUYER_HEADERS);
  const rows = sheetToObjects_(SHEET_BUYERS);
  const hit = rows.find(b => (b.BuyerKey || buyerKey_(b.BuyerName)) === key);
  if (hit) {
    const projects = String(hit.Projects || '').split(',').map(t => t.trim()).filter(Boolean);
    if (tag && projects.indexOf(tag) === -1) {
      projects.push(tag);
      const sh = getSheet_(SHEET_BUYERS);
      const col = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].indexOf('Projects');
      if (col !== -1) sh.getRange(hit.__row, col + 1).setValue(projects.join(', '));
    }
    return false;
  }
  appendByHeaders_(SHEET_BUYERS, { BuyerID: newId_('BUY'), BuyerName: name, BuyerKey: key, Phone: '', Projects: tag || '', Note: '', IsActive: true, CreatedAt: new Date() });
  return true;
}

/**
 * ใช้ตอนบันทึกการส่งจำหน่าย: ถ้าผู้รับซื้อรายนี้มีอยู่ในระบบแล้ว (แม้สะกด/คำนำหน้าต่างกันเล็กน้อย) จะใช้ชื่อเดิมในระบบให้อัตโนมัติ
 * ถ้าเป็นรายใหม่จะเพิ่มเข้าทะเบียนผู้รับซื้อ (ชีต Buyers) อัตโนมัติ — คืนชื่อที่ใช้บันทึกจริง
 */
function ensureBuyer_(name, projectTag) {
  name = String(name || '').trim().replace(/\s+/g, ' ');
  if (!name) return '';
  const key = buyerKey_(name);
  let canonical = name;
  try {
    const hit = getBuyerList_().find(b => b.BuyerKey === key);
    if (hit) canonical = hit.BuyerName;
  } catch (e) { /* ไม่มีรายชื่อเดิม — ใช้ชื่อที่กรอก */ }
  registerBuyer_(canonical, key, projectTag);
  return canonical;
}

/** นำรายชื่อผู้ซื้อที่เคยบันทึกไว้แล้วมาลงทะเบียนใน Buyers — ทำงานซ้ำได้อย่างปลอดภัย */
function ensureBuyersSeeded_() {
  ensureSheet_(SHEET_BUYERS, BUYER_HEADERS);
  let added = 0;
  getBuyerList_().forEach(b => {
    if (registerBuyer_(b.BuyerName, b.BuyerKey, b.Projects.split(',')[0].trim())) added++;
    b.Projects.split(',').slice(1).forEach(t => registerBuyer_(b.BuyerName, b.BuyerKey, t.trim()));
  });
  if (added) Logger.log('เพิ่มผู้รับซื้อในทะเบียน (ชีต Buyers): ' + added + ' ราย');
}

/** ทะเบียนผู้รับซื้อสำหรับหน้าเว็บ (สาธารณะ — ชื่อร้านที่เคยปรากฏในรายการจำหน่ายอยู่แล้ว) */
function getBuyers() {
  return getBuyerList_();
}

// ---------- มุมมองการส่งจำหน่ายรายรอบ (ใช้ร่วมกัน: ตลับหมึก / ทอดไม่ทิ้ง / จากครัวสู่ท้องฟ้า) ----------
/**
 * items: [{ID, Date, Buyer, Quantity, Unit, Base, UnitPrice, Amount, Note, RecordedByName}]
 *   Base = ปริมาณในหน่วยฐานของโครงการ (ลิตร หรือ ตลับ) — ราคาต่อหน่วยฐาน = จำนวนเงิน ÷ Base
 */
function buildShipView_(items, year) {
  // _ix = ลำดับแถวในชีต ใช้ตัดสินเมื่อวันที่เท่ากัน (แถวที่บันทึกทีหลังถือเป็นรอบล่าสุด)
  const valid = items.filter(i => i.Date instanceof Date && !isNaN(i.Date.getTime())).map((i, ix) => Object.assign({}, i, { _ix: ix }));
  const newestFirst = (a, b) => (b.Date - a.Date) || (b._ix - a._ix);
  const years = Array.from(new Set(valid.map(i => i.Date.getFullYear()))).sort((a, b) => b - a);
  const targetYear = Number(year) || years[0] || new Date().getFullYear();

  const stats = list => {
    const qty = list.reduce((s, i) => s + i.Base, 0);
    const amount = list.reduce((s, i) => s + i.Amount, 0);
    const priced = list.filter(i => i.Base > 0 && i.Amount > 0).map(i => i.Amount / i.Base);
    const last = list.slice().sort(newestFirst)[0];
    return {
      rounds: list.length, qty: round2_(qty), amount: round2_(amount),
      avgPrice: qty > 0 ? round2_(amount / qty) : 0,
      minPrice: priced.length ? round2_(Math.min.apply(null, priced)) : 0,
      maxPrice: priced.length ? round2_(Math.max.apply(null, priced)) : 0,
      lastDate: last ? toIsoDate_(last.Date) : null,
      lastPrice: last && last.Base > 0 ? round2_(last.Amount / last.Base) : 0
    };
  };

  const yearItems = valid.filter(i => i.Date.getFullYear() === targetYear);
  const monthly = Array.from({ length: 12 }, () => ({ qty: 0, amount: 0, rounds: 0 }));
  yearItems.forEach(i => {
    const m = monthly[i.Date.getMonth()];
    m.qty += i.Base; m.amount += i.Amount; m.rounds++;
  });
  monthly.forEach(m => { m.qty = round2_(m.qty); m.amount = round2_(m.amount); });

  const rounds = yearItems.slice().sort(newestFirst).map(i => ({
    ID: i.ID, Date: toIsoDate_(i.Date), Buyer: i.Buyer, Quantity: i.Quantity, Unit: i.Unit, Base: i.Base,
    UnitPrice: i.UnitPrice, PricePerBase: i.Base > 0 ? round2_(i.Amount / i.Base) : 0, Amount: i.Amount,
    Note: i.Note, RecordedByName: i.RecordedByName
  }));

  const yearly = years.map(y => Object.assign({ year: y }, stats(valid.filter(i => i.Date.getFullYear() === y))));

  const bg = {};
  valid.forEach(i => {
    const k = buyerKey_(i.Buyer) || i.Buyer;
    if (!bg[k]) bg[k] = { names: {}, list: [] };
    bg[k].names[i.Buyer] = (bg[k].names[i.Buyer] || 0) + 1;
    bg[k].list.push(i);
  });
  const buyers = Object.keys(bg).map(k => {
    let name = '', best = -1;
    Object.keys(bg[k].names).forEach(n => { if (bg[k].names[n] > best) { best = bg[k].names[n]; name = n; } });
    const st = stats(bg[k].list);
    return { BuyerName: name, rounds: st.rounds, qty: st.qty, amount: st.amount, avgPrice: st.avgPrice, lastDate: st.lastDate };
  }).sort((a, b) => b.amount - a.amount || b.rounds - a.rounds);

  return {
    year: targetYear, years: years.length ? years : [targetYear],
    rounds: rounds, monthly: monthly, yearly: yearly, buyers: buyers,
    yearStats: stats(yearItems), allTime: stats(valid)
  };
}

/** อ่านรายการส่งจำหน่ายจากชีตรูปแบบ SHIP_HEADERS แล้วแปลงเป็น items สำหรับ buildShipView_ */
function readShipItems_(sheetName) {
  return readRowsIfSheet_(sheetName).filter(r => r.ShipmentDate).map(r => {
    const qty = numOr0_(r.Quantity);
    const liters = (r.Liters !== '' && r.Liters != null && !isNaN(Number(r.Liters))) ? Number(r.Liters) : qty;
    const amount = numOr0_(r.Amount);
    const unitPrice = numOr0_(r.UnitPrice) > 0 ? numOr0_(r.UnitPrice) : (qty > 0 ? amount / qty : 0);
    return {
      ID: r.ShipmentID, Date: new Date(r.ShipmentDate), Buyer: String(r.BuyerName || ''), Quantity: qty,
      Unit: String(r.Unit || 'ลิตร'), Base: liters, UnitPrice: round2_(unitPrice), Amount: amount,
      Note: r.Note || '', RecordedByName: r.RecordedByName || ''
    };
  });
}

function shipNormalize_(d, opts) {
  const date = d.Date ? new Date(d.Date) : new Date();
  if (isNaN(date.getTime())) throw new Error('วันที่ไม่ถูกต้อง');
  const buyer = String(d.BuyerName || '').trim();
  if (!buyer) throw new Error('กรุณากรอกชื่อผู้รับซื้อ/ผู้รับ');
  const qty = Number(d.Quantity);
  if (!qty || qty <= 0) throw new Error('กรุณาระบุจำนวนให้ถูกต้อง');
  const unit = String(d.Unit || opts.units[0]).trim();
  if (opts.units.indexOf(unit) === -1) throw new Error('หน่วยไม่ถูกต้อง (เลือกได้: ' + opts.units.join(' / ') + ')');
  const liters = unit === 'ปี๊บ' ? round2_(qty * opts.litersPerCan) : qty;
  const price = shipResolvePrice_(qty, d.UnitPrice, d.Amount);
  return { date: date, buyer: buyer, qty: qty, unit: unit, liters: liters, unitPrice: price.unitPrice, amount: price.amount, note: String(d.Note || '').trim() };
}

function shipAdd_(sheetName, idPrefix, tag, auth, data, opts) {
  ensureSheet_(sheetName, SHIP_HEADERS);
  const n = shipNormalize_(data || {}, opts);
  const buyer = ensureBuyer_(n.buyer, tag);
  const now = new Date();
  const id = newId_(idPrefix);
  appendByHeaders_(sheetName, {
    ShipmentID: id, ShipmentDate: n.date, BuyerName: buyer, Quantity: n.qty, Unit: n.unit, Liters: n.liters,
    UnitPrice: n.unitPrice, Amount: n.amount, Note: n.note, RecordedByName: auth.departmentName, CreatedAt: now, UpdatedAt: now
  });
  return { ok: true, id: id, buyerName: buyer };
}

function shipUpdate_(sheetName, tag, id, data, opts) {
  data = data || {};
  const found = findRowById_(sheetName, 'ShipmentID', id);
  if (!found) throw new Error('ไม่พบรายการนี้');
  const cur = {};
  found.headers.forEach((h, i) => { cur[h] = found.values[i]; });
  const pick = (k, curKey) => (data[k] !== undefined ? data[k] : cur[curKey]);
  const n = shipNormalize_({
    Date: pick('Date', 'ShipmentDate'), BuyerName: pick('BuyerName', 'BuyerName'), Quantity: pick('Quantity', 'Quantity'),
    Unit: pick('Unit', 'Unit') || opts.units[0], UnitPrice: pick('UnitPrice', 'UnitPrice'), Amount: pick('Amount', 'Amount'), Note: pick('Note', 'Note')
  }, opts);
  const buyer = ensureBuyer_(n.buyer, tag);
  const fields = {
    ShipmentDate: n.date, BuyerName: buyer, Quantity: n.qty, Unit: n.unit, Liters: n.liters,
    UnitPrice: n.unitPrice, Amount: n.amount, Note: n.note, UpdatedAt: new Date()
  };
  const sh = getSheet_(sheetName);
  Object.keys(fields).forEach(key => {
    const col = found.headers.indexOf(key);
    if (col !== -1) sh.getRange(found.rowIndex, col + 1).setValue(fields[key]);
  });
  return { ok: true };
}

function shipDelete_(sheetName, id) {
  const found = findRowById_(sheetName, 'ShipmentID', id);
  if (!found) throw new Error('ไม่พบรายการนี้');
  getSheet_(sheetName).deleteRow(found.rowIndex);
  return { ok: true };
}

// ---------- ทอดไม่ทิ้ง: การส่งน้ำมันรายรอบ ----------
const FTF_SHIP_OPTS = { units: ['ลิตร'], litersPerCan: 1 };

/** การส่งน้ำมันของโครงการทอดไม่ทิ้งรายรอบ (จำนวนลิตร + ราคา) ของปีที่ระบุ พร้อมสรุปรายเดือน/รายปี/รายผู้รับซื้อ */
function getFryToFlyShipments(year) {
  const view = buildShipView_(readShipItems_(SHEET_FTF_SHIP), year);
  let collectedLiters = 0;
  try {
    const env = getEnvironmentalSettings_();
    collectedLiters = round2_(getFryToFlyData().summary.totalReturned * env.ftfLitersPerGallon);
  } catch (e) { /* ยังไม่มีข้อมูลการบริจาค */ }
  view.collectedLiters = collectedLiters;
  view.pendingLiters = Math.max(0, round2_(collectedLiters - view.allTime.qty));
  view.buyerOptions = safeBuyerOptions_();
  return view;
}

function addFryToFlyShipment(departmentId, pin, data) {
  const auth = assertAuthorized_(departmentId, pin);
  const res = shipAdd_(SHEET_FTF_SHIP, 'FTFS', 'ทอดไม่ทิ้ง', auth, data, FTF_SHIP_OPTS);
  syncFtfLinkedProject_();
  return res;
}
function updateFryToFlyShipment(departmentId, pin, id, data) {
  assertAuthorized_(departmentId, pin);
  const res = shipUpdate_(SHEET_FTF_SHIP, 'ทอดไม่ทิ้ง', id, data, FTF_SHIP_OPTS);
  syncFtfLinkedProject_();
  return res;
}
function deleteFryToFlyShipment(departmentId, pin, id) {
  assertAuthorized_(departmentId, pin);
  const res = shipDelete_(SHEET_FTF_SHIP, id);
  syncFtfLinkedProject_();
  return res;
}

/** ข้อความสถานะการส่งน้ำมัน ต่อท้ายสถานะโครงการทอดไม่ทิ้งในหน้าภาพรวม */
function ftfShipStatusText_() {
  try {
    const a = buildShipView_(readShipItems_(SHEET_FTF_SHIP)).allTime;
    if (!a.rounds) return '';
    return ' • ส่งน้ำมันแล้ว ' + a.rounds + ' รอบ ' + a.qty.toLocaleString() + ' ลิตร (รายได้ ' + a.amount.toLocaleString() + ' บาท)';
  } catch (e) { return ''; }
}

// ---------- จากครัวสู่ท้องฟ้า (ฝ่ายโภชนาการ) ----------
function kitchenOpts_() {
  return { units: KIT_UNITS, litersPerCan: getEnvironmentalSettings_().kitchenLitersPerCan };
}

/** ภาพรวมโครงการ "จากครัวสู่ท้องฟ้า": การส่งน้ำมันใช้แล้วรายรอบ (ลิตรหรือปี๊บ + ราคาแต่ละรอบ) ของปีที่ระบุ */
function getKitchenOverview(year) {
  const env = getEnvironmentalSettings_();
  const view = buildShipView_(readShipItems_(SHEET_KIT_SHIP), year);
  view.allTime.cans = round2_(view.allTime.qty / env.kitchenLitersPerCan);
  view.yearStats.cans = round2_(view.yearStats.qty / env.kitchenLitersPerCan);
  view.settings = { litersPerCan: env.kitchenLitersPerCan, co2PerLiterKg: env.kitchenCo2PerLiterKg, units: KIT_UNITS };
  view.co2SavedKg = round2_(view.allTime.qty * env.kitchenCo2PerLiterKg);
  view.buyerOptions = safeBuyerOptions_();
  return view;
}

function addKitchenShipment(departmentId, pin, data) {
  const auth = assertAuthorized_(departmentId, pin);
  const res = shipAdd_(SHEET_KIT_SHIP, 'KITS', 'จากครัวสู่ท้องฟ้า', auth, data, kitchenOpts_());
  syncKitchenLinkedProject_();
  return res;
}
function updateKitchenShipment(departmentId, pin, id, data) {
  assertAuthorized_(departmentId, pin);
  const res = shipUpdate_(SHEET_KIT_SHIP, 'จากครัวสู่ท้องฟ้า', id, data, kitchenOpts_());
  syncKitchenLinkedProject_();
  return res;
}
function deleteKitchenShipment(departmentId, pin, id) {
  assertAuthorized_(departmentId, pin);
  const res = shipDelete_(SHEET_KIT_SHIP, id);
  syncKitchenLinkedProject_();
  return res;
}

function syncKitchenLinkedProject_() {
  try {
    const a = buildShipView_(readShipItems_(SHEET_KIT_SHIP)).allTime;
    updateLinkedProjectFields_(LINKED_PROJECT_KIT_ID, {
      CurrentStatus: a.rounds
        ? `ส่งน้ำมันใช้แล้ว ${a.rounds} รอบ • ${a.qty.toLocaleString()} ลิตร • รายได้ ${a.amount.toLocaleString()} บาท • ราคาเฉลี่ย ${a.avgPrice.toLocaleString()} บาท/ลิตร`
        : 'เริ่มโครงการ — ยังไม่มีการส่งน้ำมัน (ดูรายละเอียดที่แท็บ "จากครัวสู่ท้องฟ้า")',
      Progress: yearProgressPct_()
    });
  } catch (e) { /* ยังไม่มีชีต/ข้อมูล ก็ข้ามไปเงียบๆ */ }
}

// ---------- ปฏิทินเก่าเราขอ (บริจาคทำอักษรเบรลล์) ----------
function calendarItems_() {
  return readRowsIfSheet_(SHEET_CAL_DON).filter(r => r.DonationDate).map(r => ({
    ID: r.DonationID, Date: new Date(r.DonationDate), Round: String(r.RoundName || '').trim() || 'ไม่ระบุรอบ',
    Type: String(r.DonorType || '').trim() || 'อื่นๆ', Name: String(r.DonorName || '').trim(), Kg: numOr0_(r.WeightKg),
    Note: r.Note || '', RecordedByName: r.RecordedByName || ''
  })).filter(i => !isNaN(i.Date.getTime()));
}

/** ภาพรวมโครงการปฏิทินเก่าเราขอของปีที่ระบุ: น้ำหนักรายรอบ/รายเดือน/รายประเภทผู้บริจาค/อันดับผู้บริจาค + ผลช่วยรักษาสิ่งแวดล้อม */
function getCalendarOverview(year) {
  const env = getEnvironmentalSettings_();
  const items = calendarItems_();
  const treesOf = kg => round2_(kg / 1000 * env.calendarTreesPerTon);
  const co2Of = kg => round2_(kg * env.calendarCo2PerKg);
  const dkey = n => buyerKey_(n) || n;

  const years = Array.from(new Set(items.map(i => i.Date.getFullYear()))).sort((a, b) => b - a);
  const targetYear = Number(year) || years[0] || new Date().getFullYear();
  const yearItems = items.filter(i => i.Date.getFullYear() === targetYear);

  const stats = list => {
    const kg = list.reduce((s, i) => s + i.Kg, 0);
    return {
      kg: round2_(kg), donations: list.length,
      donors: new Set(list.map(i => dkey(i.Name))).size,
      rounds: new Set(list.map(i => i.Round)).size,
      treesSaved: treesOf(kg), co2SavedKg: co2Of(kg)
    };
  };
  const typeNames = CAL_DONOR_TYPES.concat(Array.from(new Set(items.map(i => i.Type))).filter(t => CAL_DONOR_TYPES.indexOf(t) === -1));
  const byType = list => typeNames.map(t => {
    const l = list.filter(i => i.Type === t);
    return { type: t, kg: round2_(l.reduce((s, i) => s + i.Kg, 0)), donors: new Set(l.map(i => dkey(i.Name))).size, donations: l.length };
  }).filter(r => CAL_DONOR_TYPES.indexOf(r.type) !== -1 || r.donations > 0);

  const monthly = Array.from({ length: 12 }, () => ({ kg: 0, donations: 0 }));
  yearItems.forEach(i => { const m = monthly[i.Date.getMonth()]; m.kg += i.Kg; m.donations++; });
  monthly.forEach(m => { m.kg = round2_(m.kg); });

  const roundGroups = {};
  yearItems.forEach(i => {
    const g = roundGroups[i.Round] || (roundGroups[i.Round] = { RoundName: i.Round, first: i.Date, last: i.Date, list: [] });
    if (i.Date < g.first) g.first = i.Date;
    if (i.Date > g.last) g.last = i.Date;
    g.list.push(i);
  });
  const rounds = Object.keys(roundGroups).map(k => {
    const g = roundGroups[k];
    const bt = {};
    CAL_DONOR_TYPES.forEach(t => { bt[t] = round2_(g.list.filter(i => i.Type === t).reduce((s, i) => s + i.Kg, 0)); });
    const kg = g.list.reduce((s, i) => s + i.Kg, 0);
    return {
      RoundName: g.RoundName, firstDate: toIsoDate_(g.first), lastDate: toIsoDate_(g.last), donations: g.list.length,
      donors: new Set(g.list.map(i => dkey(i.Name))).size, kg: round2_(kg), byType: bt, treesSaved: treesOf(kg)
    };
  }).sort((a, b) => new Date(b.firstDate) - new Date(a.firstDate));

  const roundFirst = {};
  items.forEach(i => { if (!roundFirst[i.Round] || i.Date < roundFirst[i.Round]) roundFirst[i.Round] = i.Date; });
  const roundOptions = Object.keys(roundFirst).sort((a, b) => roundFirst[b] - roundFirst[a]);

  const dg = {};
  items.forEach(i => {
    const k = dkey(i.Name);
    if (!k) return;
    const g = dg[k] || (dg[k] = { names: {}, list: [] });
    g.names[i.Name] = (g.names[i.Name] || 0) + 1;
    g.list.push(i);
  });
  const donorLeaderboard = Object.keys(dg).map(k => {
    const g = dg[k];
    let name = '', best = -1;
    Object.keys(g.names).forEach(n => { if (g.names[n] > best) { best = g.names[n]; name = n; } });
    const latest = g.list.slice().sort((a, b) => b.Date - a.Date)[0];
    const kg = g.list.reduce((s, i) => s + i.Kg, 0);
    return {
      DonorName: name, DonorType: latest.Type, kg: round2_(kg), times: g.list.length,
      rounds: new Set(g.list.map(i => i.Round)).size, lastDate: toIsoDate_(latest.Date)
    };
  }).sort((a, b) => b.kg - a.kg).map((r, i) => Object.assign(r, { rank: i + 1 }));

  const donorOptions = {};
  CAL_DONOR_TYPES.forEach(t => { donorOptions[t] = []; });
  donorLeaderboard.forEach(d => { (donorOptions[d.DonorType] = donorOptions[d.DonorType] || []).push(d.DonorName); });
  try {
    const depts = sheetToObjects_(SHEET_CTG_DEPTS).filter(d => d.IsActive !== false).map(d => d.DepartmentName);
    donorOptions['หน่วยงาน'] = Array.from(new Set(donorOptions['หน่วยงาน'].concat(depts)));
  } catch (e) { /* ไม่มีรายชื่อหน่วยงาน */ }

  const yearlySummary = years.map(y => Object.assign({ year: y }, stats(items.filter(i => i.Date.getFullYear() === y))));
  const donations = yearItems.slice().sort((a, b) => b.Date - a.Date).map(i => ({
    DonationID: i.ID, DonationDate: toIsoDate_(i.Date), RoundName: i.Round, DonorType: i.Type, DonorName: i.Name,
    WeightKg: round2_(i.Kg), Note: i.Note
  }));
  const allStats = stats(items);
  const last = items.slice().sort((a, b) => b.Date - a.Date)[0];

  return {
    year: targetYear, years: years.length ? years : [targetYear], monthNames: PL_MONTH_NAMES,
    donorTypes: CAL_DONOR_TYPES, donations: donations, rounds: rounds, roundOptions: roundOptions,
    monthly: monthly, yearly: yearlySummary, donorLeaderboard: donorLeaderboard, donorOptions: donorOptions,
    typeYear: byType(yearItems), typeAll: byType(items), yearStats: stats(yearItems),
    allTime: Object.assign({ totalKg: allStats.kg, lastDate: last ? toIsoDate_(last.Date) : null }, allStats),
    settings: { treesPerTon: env.calendarTreesPerTon, co2PerKg: env.calendarCo2PerKg }
  };
}

function calendarNormalize_(d) {
  d = d || {};
  const date = d.DonationDate ? new Date(d.DonationDate) : new Date();
  if (isNaN(date.getTime())) throw new Error('วันที่ไม่ถูกต้อง');
  const type = String(d.DonorType || '').trim();
  if (CAL_DONOR_TYPES.indexOf(type) === -1) throw new Error('กรุณาเลือกประเภทผู้บริจาค');
  let name = String(d.DonorName || '').trim().replace(/\s+/g, ' ');
  if (!name) throw new Error('กรุณากรอกชื่อผู้บริจาค');
  const kg = Number(d.WeightKg);
  if (!kg || kg <= 0) throw new Error('กรุณาระบุน้ำหนักเป็นกิโลกรัมให้ถูกต้อง');
  const round = String(d.RoundName || '').trim() || ('รอบ ' + toIsoDate_(date).slice(0, 7));
  // ผู้บริจาคที่เคยบริจาคแล้ว (สะกด/คำนำหน้าต่างกันเล็กน้อย) ใช้ชื่อเดิมในระบบ
  const key = buyerKey_(name);
  const seen = {};
  calendarItems_().forEach(i => { if (buyerKey_(i.Name) === key) seen[i.Name] = (seen[i.Name] || 0) + 1; });
  let best = -1;
  Object.keys(seen).forEach(n => { if (seen[n] > best) { best = seen[n]; name = n; } });
  return { date: date, type: type, name: name, kg: round2_(kg), round: round, note: String(d.Note || '').trim() };
}

function addCalendarDonation(departmentId, pin, data) {
  const auth = assertAuthorized_(departmentId, pin);
  const n = calendarNormalize_(data);
  ensureSheet_(SHEET_CAL_DON, CAL_HEADERS);
  const now = new Date();
  const id = newId_('CAL');
  appendByHeaders_(SHEET_CAL_DON, {
    DonationID: id, DonationDate: n.date, RoundName: n.round, DonorType: n.type, DonorName: n.name, WeightKg: n.kg,
    Note: n.note, RecordedByName: auth.departmentName, CreatedAt: now, UpdatedAt: now
  });
  syncCalendarLinkedProject_();
  return { ok: true, id: id, donorName: n.name };
}

function updateCalendarDonation(departmentId, pin, donationId, data) {
  assertAuthorized_(departmentId, pin);
  data = data || {};
  const found = findRowById_(SHEET_CAL_DON, 'DonationID', donationId);
  if (!found) throw new Error('ไม่พบรายการบริจาคนี้');
  const cur = {};
  found.headers.forEach((h, i) => { cur[h] = found.values[i]; });
  const pick = (k) => (data[k] !== undefined ? data[k] : cur[k]);
  const n = calendarNormalize_({
    DonationDate: pick('DonationDate'), RoundName: pick('RoundName'), DonorType: pick('DonorType'),
    DonorName: pick('DonorName'), WeightKg: pick('WeightKg'), Note: pick('Note')
  });
  const fields = { DonationDate: n.date, RoundName: n.round, DonorType: n.type, DonorName: n.name, WeightKg: n.kg, Note: n.note, UpdatedAt: new Date() };
  const sh = getSheet_(SHEET_CAL_DON);
  Object.keys(fields).forEach(key => {
    const col = found.headers.indexOf(key);
    if (col !== -1) sh.getRange(found.rowIndex, col + 1).setValue(fields[key]);
  });
  syncCalendarLinkedProject_();
  return { ok: true };
}

function deleteCalendarDonation(departmentId, pin, donationId) {
  assertAuthorized_(departmentId, pin);
  const found = findRowById_(SHEET_CAL_DON, 'DonationID', donationId);
  if (!found) throw new Error('ไม่พบรายการบริจาคนี้');
  getSheet_(SHEET_CAL_DON).deleteRow(found.rowIndex);
  syncCalendarLinkedProject_();
  return { ok: true };
}

function syncCalendarLinkedProject_() {
  try {
    const a = getCalendarOverview().allTime;
    updateLinkedProjectFields_(LINKED_PROJECT_CAL_ID, {
      CurrentStatus: a.donations
        ? `บริจาคแล้ว ${a.rounds} รอบ • ${a.totalKg.toLocaleString()} กก. • ผู้ร่วมบริจาค ${a.donors} ราย • ช่วยรักษาต้นไม้ได้ประมาณ ${a.treesSaved.toLocaleString()} ต้น`
        : 'เริ่มโครงการ — ยังไม่มีการบริจาค (ดูรายละเอียดที่แท็บ "ปฏิทินเก่าเราขอ")',
      Progress: yearProgressPct_()
    });
  } catch (e) { /* ยังไม่มีชีต/ข้อมูล ก็ข้ามไปเงียบๆ */ }
}


// ============================= สรุปรายเดือน / รายปี / สะสม ของทุกโครงการ =============================
/**
 * สรุปผลตามช่วงเวลาของแต่ละโครงการ ในรูปแบบเดียวกันทุกโครงการ (ใช้กับการ์ด "สรุปรายเดือน • รายปี • สะสม")
 * key: 'ftf' | 'paperless' | 'cartridge' | 'kitchen' | 'calendar' | 'share'
 * คืนค่า: columns (นิยามคอลัมน์), monthly (12 เดือนของปีที่เลือก), yearly (ทุกปี), cumulative (ยอดสะสมถึงสิ้นปี),
 *         total (สะสมทั้งโครงการ), yearTotal (รวมของปีที่เลือก)
 * ตัวเลขคิดจากข้อมูลที่บันทึกจริงในแต่ละชีต (ไม่มีข้อมูลเพิ่มใหม่) — ค่า CO2/ต้นไม้ใช้ตัวแปรจาก EnvironmentalSettings
 */
function getPeriodSummary(key, year) {
  const env = getEnvironmentalSettings_();
  const safe = fn => { try { return fn(); } catch (e) { return []; } };
  const dk = n => buyerKey_(n) || String(n || '');
  const sumOf = (list, f) => list.reduce((s, i) => s + (Number(f(i)) || 0), 0);
  let columns = [], items = [], agg = null;

  if (key === 'ftf') {
    columns = [
      { key: 'gallons', label: 'น้ำมันที่รับบริจาค', unit: 'แกลลอน', dec: 1 },
      { key: 'liters', label: 'คิดเป็น', unit: 'ลิตร', dec: 1 },
      { key: 'donors', label: 'ผู้ร่วมบริจาค', unit: 'คน', dec: 0 },
      { key: 'shipRounds', label: 'ส่งออกไปแล้ว', unit: 'รอบ', dec: 0 },
      { key: 'shipLiters', label: 'ปริมาณที่ส่ง', unit: 'ลิตร', dec: 1 },
      { key: 'shipAmount', label: 'รายได้จากการส่ง', unit: 'บาท', dec: 0 },
      { key: 'co2', label: 'CO2 ที่ลดได้', unit: 'กก.', dec: 1 }
    ];
    safe(() => sheetToObjects_(SHEET_FTF_LOG)).filter(r => isVerified_(r.VerifyStatus)).forEach(r => {
      const d = r.Timestamp ? new Date(r.Timestamp) : null;
      if (d && !isNaN(d.getTime()) && r.ActionType === 'คืนแกลลอน') items.push({ d: d, t: 'don', gal: Number(r.Quantity) || 0, who: String(r.StaffCode || '') });
    });
    safe(() => readShipItems_(SHEET_FTF_SHIP)).forEach(s => { if (s.Date instanceof Date && !isNaN(s.Date.getTime())) items.push({ d: s.Date, t: 'ship', liters: s.Base, amount: s.Amount }); });
    agg = L => {
      const don = L.filter(i => i.t === 'don'), sh = L.filter(i => i.t === 'ship');
      const gal = sumOf(don, i => i.gal), liters = gal * env.ftfLitersPerGallon;
      return { gallons: gal, liters: liters, donors: new Set(don.filter(i => i.gal > 0).map(i => i.who)).size,
               shipRounds: sh.length, shipLiters: sumOf(sh, i => i.liters), shipAmount: sumOf(sh, i => i.amount), co2: liters * env.ftfCo2PerLiterKg };
    };
  } else if (key === 'paperless') {
    columns = [
      { key: 'depts', label: 'หน่วยงานที่รายงาน', unit: 'หน่วยงาน', dec: 0 },
      { key: 'items', label: 'เรื่องที่ไม่ต้องพิมพ์', unit: 'เรื่อง', dec: 0 },
      { key: 'sheets', label: 'กระดาษที่ลดได้', unit: 'แผ่น', dec: 0 },
      { key: 'trees', label: 'ต้นไม้ที่ช่วยรักษา', unit: 'ต้น', dec: 2 },
      { key: 'co2', label: 'CO2 ที่ลดได้', unit: 'กก.', dec: 1 }
    ];
    safe(() => sheetToObjects_(SHEET_PL_LOG)).forEach(r => {
      const y = Number(r.Year), m = Number(r.Month);
      if (!y || !m) return;
      items.push({ d: new Date(y, m - 1, 1), dept: String(r.DepartmentID || r.DepartmentName || ''), n: Number(r.ItemCount) || 0, sheets: Number(r.SheetsReduced) || 0 });
    });
    agg = L => {
      const sheets = sumOf(L, i => i.sheets);
      return { depts: new Set(L.filter(i => i.sheets > 0 || i.n > 0).map(i => i.dept)).size, items: sumOf(L, i => i.n), sheets: sheets,
               trees: sheets / env.paperSheetsPerTree, co2: sheets * env.paperCo2PerSheetKg };
    };
  } else if (key === 'cartridge') {
    const ctg = getCartridgeSettings_();
    columns = [
      { key: 'contributed', label: 'ส่งเข้าโครงการ', unit: 'ตลับ', dec: 0 },
      { key: 'soldQty', label: 'จำหน่ายแล้ว', unit: 'ตลับ', dec: 0 },
      { key: 'amount', label: 'รายได้', unit: 'บาท', dec: 0 },
      { key: 'avgPrice', label: 'ราคาเฉลี่ย', unit: 'บาท/ตลับ', dec: 2 },
      { key: 'weightKg', label: 'น้ำหนักรีไซเคิล', unit: 'กก.', dec: 1 },
      { key: 'co2', label: 'CO2 ที่ลดได้', unit: 'กก.', dec: 1 }
    ];
    safe(() => sheetToObjects_(SHEET_CTG_SALES)).forEach(r => {
      const d = r.SaleDate ? new Date(r.SaleDate) : null;
      if (d && !isNaN(d.getTime())) items.push({ d: d, t: 'sale', qty: Number(r.Quantity) || 0, amount: Number(r.Amount) || 0 });
    });
    safe(() => sheetToObjects_(SHEET_CTG_CONTRIB)).filter(r => isVerified_(r.VerifyStatus)).forEach(r => {
      const d = r.ContributionDate ? new Date(r.ContributionDate) : null;
      if (d && !isNaN(d.getTime())) items.push({ d: d, t: 'in', qty: Number(r.Quantity) || 0 });
    });
    agg = L => {
      const sales = L.filter(i => i.t === 'sale'), inn = L.filter(i => i.t === 'in');
      const qty = sumOf(sales, i => i.qty), amount = sumOf(sales, i => i.amount), w = qty * ctg.weightPerUnitKg;
      return { contributed: sumOf(inn, i => i.qty), soldQty: qty, amount: amount, avgPrice: qty > 0 ? amount / qty : 0, weightKg: w, co2: w * env.cartridgeCo2PerKg };
    };
  } else if (key === 'kitchen') {
    columns = [
      { key: 'rounds', label: 'จำนวนรอบที่ส่ง', unit: 'รอบ', dec: 0 },
      { key: 'liters', label: 'ปริมาณน้ำมัน', unit: 'ลิตร', dec: 1 },
      { key: 'cans', label: 'คิดเป็น', unit: 'ปี๊บ', dec: 1 },
      { key: 'amount', label: 'รายได้', unit: 'บาท', dec: 0 },
      { key: 'avgPrice', label: 'ราคาเฉลี่ย', unit: 'บาท/ลิตร', dec: 2 },
      { key: 'co2', label: 'CO2 ที่ลดได้', unit: 'กก.', dec: 1 }
    ];
    safe(() => readShipItems_(SHEET_KIT_SHIP)).forEach(s => { if (s.Date instanceof Date && !isNaN(s.Date.getTime())) items.push({ d: s.Date, liters: s.Base, amount: s.Amount }); });
    agg = L => {
      const liters = sumOf(L, i => i.liters), amount = sumOf(L, i => i.amount);
      return { rounds: L.length, liters: liters, cans: liters / env.kitchenLitersPerCan, amount: amount, avgPrice: liters > 0 ? amount / liters : 0, co2: liters * env.kitchenCo2PerLiterKg };
    };
  } else if (key === 'calendar') {
    columns = [
      { key: 'rounds', label: 'จำนวนรอบ', unit: 'รอบ', dec: 0 },
      { key: 'donors', label: 'ผู้บริจาค', unit: 'ราย', dec: 0 },
      { key: 'kg', label: 'น้ำหนักปฏิทินเก่า', unit: 'กก.', dec: 1 },
      { key: 'trees', label: 'ต้นไม้ที่ช่วยรักษา', unit: 'ต้น', dec: 2 },
      { key: 'co2', label: 'CO2 ที่ลดได้', unit: 'กก.', dec: 1 }
    ];
    safe(() => calendarItems_()).forEach(i => items.push({ d: i.Date, round: i.Round, who: dk(i.Name), kg: i.Kg }));
    agg = L => {
      const kg = sumOf(L, i => i.kg);
      return { rounds: new Set(L.map(i => i.round)).size, donors: new Set(L.map(i => i.who)).size, kg: kg,
               trees: kg / 1000 * env.calendarTreesPerTon, co2: kg * env.calendarCo2PerKg };
    };
  } else if (key === 'share') {
    columns = [
      { key: 'count', label: 'รายการที่แบ่งปัน', unit: 'รายการ', dec: 0 },
      { key: 'depts', label: 'หน่วยงานที่รับ', unit: 'หน่วยงาน', dec: 0 },
      { key: 'value', label: 'มูลค่าที่ประหยัดงบ', unit: 'บาท', dec: 0 },
      { key: 'reams', label: 'กระดาษ', unit: 'รีม', dec: 0 },
      { key: 'cartridges', label: 'ตลับหมึก/โทนเนอร์', unit: 'ตลับ', dec: 0 },
      { key: 'trees', label: 'ต้นไม้ที่ช่วยรักษา', unit: 'ต้น', dec: 2 },
      { key: 'co2', label: 'CO2 ที่ลดได้', unit: 'กก.', dec: 1 }
    ];
    safe(() => shareRows_(SHEET_SHARE_WD)).forEach(w => {
      const d = w.WithdrawDate ? new Date(w.WithdrawDate) : null;
      if (d && !isNaN(d.getTime())) items.push({ d: d, raw: w });
    });
    agg = L => {
      const raws = L.map(i => i.raw);
      const imp = getShareImpact_(raws);
      const value = raws.reduce((s, w) => s + (shareHasNum_(w.UnitPrice) ? shareNum_(w.Quantity) * Number(w.UnitPrice) : 0), 0);
      return { count: raws.length, depts: new Set(raws.map(w => shareText_(w.DepartmentName)).filter(Boolean)).size, value: value,
               reams: imp.paperReams, cartridges: imp.cartridges, trees: imp.treesSaved, co2: imp.totalCo2Kg };
    };
  } else if (key === 'em') {
    columns = [
      { key: 'batches', label: 'รอบที่เริ่มหมัก', unit: 'รอบ', dec: 0 },
      { key: 'barrels', label: 'จำนวนถังที่หมัก', unit: 'ถัง', dec: 0 },
      { key: 'inputKg', label: 'วัตถุดิบที่นำเข้ารวม', unit: 'กก.', dec: 1 },
      { key: 'foodKg', label: 'เศษอาหาร', unit: 'กก.', dec: 1 },
      { key: 'fruitVegKg', label: 'ผักและผลไม้', unit: 'กก.', dec: 1 },
      { key: 'meatKg', label: 'เนื้อสัตว์', unit: 'กก.', dec: 1 },
      { key: 'emLiters', label: 'น้ำยา EM ที่ผลิตได้', unit: 'ลิตร', dec: 1 },
      { key: 'usedLiters', label: 'น้ำยา EM ที่นำไปใช้', unit: 'ลิตร', dec: 1 }
    ];
    safe(() => readRowsIfSheet_(SHEET_EM_BATCHES)).filter(b => b.BatchID && (b.Status || EM_STATUS_FERMENT) !== EM_STATUS_CANCEL).forEach(b => {
      const sn = emDayNum_(b.StartDate);
      if (sn !== null) items.push({ d: new Date(sn * 86400000 + 12 * 3600000), t: 'start', barrels: Number(b.BarrelCount) || 0, kg: emInputKg_(b), food: Number(b.FoodKg) || 0, fv: Number(b.FruitVegKg) || 0, meat: Number(b.MeatKg) || 0 });
      const hn = emDayNum_(b.HarvestDate);
      if ((b.Status === EM_STATUS_DONE) && hn !== null) items.push({ d: new Date(hn * 86400000 + 12 * 3600000), t: 'har', l: Number(b.YieldLiters) || 0 });
    });
    safe(() => readRowsIfSheet_(SHEET_EM_USAGE)).filter(u => u.UsageID && isVerified_(u.VerifyStatus)).forEach(u => {
      const un = emDayNum_(u.UsageDate);
      if (un !== null) items.push({ d: new Date(un * 86400000 + 12 * 3600000), t: 'use', l: Number(u.Liters) || 0 });
    });
    agg = L => {
      const st = L.filter(i => i.t === 'start');
      return { batches: st.length, barrels: sumOf(st, i => i.barrels), inputKg: sumOf(st, i => i.kg), foodKg: sumOf(st, i => i.food), fruitVegKg: sumOf(st, i => i.fv), meatKg: sumOf(st, i => i.meat),
               emLiters: sumOf(L.filter(i => i.t === 'har'), i => i.l), usedLiters: sumOf(L.filter(i => i.t === 'use'), i => i.l) };
    };
  } else if (key === 'compost') {
    columns = [
      { key: 'batches', label: 'รอบที่เริ่มหมัก', unit: 'รอบ', dec: 0 },
      { key: 'piles', label: 'จำนวนกองที่หมัก', unit: 'กอง', dec: 0 },
      { key: 'inputKg', label: 'วัตถุดิบที่นำเข้ารวม', unit: 'กก.', dec: 1 },
      { key: 'dryKg', label: 'ใบไม้แห้ง', unit: 'กก.', dec: 1 },
      { key: 'freshKg', label: 'ใบไม้สด/เศษหญ้า', unit: 'กก.', dec: 1 },
      { key: 'branchKg', label: 'กิ่งไม้/เศษไม้', unit: 'กก.', dec: 1 },
      { key: 'cpKg', label: 'ปุ๋ยหมักที่ผลิตได้', unit: 'กก.', dec: 1 },
      { key: 'usedKg', label: 'ปุ๋ยหมักที่นำไปใช้', unit: 'กก.', dec: 1 }
    ];
    safe(() => readRowsIfSheet_(SHEET_CP_BATCHES)).filter(b => b.BatchID && (b.Status || CP_STATUS_FERMENT) !== CP_STATUS_CANCEL).forEach(b => {
      const sn = cpDayNum_(b.StartDate);
      if (sn !== null) items.push({ d: new Date(sn * 86400000 + 12 * 3600000), t: 'start', piles: Number(b.PileCount) || 0, kg: cpInputKg_(b), food: Number(b.DryKg) || 0, fv: Number(b.FreshKg) || 0, meat: Number(b.BranchKg) || 0 });
      const hn = cpDayNum_(b.HarvestDate);
      if ((b.Status === CP_STATUS_DONE) && hn !== null) items.push({ d: new Date(hn * 86400000 + 12 * 3600000), t: 'har', l: Number(b.YieldKg) || 0 });
    });
    safe(() => readRowsIfSheet_(SHEET_CP_USAGE)).filter(u => u.UsageID && isVerified_(u.VerifyStatus)).forEach(u => {
      const un = cpDayNum_(u.UsageDate);
      if (un !== null) items.push({ d: new Date(un * 86400000 + 12 * 3600000), t: 'use', l: Number(u.Kg) || 0 });
    });
    agg = L => {
      const st = L.filter(i => i.t === 'start');
      return { batches: st.length, piles: sumOf(st, i => i.piles), inputKg: sumOf(st, i => i.kg), dryKg: sumOf(st, i => i.food), freshKg: sumOf(st, i => i.fv), branchKg: sumOf(st, i => i.meat),
               cpKg: sumOf(L.filter(i => i.t === 'har'), i => i.l), usedKg: sumOf(L.filter(i => i.t === 'use'), i => i.l) };
    };
  } else {
    throw new Error('ไม่รู้จักโครงการ: ' + key);
  }

  const round = o => { const r = {}; Object.keys(o).forEach(k => { r[k] = round2_(o[k]); }); return r; };
  const years = Array.from(new Set(items.map(i => i.d.getFullYear()))).sort((a, b) => b - a);
  const targetYear = Number(year) || years[0] || new Date().getFullYear();
  const inYear = y => items.filter(i => i.d.getFullYear() === y);
  const monthly = Array.from({ length: 12 }, (_, m) => Object.assign({ month: m + 1 }, round(agg(inYear(targetYear).filter(i => i.d.getMonth() === m)))));
  const yearly = years.map(y => Object.assign({ year: y }, round(agg(inYear(y)))));
  const cumulative = years.slice().sort((a, b) => a - b).map(y => Object.assign({ year: y }, round(agg(items.filter(i => i.d.getFullYear() <= y)))));
  return {
    key: key, year: targetYear, years: years.length ? years : [targetYear], columns: columns,
    monthly: monthly, yearly: yearly, cumulative: cumulative.reverse(),
    yearTotal: round(agg(inYear(targetYear))), total: round(agg(items)), hasData: items.length > 0
  };
}


// ============================================================================================
// ผู้ช่วยตอบคำถามอัตโนมัติ (เฉพาะเรื่องระบบ) + ส่งอีเมลสอบถามเจ้าหน้าที่ + กล่องข้อความผู้ดูแล
// ============================================================================================
// - ผู้ช่วยเป็น "ระบบตอบตามกฎ" จากข้อมูลจริงในชีต (ไม่เรียก AI ภายนอก) จึงตอบได้เฉพาะเรื่องของระบบนี้เท่านั้น
//   คำถามนอกเรื่องจะถูกปฏิเสธอย่างสุภาพ และไม่เปิดเผยข้อมูลลับ (เช่น PIN) ไม่ว่าจะถามอย่างไร
// - ถ้าตอบไม่ได้ ผู้ถามส่งอีเมลสอบถามเจ้าหน้าที่ได้จากฟอร์ม → บันทึกในชีต Inquiries + ส่งอีเมลแจ้งผู้ดูแลระบบ
//   ผู้รับอีเมล: ชีต EnvironmentalSettings คีย์ InquiryEmail (คั่นหลายอีเมลด้วย ,) → ถ้าไม่ระบุใช้อีเมลของหน่วยงานที่เป็นผู้ดูแลระบบ
const SHEET_INQUIRIES = 'Inquiries';
const INQUIRY_HEADERS = ['InquiryID', 'CreatedAt', 'AskerType', 'AskerName', 'Organization', 'Contact', 'Topic', 'Question',
                         'Status', 'Reply', 'RepliedBy', 'RepliedAt', 'EmailSent'];
const INQUIRY_TYPES = ['เจ้าหน้าที่โรงพยาบาล', 'บริษัท/องค์กรภายนอก', 'บุคคลทั่วไป'];
const INQUIRY_STATUS_NEW = 'รอตอบ';
const INQUIRY_STATUS_DONE = 'ตอบแล้ว';

const ASSISTANT_CAMPAIGNS = [
  { key: 'ftf', name: 'ทอดไม่ทิ้ง', tab: 'frytofly', linked: LINKED_PROJECT_FTF_ID,
    re: /ทอดไม่ทิ้ง|แกลลอน|green\s*somdej|แต้ม|บางจาก|saf|น้ำมันพืช|น้ำมันสัตว์/i },
  { key: 'kitchen', name: 'จากครัวสู่ท้องฟ้า', tab: 'kitchen', linked: LINKED_PROJECT_KIT_ID,
    re: /จากครัวสู่ท้องฟ้า|ครัว|โภชนาการ|ปี๊บ|ปั๊บ/i },
  { key: 'paperless', name: 'Paperless (ลดการใช้กระดาษ)', tab: 'paperless', linked: LINKED_PROJECT_PL_ID,
    re: /paperless|เปเปอร์เลส|ลดกระดาษ|ลดการใช้กระดาษ|เอกสารอิเล็กทรอนิกส์/i },
  { key: 'cartridge', name: 'ตลับหมึกเก่าเราขอ', tab: 'cartridge', linked: LINKED_PROJECT_CTG_ID,
    re: /ตลับหมึก|หมึกพิมพ์|โทนเนอร์|toner/i },
  { key: 'calendar', name: 'ปฏิทินเก่าเราขอ', tab: 'calendar', linked: LINKED_PROJECT_CAL_ID,
    re: /ปฏิทิน|เบรลล์|เบลล์|braille|ผู้พิการทางสายตา|คนตาบอด/i },
  { key: 'share', name: 'คัดแยก แบ่งปัน', tab: 'share', linked: LINKED_PROJECT_SHARE_ID,
    re: /คัดแยก|แบ่งปัน|วัสดุเหลือใช้|เครื่องใช้สำนักงาน|เบิกวัสดุ|ขอเบิก/i },
  { key: 'em', name: 'น้ำยา EM จากเศษอาหาร', tab: 'em', linked: LINKED_PROJECT_EM_ID,
    re: /น้ำยา\s*em|\bem\b|อีเอ็ม|เศษอาหาร|หมักน้ำ|จุลินทรีย์|หมักเศษ/i },
  { key: 'compost', name: 'ปุ๋ยหมักจากขยะใบไม้', tab: 'compost', linked: LINKED_PROJECT_CP_ID,
    re: /ปุ๋ยหมัก|ขยะใบไม้|ใบไม้|เศษหญ้า|กิ่งไม้|หมักใบ/i }
];

const ASSISTANT_ABOUT = {
  ftf: 'แคมเปญขับเคลื่อนสิ่งแวดล้อมของกลุ่มบริษัทบางจาก — รวบรวมน้ำมันพืช/สัตว์ใช้แล้วจากบุคลากร เพื่อนำไปแปรรูปเป็นน้ำมันเชื้อเพลิงอากาศยานยั่งยืน (SAF) ลดการปล่อยก๊าซเรือนกระจกได้สูงสุดถึง 80% เป็นโครงการ "รับบริจาค" น้ำมันจากบุคลากร',
  kitchen: 'โครงการของฝ่ายโภชนาการ — รวบรวมน้ำมันทอดอาหารใช้แล้วจากครัวโรงพยาบาล ส่งให้ผู้รับซื้อเป็นรอบๆ เพื่อนำไปแปรรูปเป็นพลังงานทดแทน แทนการทิ้งลงท่อระบายน้ำ (ต่างจากทอดไม่ทิ้งที่เป็นการรับบริจาคจากบุคลากร)',
  paperless: 'แต่ละหน่วยงานบันทึกจำนวนเรื่องที่เปลี่ยนมาใช้ระบบอิเล็กทรอนิกส์ และปริมาณกระดาษที่ลดลงได้ (แผ่น) เป็นประจำทุกเดือน เพื่อติดตามภาพรวมการลดการใช้กระดาษของโรงพยาบาล',
  cartridge: 'รวบรวมตลับหมึกพิมพ์ใช้แล้วจากทุกหน่วยงาน ส่งขายให้ร้านรับซื้อเป็นรอบๆ ลดปริมาณขยะอันตรายที่ต้องเสียค่าทำลาย และสร้างรายได้ให้โรงพยาบาล',
  calendar: 'รับบริจาคปฏิทินเก่าจากหน่วยงาน เจ้าหน้าที่ และบริษัท/องค์กรภายนอก เป็นรอบๆ เพื่อส่งไปจัดทำอักษรเบรลล์ให้ผู้พิการทางสายตา นับเป็นน้ำหนัก (กก.) เพื่อคำนวณการช่วยรักษาสิ่งแวดล้อม',
  compost: 'นำขยะใบไม้ ใบไม้แห้ง/สด เศษหญ้า และกิ่งไม้ที่ตัดแต่งในพื้นที่โรงพยาบาลมาหมักเป็นปุ๋ยหมักตามรอบการหมัก 7 / 15 / 45 วัน เมื่อครบกำหนดจะเก็บเป็นปุ๋ยหมักแล้วนำไปใช้ปรับปรุงดิน ปลูกต้นไม้ในสวนหย่อมและแปลงผัก ระบบบันทึกจำนวนกอง ปริมาณปุ๋ยหมัก การนำไปใช้ และคงเหลือ',
  em: 'นำเศษอาหาร ผักและผลไม้ และเนื้อสัตว์มาหมักเป็นน้ำยา EM (จุลินทรีย์ที่มีประโยชน์) ตามรอบการหมัก 7 / 15 / 45 วัน เมื่อครบกำหนดจะเก็บเป็นน้ำยา EM แล้วนำไปใช้บำบัดน้ำเสีย ดับกลิ่น หรือปรับปรุงดินในพื้นที่ต่างๆ ระบบบันทึกจำนวนถัง ปริมาณน้ำยา การนำไปใช้ และคงเหลือ',
  share: 'ทุกหน่วยงานช่วยคัดแยกเครื่องใช้สำนักงานและวัสดุที่เกินจำเป็น ส่งเข้าโครงการ ผู้ดูแลตรวจสอบแล้วรับเข้าคลังรอแบ่งปัน เพื่อให้หน่วยงานอื่นนำไปใช้แทนการจัดซื้อใหม่ ลดการสั่งซื้อซ้ำซ้อน ลดของเหลือทิ้ง และประหยัดงบประมาณ'
};
const ASSISTANT_HOW = {
  ftf: 'นำน้ำมันพืช/สัตว์ที่ใช้แล้วมาที่จุดรับบริจาค ชั้น 2 อาคารบริการสนับสนุนทางการแพทย์ วันจันทร์–ศุกร์ เวลา 09.00–15.00 น. แล้วลงทะเบียนรับ/ส่งคืนแกลลอนในระบบ (ปุ่ม "ลงทะเบียนรับ / ส่งคืนแกลลอน" ในแท็บทอดไม่ทิ้ง หรือสแกน QR Code ที่จุดรับ) บริจาค 1 แกลลอน (≈ 4 ลิตร) ได้ 50 แต้ม',
  kitchen: 'เป็นโครงการภายในของฝ่ายโภชนาการ — ผู้ดูแลระบบเป็นผู้บันทึกรอบการส่งน้ำมัน (เป็นลิตรหรือปี๊บ) และราคาของแต่ละรอบ หากต้องการบริจาคน้ำมันใช้แล้วในฐานะบุคลากรทั่วไป ให้ดูโครงการ "ทอดไม่ทิ้ง"',
  paperless: 'หน่วยงานเข้าสู่ระบบด้วย PIN ของหน่วยงาน (แท็บ Paperless → "เข้าสู่ระบบหน่วยงาน") แล้วบันทึกจำนวนเรื่องและจำนวนแผ่นกระดาษที่ลดได้ประจำทุกเดือน หากไม่ทราบ PIN ให้สอบถามผู้ดูแลระบบ',
  cartridge: 'หน่วยงานส่งตลับหมึกใช้แล้วเข้าโครงการ และแจ้งจำนวนผ่านฟอร์ม "ส่งตลับหมึกเข้าโครงการ" (ปุ่มในแท็บตลับหมึกเก่าเราขอ หรือสแกน QR Code) ผู้ดูแลจะรวบรวมแล้วส่งขายให้ร้านรับซื้อเป็นรอบๆ',
  calendar: 'หน่วยงาน เจ้าหน้าที่ หรือบริษัท/องค์กรภายนอก นำปฏิทินเก่ามาบริจาคเป็นรอบๆ เจ้าหน้าที่ผู้ดูแลจะชั่งน้ำหนักและบันทึกชื่อผู้บริจาคในระบบ ปฏิทินที่ได้จะส่งไปจัดทำอักษรเบรลล์ (หากต้องการนัดส่งมอบ กรุณาส่งอีเมลสอบถามเจ้าหน้าที่)',
  compost: 'เป็นโครงการที่ผู้ดูแลระบบบันทึกรอบการหมัก (ปริมาณใบไม้ กิ่งไม้ จำนวนกอง วันที่เริ่ม) ผู้ปฏิบัติงานบันทึกการหมักประจำวันผ่านลิงก์/QR Code และตรวจอุณหภูมิกองทุก 7 วัน เมื่อครบกำหนดเก็บเป็นปุ๋ยหมักและบันทึกการนำไปใช้ — บุคลากรทุกคนดูความคืบหน้า ปริมาณคงเหลือ และผู้นำไปใช้ได้ที่แท็บ "ปุ๋ยหมักใบไม้"',
  em: 'เป็นโครงการที่ผู้ดูแลระบบบันทึกรอบการหมัก (ปริมาณเศษอาหาร จำนวนถัง วันที่เริ่ม) บันทึกผลตรวจทุก 7 วัน เก็บเป็นน้ำยา EM เมื่อครบกำหนด และบันทึกการนำไปใช้ — บุคลากรทุกคนดูความคืบหน้า ปริมาณคงเหลือ และผู้นำไปใช้ได้ที่แท็บ "น้ำยา EM"',
  share: 'ส่งรายการวัสดุที่เกินจำเป็นเข้าโครงการผ่านปุ่ม "ส่งรายการเข้าโครงการ" (หรือสแกน QR Code) ผู้ดูแลตรวจสอบแล้วรับเข้าคลัง จากนั้นหน่วยงานอื่นดูรายการคงเหลือในแท็บคัดแยก แบ่งปัน ก่อนตัดสินใจซื้อใหม่ และแจ้งผู้ดูแลระบบเพื่อบันทึกการเบิก/แบ่งปัน'
};

function assistantFmt_(n, d) {
  return (Number(n) || 0).toLocaleString('en-US', { maximumFractionDigits: d === undefined ? 2 : d });
}

function assistantProjectRow_(c) {
  try { return sheetToObjects_(SHEET_PROJECTS).find(p => p.ProjectID === c.linked) || null; } catch (e) { return null; }
}

function assistantStats_(key) {
  let t = null;
  try { t = getPeriodSummary(key).total; } catch (e) { return ''; }
  if (!t) return '';
  if (key === 'ftf') {
    let people = 0;
    try { people = getFryToFlyData().summary.totalParticipants; } catch (e) { /* ไม่มีข้อมูล */ }
    return `สะสมถึงปัจจุบัน: รับบริจาคน้ำมัน ${assistantFmt_(t.gallons, 1)} แกลลอน (≈ ${assistantFmt_(t.liters, 1)} ลิตร) จากผู้ร่วมโครงการ ${assistantFmt_(people, 0)} คน • ส่งออกไปแล้ว ${assistantFmt_(t.shipRounds, 0)} รอบ (${assistantFmt_(t.shipLiters, 1)} ลิตร) • ลด CO2 ≈ ${assistantFmt_(t.co2, 1)} กก.`;
  }
  if (key === 'kitchen') return `สะสมถึงปัจจุบัน: ส่งน้ำมันแล้ว ${assistantFmt_(t.rounds, 0)} รอบ รวม ${assistantFmt_(t.liters, 1)} ลิตร (≈ ${assistantFmt_(t.cans, 1)} ปี๊บ) • รายได้ ${assistantFmt_(t.amount, 0)} บาท • ลด CO2 ≈ ${assistantFmt_(t.co2, 1)} กก.`;
  if (key === 'paperless') return `สะสมถึงปัจจุบัน: ลดกระดาษได้ ${assistantFmt_(t.sheets, 0)} แผ่น จาก ${assistantFmt_(t.items, 0)} เรื่อง • เทียบเท่าต้นไม้ ${assistantFmt_(t.trees, 2)} ต้น • ลด CO2 ≈ ${assistantFmt_(t.co2, 1)} กก.`;
  if (key === 'cartridge') return `สะสมถึงปัจจุบัน: จำหน่ายตลับหมึกแล้ว ${assistantFmt_(t.soldQty, 0)} ตลับ (≈ ${assistantFmt_(t.weightKg, 1)} กก.) • รายได้ ${assistantFmt_(t.amount, 0)} บาท • ลด CO2 ≈ ${assistantFmt_(t.co2, 1)} กก.`;
  if (key === 'calendar') return `สะสมถึงปัจจุบัน: รับบริจาค ${assistantFmt_(t.rounds, 0)} รอบ จากผู้บริจาค ${assistantFmt_(t.donors, 0)} ราย รวม ${assistantFmt_(t.kg, 1)} กก. • เทียบเท่าต้นไม้ที่ช่วยรักษา ${assistantFmt_(t.trees, 2)} ต้น`;
  if (key === 'compost') return `สะสมถึงปัจจุบัน: หมักแล้ว ${assistantFmt_(t.batches, 0)} รอบ (${assistantFmt_(t.piles, 0)} กอง) จากใบไม้และกิ่งไม้ ${assistantFmt_(t.inputKg, 1)} กก. • ได้ปุ๋ยหมัก ${assistantFmt_(t.cpKg, 1)} กก. • นำไปใช้แล้ว ${assistantFmt_(t.usedKg, 1)} กก.`;
  if (key === 'em') return `สะสมถึงปัจจุบัน: หมักแล้ว ${assistantFmt_(t.batches, 0)} รอบ (${assistantFmt_(t.barrels, 0)} ถัง) จากเศษอาหาร ${assistantFmt_(t.inputKg, 1)} กก. • ได้น้ำยา EM ${assistantFmt_(t.emLiters, 1)} ลิตร • นำไปใช้แล้ว ${assistantFmt_(t.usedLiters, 1)} ลิตร`;
  if (key === 'share') return `สะสมถึงปัจจุบัน: แบ่งปันแล้ว ${assistantFmt_(t.count, 0)} รายการ (${assistantFmt_(t.depts, 0)} หน่วยงาน) • มูลค่าที่ประหยัดงบ ${assistantFmt_(t.value, 0)} บาท`;
  return '';
}

function assistantPrice_(key) {
  let a = null, unit = 'บาท/ลิตร';
  try {
    if (key === 'ftf') a = getFryToFlyShipments().allTime;
    else if (key === 'kitchen') a = getKitchenOverview().allTime;
    else if (key === 'cartridge') { a = getCartridgeOverview().shipments.allTime; unit = 'บาท/ตลับ'; }
  } catch (e) { a = null; }
  if (!a || !a.rounds) return 'ยังไม่มีข้อมูลการส่งจำหน่ายที่บันทึกไว้ในระบบ';
  return `ส่งจำหน่ายแล้ว ${assistantFmt_(a.rounds, 0)} รอบ ราคาแต่ละรอบไม่เท่ากัน • รอบล่าสุด (${a.lastDate || '-'}) ราคา ${assistantFmt_(a.lastPrice)} ${unit} • ราคาเฉลี่ยตลอดโครงการ ${assistantFmt_(a.avgPrice)} ${unit} (ต่ำสุด ${assistantFmt_(a.minPrice)} – สูงสุด ${assistantFmt_(a.maxPrice)}) • รายได้รวม ${assistantFmt_(a.amount, 0)} บาท`;
}

function assistantStatus_(c) {
  const p = assistantProjectRow_(c);
  if (!p) return `ยังไม่พบข้อมูลสถานะของโครงการ "${c.name}" ในระบบ`;
  return `สถานะโครงการ "${c.name}": ${p.CurrentStatus || 'ยังไม่มีข้อมูลสถานะ'} (ความคืบหน้า ${p.Progress || 0}%, หมวดสถานะ: ${p.StatusCategory || '-'})`;
}

function assistantContact_(c) {
  if (c.key === 'ftf') return 'ผู้ประสานงานโครงการทอดไม่ทิ้ง: นางสาวสุชาดา รัตนบัวแก้ว โทร. 038-320200 ต่อ 20510 • จุดรับบริจาค ชั้น 2 อาคารบริการสนับสนุนทางการแพทย์ วันจันทร์–ศุกร์ 09.00–15.00 น.';
  const p = assistantProjectRow_(c);
  const who = p ? `${p.ResponsiblePerson || ''} ${p.ResponsibleUnit || p.Department || ''}`.trim() : '';
  return (who ? `โครงการ "${c.name}" ดูแลโดย ${who}` : `ยังไม่มีข้อมูลผู้รับผิดชอบโครงการ "${c.name}" ในระบบ`) + ' — หากต้องการติดต่อโดยตรง กดปุ่ม "ส่งอีเมลสอบถามเจ้าหน้าที่" ด้านล่างได้เลยครับ';
}

/** เลือกคำตอบของโครงการหนึ่งตามเจตนาของคำถาม */
function assistantCampaignAnswer_(c, q, intents) {
  if (intents.contact) return assistantContact_(c);
  if (intents.points && c.key === 'ftf') {
    return `บริจาคน้ำมัน 1 แกลลอน (≈ 4 ลิตร) ได้ ${FTF_POINTS_PER_GALLON} แต้ม • ระดับสมาชิกตามแต้มสะสม: ` +
      FTF_TIERS.slice().reverse().map(t => `${t.label} (ตั้งแต่ ${t.min} แต้ม)`).join(' → ') + ' • ดูอันดับและแต้มของตัวเองได้ที่แท็บทอดไม่ทิ้ง';
  }
  if (intents.price) return `[${c.name}] ` + (c.key === 'ftf' || c.key === 'kitchen' || c.key === 'cartridge' ? assistantPrice_(c.key) : 'โครงการนี้ไม่มีการจำหน่ายเป็นรอบที่คิดราคา');
  if (intents.stats) return `[${c.name}] ` + (assistantStats_(c.key) || 'ยังไม่มีข้อมูลสะสม');
  if (intents.status) return assistantStatus_(c);
  if (intents.join) return `วิธีเข้าร่วมโครงการ "${c.name}": ${ASSISTANT_HOW[c.key]}`;
  const st = assistantStats_(c.key);
  return `โครงการ "${c.name}": ${ASSISTANT_ABOUT[c.key]}\n\nวิธีเข้าร่วม: ${ASSISTANT_HOW[c.key]}` + (st ? `\n\n${st}` : '');
}

const ASSISTANT_OFFTOPIC_TEXT = 'ขออภัยครับ ผู้ช่วยนี้ตอบได้เฉพาะเรื่องเกี่ยวกับ "ระบบติดตามโครงการสิ่งแวดล้อมโรงพยาบาล" เท่านั้น ' +
  '(เช่น มีโครงการอะไรบ้าง สถานะโครงการ วิธีเข้าร่วม/บริจาค ผลลัพธ์ด้านสิ่งแวดล้อม การใช้งานระบบ) ลองถามเรื่องเหล่านี้ได้เลยครับ หรือส่งอีเมลสอบถามเจ้าหน้าที่';

/**
 * ผู้ช่วยตอบคำถามอัตโนมัติ — ใครก็ถามได้ (ไม่ต้องมี PIN) ตอบจากข้อมูลจริงในระบบเท่านั้น
 * คืนค่า { answer, kind: 'answer'|'offtopic'|'unknown'|'refused', suggestEmail, tabs:[{label,tab}], followUps:[...] }
 */
function askAssistant(question) {
  const q = String(question || '').replace(/\s+/g, ' ').trim().slice(0, 400);
  if (!q) throw new Error('กรุณาพิมพ์คำถาม');
  const res = (answer, kind, extra) => Object.assign({ answer: answer, kind: kind, suggestEmail: kind === 'unknown' || kind === 'offtopic', tabs: [], followUps: [] }, extra || {});

  // 1) ขอข้อมูลลับ (PIN/รหัสผ่าน) → ปฏิเสธเสมอ
  if (/(pin|พิน|รหัสผ่าน|password|รหัสผู้ดูแล)/i.test(q) && /(บอก|ขอ|ส่ง|ให้|อะไร|คือ|ลืม|เท่าไร|ค่าเริ่มต้น|default)/i.test(q)) {
    return res('ขออภัยครับ ผมไม่สามารถแจ้งหรือกู้คืน PIN/รหัสผ่านได้ หากลืม PIN หรือต้องการสิทธิ์ผู้ดูแลระบบ กรุณาติดต่อผู้ดูแลระบบของโรงพยาบาลโดยตรง (ส่งอีเมลสอบถามเจ้าหน้าที่ได้ที่ปุ่มด้านล่าง)', 'refused', { suggestEmail: true });
  }

  const intents = {
    greet: /^(สวัสดี|หวัดดี|ดีครับ|ดีค่ะ|hello|hi|hey)/i.test(q),
    thanks: /(ขอบคุณ|ขอบใจ|thank)/i.test(q),
    help: /(ทำอะไรได้|ช่วยอะไร|ใช้งานยังไง|ถามอะไรได้|help|คุณคือใคร|เป็นใคร)/i.test(q),
    list: /(มีโครงการ|โครงการอะไร|โครงการทั้งหมด|รายชื่อโครงการ|กี่โครงการ|โครงการอะไรบ้าง)/.test(q),
    join: /(เข้าร่วม|ร่วมโครงการ|ร่วมกิจกรรม|บริจาค|นำมา|ส่งมา|ส่งยังไง|ส่งอย่างไร|ทำอย่างไร|ทำยังไง|วิธี|ขั้นตอน|ที่ไหน|จุดรับ|เวลา|เปิดรับ|ลงทะเบียน|สมัคร)/.test(q),
    status: /(สถานะ|คืบหน้า|ไปถึงไหน|progress|status|เสร็จหรือยัง)/i.test(q),
    stats: /(กี่|เท่าไร|เท่าไหร่|สะสม|ยอด|ทั้งหมด|จำนวน|ปริมาณ|น้ำหนัก|สถิติ|ผลลัพธ์|ผลงาน)/.test(q),
    contact: /(ติดต่อ|โทร|เบอร์|อีเมล|email|ผู้ประสานงาน|ผู้รับผิดชอบ|ใครดูแล|ถามใคร)/i.test(q),
    price: /(ราคา|รายได้|ขายได้|ขายเท่า|กี่บาท|ราคาต่อ)/.test(q),
    points: /(แต้ม|ระดับสมาชิก|ระดับทอง|ระดับเงิน|ทองแดง|สิทธิประโยชน์)/.test(q),
    env: /(ต้นไม้|co2|co₂|คาร์บอน|ก๊าซเรือนกระจก|โลกร้อน|สิ่งแวดล้อม|ผลกระทบ)/i.test(q),
    external: /(บริษัทภายนอก|องค์กรภายนอก|บุคคลภายนอก|คนนอก|ภายนอกโรงพยาบาล|บริษัท)/.test(q),
    login: /(เข้าสู่ระบบ|login|ล็อกอิน|ผู้ดูแลระบบ|แอดมิน|admin|บันทึกข้อมูล|แก้ไขข้อมูล)/i.test(q),
    exportDoc: /(ส่งออก|pdf|excel|เอ็กเซล|พิมพ์รายงาน|ดาวน์โหลด)/i.test(q),
    qr: /(qr|คิวอาร์|ลิงก์|แชร์)/i.test(q),
    photo: /(รูปภาพ|ภาพกิจกรรม|อัปโหลดรูป|เพิ่มรูป|แกลเลอรี)/i.test(q)
  };

  if (intents.greet && q.length <= 20) {
    return res('สวัสดีครับ 😊 ผมเป็นผู้ช่วยอัตโนมัติของระบบติดตามโครงการสิ่งแวดล้อมโรงพยาบาล ถามเรื่องโครงการ สถานะ วิธีเข้าร่วม ผลลัพธ์ด้านสิ่งแวดล้อม หรือการใช้งานระบบได้เลยครับ', 'answer', { followUps: ['มีโครงการอะไรบ้าง', 'วิธีบริจาคน้ำมันใช้แล้ว', 'ผลลัพธ์ด้านสิ่งแวดล้อมรวม'] });
  }
  if (intents.thanks && q.length <= 30) return res('ยินดีครับ 🌱 มีอะไรอยากถามเพิ่มเติมเกี่ยวกับโครงการได้เลย', 'answer');
  if (intents.help) {
    return res('ผมตอบได้เฉพาะเรื่องของระบบนี้ เช่น\n• มีโครงการอะไรบ้าง และสถานะเป็นอย่างไร\n• วิธีเข้าร่วม/บริจาคของแต่ละโครงการ (ทอดไม่ทิ้ง, Paperless, ตลับหมึก, จากครัวสู่ท้องฟ้า, ปฏิทินเก่า, คัดแยก แบ่งปัน)\n• ยอดสะสม ราคา และผลลัพธ์ด้านสิ่งแวดล้อม (ต้นไม้/CO2)\n• การใช้งานระบบ เช่น ส่งออกรายงาน ลิงก์/QR Code\nถ้าผมตอบไม่ได้ ส่งอีเมลสอบถามเจ้าหน้าที่ได้ที่ปุ่มด้านล่างครับ', 'answer', { followUps: ['มีโครงการอะไรบ้าง', 'ติดต่อใคร'] });
  }

  const matched = ASSISTANT_CAMPAIGNS.filter(c => c.re.test(q));

  // 2) รายชื่อโครงการทั้งหมด
  if (intents.list && !matched.length) {
    let rows = [];
    try { rows = sheetToObjects_(SHEET_PROJECTS).filter(p => !isTruthyFlag_(p.IsArchived)); } catch (e) { rows = []; }
    if (!rows.length) return res('ตอนนี้ยังไม่มีโครงการในระบบครับ', 'answer');
    const lines = rows.slice(0, 20).map(p => `• ${p.ProjectName} — ${p.StatusCategory || 'ไม่ระบุสถานะ'} (${p.Progress || 0}%)${p.Department ? ' • ' + p.Department : ''}`);
    return res(`ขณะนี้มี ${rows.length} โครงการในระบบ:\n${lines.join('\n')}\n\nถามต่อเรื่องโครงการไหนก็ได้ เช่น "วิธีเข้าร่วม…" หรือ "สถานะ…"`, 'answer',
      { tabs: [{ label: 'เปิดภาพรวมโครงการ', tab: 'projects' }], followUps: ['วิธีบริจาคน้ำมันใช้แล้ว', 'ผลลัพธ์ด้านสิ่งแวดล้อมรวม'] });
  }

  // 3) คำถามที่ระบุโครงการ
  if (matched.length) {
    // "น้ำมัน" ทั้งสองโครงการ: ถ้าถามทั้งคู่ให้ตอบทั้งคู่แบบสั้น
    const list = matched.slice(0, 2);
    const parts = list.map(c => assistantCampaignAnswer_(c, q, intents));
    const tabs = list.map(c => ({ label: 'เปิดแท็บ ' + c.name, tab: c.tab }));
    const c0 = list[0];
    return res(parts.join('\n\n— — —\n\n'), 'answer', {
      tabs: tabs,
      followUps: list.length === 1 ? ['วิธีเข้าร่วม' + c0.name, 'สถานะ' + c0.name, 'ยอดสะสม' + c0.name, 'ติดต่อ' + c0.name].slice(0, 4) : []
    });
  }

  // 4) บริษัท/บุคคลภายนอก
  if (intents.external && /(บริจาค|ร่วม|ส่ง|เข้าร่วม|ได้ไหม|ได้หรือไม่|สนใจ)/.test(q)) {
    return res('บริษัท/องค์กรภายนอก ร่วมโครงการ "ปฏิทินเก่าเราขอ" ได้ โดยบริจาคปฏิทินเก่าเพื่อนำไปจัดทำอักษรเบรลล์ให้ผู้พิการทางสายตา (ระบบบันทึกประเภทผู้บริจาคเป็น "บริษัท/องค์กรภายนอก") ส่วนโครงการอื่นๆ ส่วนใหญ่เป็นกิจกรรมภายในโรงพยาบาล หากสนใจร่วมโครงการใดหรือต้องการนัดหมายส่งมอบ กรุณาส่งอีเมลสอบถามเจ้าหน้าที่ที่ปุ่มด้านล่าง ระบุชื่อบริษัทและช่องทางติดต่อได้เลยครับ', 'answer',
      { suggestEmail: true, tabs: [{ label: 'เปิดแท็บ ปฏิทินเก่าเราขอ', tab: 'calendar' }] });
  }

  // 5) ผลลัพธ์สิ่งแวดล้อมรวม
  if (intents.env) {
    try {
      const o = getEnvironmentalImpactSummary().overview;
      return res(`ผลลัพธ์ด้านสิ่งแวดล้อมรวมทุกโครงการ (ประมาณ): ลด CO2 ได้ ≈ ${assistantFmt_(o.totalCo2SavedKg, 1)} กก. (${assistantFmt_(o.totalCo2SavedTon, 2)} ตัน) • ต้นไม้ที่ช่วยรักษาได้ ${assistantFmt_(o.treesSavedDirect, 2)} ต้น (นับเฉพาะกระดาษ) • เทียบเท่าต้นไม้ดูดซับ CO2 ใน 1 ปี ≈ ${assistantFmt_(o.treesEquivalentFromCo2, 0)} ต้น\nตัวเลขเป็นค่าประมาณจากสมมติฐานทั่วไป ดูวิธีคำนวณละเอียดได้ที่หน้ารายงานสิ่งแวดล้อม`, 'answer',
        { tabs: [{ label: 'เปิดรายงานสิ่งแวดล้อม', tab: 'envreport' }], followUps: ['ต้นไม้ที่ช่วยรักษาได้คืออะไร'] });
    } catch (e) {
      return res('ยังไม่สามารถดึงข้อมูลผลลัพธ์ด้านสิ่งแวดล้อมได้ในขณะนี้ครับ ลองเปิดหน้ารายงานสิ่งแวดล้อมในระบบ', 'answer', { tabs: [{ label: 'เปิดรายงานสิ่งแวดล้อม', tab: 'envreport' }] });
    }
  }

  // 6) การใช้งานระบบ
  if (intents.login) {
    return res('ทุกคนดูข้อมูลโครงการได้โดยไม่ต้องเข้าสู่ระบบ ส่วนการเพิ่ม/แก้ไข/บันทึกข้อมูลทำได้เฉพาะผู้ดูแลระบบ โดยกดปุ่ม "เข้าสู่ระบบผู้ดูแล" มุมขวาบน เลือกหน่วยงานและกรอก PIN (Paperless ใช้ PIN ของแต่ละหน่วยงาน) หากลืม PIN หรืออยากขอสิทธิ์ ให้ติดต่อผู้ดูแลระบบ', 'answer', { suggestEmail: true });
  }
  if (intents.exportDoc) return res('ส่งออกได้ทั้ง Excel และ PDF (A4 แนวนอน) จากปุ่มในหน้าของแต่ละโครงการ และหน้ารายงานสิ่งแวดล้อม นอกจากนี้การ์ด "สรุปรายเดือน • รายปี • สะสม" ของแต่ละโครงการก็มีปุ่ม Excel/PDF ครับ', 'answer');
  if (intents.qr) return res('แท็บทอดไม่ทิ้ง, ตลับหมึกเก่าเราขอ และคัดแยก แบ่งปัน มีปุ่ม "แชร์ลิงก์ / QR Code" สำหรับให้บุคลากรสแกนส่งข้อมูลเข้าโครงการได้ทันที สามารถดาวน์โหลด QR Code ไปติดที่จุดรับได้', 'answer');
  if (intents.photo) return res('ผู้ดูแลระบบเพิ่มรูปกิจกรรมได้ที่ส่วน "ภาพกิจกรรม" ใต้หัวข้อของแต่ละโครงการ (เข้าสู่ระบบผู้ดูแลก่อน) ทุกคนคลิกรูปเพื่อดูภาพขยายได้ครับ', 'answer');

  // 7) โครงการทั่วไปในชีต Projects (ที่ผู้ดูแลเพิ่มเอง) — จับคู่จากชื่อโครงการ
  let rows = [];
  try { rows = sheetToObjects_(SHEET_PROJECTS).filter(p => !isTruthyFlag_(p.IsArchived)); } catch (e) { rows = []; }
  const norm = s => String(s || '').toLowerCase().replace(/^โครงการ\s*/, '').replace(/[\s"'“”()\-_.]/g, '');
  const nq = norm(q);
  const hit = rows.filter(p => !/-LINK$/.test(String(p.ProjectID))).find(p => { const n = norm(p.ProjectName); return n.length >= 3 && nq.indexOf(n) !== -1; });
  if (hit) {
    const a = tryAutoAnswerProjectQuestion_(hit, q);
    const base = a || `โครงการ "${hit.ProjectName}" (${hit.Category || 'ไม่ระบุประเภท'}): ${hit.Description || 'ยังไม่มีรายละเอียด'} • สถานะ: ${hit.CurrentStatus || '-'} (${hit.Progress || 0}%) • ผู้รับผิดชอบ: ${hit.ResponsiblePerson || '-'} (${hit.ResponsibleUnit || hit.Department || '-'})`;
    return res(base, 'answer', { tabs: [{ label: 'เปิดภาพรวมโครงการ', tab: 'projects' }] });
  }

  // 8) คำถามทั่วไปที่ยังเกี่ยวกับระบบแต่ตอบไม่ได้ vs นอกเรื่อง
  if (intents.join || intents.status || intents.stats || intents.contact || intents.price || intents.points ||
      /(โครงการ|ระบบ|โรงพยาบาล|รพ\.|ขยะ|รีไซเคิล|น้ำมัน|กระดาษ|หมึก|ปฏิทิน|บริจาค|แบ่งปัน|รายงาน|หน่วยงาน|กิจกรรม)/.test(q)) {
    return res('ขออภัยครับ ผมยังไม่มีข้อมูลเรื่องนี้ในระบบ ลองระบุชื่อโครงการให้ชัดขึ้น เช่น "ทอดไม่ทิ้ง", "Paperless", "ตลับหมึก", "จากครัวสู่ท้องฟ้า", "ปฏิทินเก่า", "คัดแยก แบ่งปัน" หรือส่งอีเมลสอบถามเจ้าหน้าที่ที่ปุ่มด้านล่างครับ', 'unknown', { followUps: ['มีโครงการอะไรบ้าง'] });
  }
  return res(ASSISTANT_OFFTOPIC_TEXT, 'offtopic', { followUps: ['มีโครงการอะไรบ้าง', 'วิธีบริจาคน้ำมันใช้แล้ว'] });
}

// ----------------------------- ส่งอีเมลสอบถามเจ้าหน้าที่ -----------------------------
const INQUIRY_EMAIL_RE = /^[^\s@,;]+@[^\s@,;]+\.[^\s@,;]+$/;

function getInquiryRecipients_() {
  let list = [];
  try {
    const row = sheetToObjects_(SHEET_ENV_SETTINGS).find(r => r.Key === 'InquiryEmail');
    if (row && row.Value) list = String(row.Value).split(/[,;\s]+/).filter(x => INQUIRY_EMAIL_RE.test(x));
  } catch (e) { /* ไม่มีชีต */ }
  if (!list.length) {
    try {
      list = sheetToObjects_(SHEET_DEPARTMENTS)
        .filter(d => isTruthyFlag_(d.IsAdmin) && isTruthyFlag_(d.IsActive !== undefined ? d.IsActive : true) && d.DepartmentEmail && INQUIRY_EMAIL_RE.test(String(d.DepartmentEmail).trim()))
        .map(d => String(d.DepartmentEmail).trim());
    } catch (e) { list = []; }
  }
  if (!list.length) {
    try { const me = Session.getEffectiveUser().getEmail(); if (me && INQUIRY_EMAIL_RE.test(me)) list = [me]; } catch (e) { /* ไม่มีสิทธิ์อ่านอีเมลเจ้าของสคริปต์ */ }
  }
  return Array.from(new Set(list));
}

/**
 * ส่งคำถาม/สอบถามถึงเจ้าหน้าที่ (เปิดให้ทุกคน ไม่ต้องมี PIN) — บันทึกลงชีต Inquiries และส่งอีเมลแจ้งผู้ดูแลระบบ
 * data: { AskerType, AskerName, Organization, Contact, Topic, Question, website(ช่องดักสแปม ต้องว่าง) }
 */
function submitInquiry(data) {
  data = data || {};
  if (String(data.website || '').trim()) throw new Error('ไม่สามารถส่งได้');
  const type = String(data.AskerType || '').trim();
  if (INQUIRY_TYPES.indexOf(type) === -1) throw new Error('กรุณาเลือกประเภทผู้สอบถาม');
  const name = String(data.AskerName || '').trim().slice(0, 100);
  if (!name) throw new Error('กรุณากรอกชื่อผู้สอบถาม');
  const org = String(data.Organization || '').trim().slice(0, 150);
  const contact = String(data.Contact || '').trim().slice(0, 120);
  if (!contact) throw new Error('กรุณากรอกช่องทางติดต่อกลับ (อีเมลหรือเบอร์โทร)');
  const isEmail = INQUIRY_EMAIL_RE.test(contact);
  if (!isEmail && !/^[0-9+\-\s()#ต่อ.]{6,}$/.test(contact)) throw new Error('ช่องทางติดต่อต้องเป็นอีเมลหรือเบอร์โทรที่ถูกต้อง');
  const topic = String(data.Topic || 'ทั่วไป').trim().slice(0, 150) || 'ทั่วไป';
  const question = String(data.Question || '').trim().slice(0, 1500);
  if (question.length < 5) throw new Error('กรุณาพิมพ์คำถามให้ละเอียดขึ้นเล็กน้อย');
  if ((type === 'บริษัท/องค์กรภายนอก') && !org) throw new Error('กรุณาระบุชื่อบริษัท/องค์กร');

  // กันส่งซ้ำ/สแปม: ผู้ติดต่อเดิมส่งซ้ำภายใน 60 วินาที และจำกัดจำนวนรวมต่อชั่วโมง
  try {
    if (typeof CacheService !== 'undefined') {
      const cache = CacheService.getScriptCache(), ck = 'inq:' + contact.toLowerCase();
      if (cache.get(ck)) throw new Error('คุณเพิ่งส่งคำถามไป กรุณารอสักครู่แล้วลองใหม่');
      cache.put(ck, '1', 60);
    }
  } catch (e) { if (/เพิ่งส่ง/.test(String(e.message))) throw e; }
  ensureSheet_(SHEET_INQUIRIES, INQUIRY_HEADERS);
  const recent = readRowsIfSheet_(SHEET_INQUIRIES).filter(r => r.CreatedAt && (Date.now() - new Date(r.CreatedAt).getTime()) < 3600 * 1000).length;
  if (recent >= 40) throw new Error('มีการส่งคำถามจำนวนมากในขณะนี้ กรุณาลองใหม่ภายหลัง');

  const now = new Date(), id = newId_('INQ');
  appendByHeaders_(SHEET_INQUIRIES, {
    InquiryID: id, CreatedAt: now, AskerType: type, AskerName: name, Organization: org, Contact: contact, Topic: topic,
    Question: question, Status: INQUIRY_STATUS_NEW, Reply: '', RepliedBy: '', RepliedAt: '', EmailSent: false
  });

  let emailed = false;
  const to = getInquiryRecipients_();
  if (to.length) {
    try {
      const opts = {
        to: to.join(','),
        subject: `[สอบถามระบบสิ่งแวดล้อม] ${topic} — ${name} (${type})`,
        body: `มีผู้สอบถามผ่านระบบติดตามโครงการสิ่งแวดล้อมโรงพยาบาล\n\n` +
          `ประเภท: ${type}\nชื่อ: ${name}\nหน่วยงาน/บริษัท: ${org || '-'}\nติดต่อกลับ: ${contact}\nเรื่อง/โครงการ: ${topic}\n\n` +
          `คำถาม:\n${question}\n\n` +
          `รหัสคำถาม: ${id}\n(ผู้ดูแลระบบดูและตอบได้ในระบบ: ปุ่มผู้ช่วย 💬 → "กล่องข้อความ")`
      };
      if (isEmail) opts.replyTo = contact;
      MailApp.sendEmail(opts);
      emailed = true;
    } catch (e) { Logger.log('ส่งอีเมลแจ้งคำถามไม่สำเร็จ: ' + e); }
  }
  if (emailed) {
    const f = findRowById_(SHEET_INQUIRIES, 'InquiryID', id);
    if (f) { const col = f.headers.indexOf('EmailSent'); if (col !== -1) getSheet_(SHEET_INQUIRIES).getRange(f.rowIndex, col + 1).setValue(true); }
  }
  return { ok: true, id: id, emailed: emailed };
}

/** ผู้ดูแลระบบ: รายการสอบถามทั้งหมด (ใหม่สุดก่อน) */
function getInquiries(departmentId, pin) {
  assertAuthorized_(departmentId, pin);
  return readRowsIfSheet_(SHEET_INQUIRIES).map(r => ({
    InquiryID: r.InquiryID, CreatedAt: toIsoDate_(r.CreatedAt), AskerType: r.AskerType, AskerName: r.AskerName, Organization: r.Organization,
    Contact: r.Contact, Topic: r.Topic, Question: r.Question, Status: r.Status || INQUIRY_STATUS_NEW, Reply: r.Reply || '',
    RepliedBy: r.RepliedBy || '', RepliedAt: toIsoDate_(r.RepliedAt), EmailSent: r.EmailSent === true || r.EmailSent === 'TRUE'
  })).sort((a, b) => (b.CreatedAt > a.CreatedAt ? 1 : b.CreatedAt < a.CreatedAt ? -1 : 0) || (b.InquiryID > a.InquiryID ? 1 : -1));
}

/** ผู้ดูแลระบบ: ตอบกลับ (ถ้าช่องทางติดต่อเป็นอีเมลและเลือกส่ง จะส่งอีเมลตอบให้ผู้ถามด้วย) */
function replyInquiry(departmentId, pin, inquiryId, replyText, sendEmail) {
  const auth = assertAuthorized_(departmentId, pin);
  const reply = String(replyText || '').trim();
  if (!reply) throw new Error('กรุณากรอกข้อความตอบกลับ');
  const found = findRowById_(SHEET_INQUIRIES, 'InquiryID', inquiryId);
  if (!found) throw new Error('ไม่พบคำถามนี้');
  const get = h => found.values[found.headers.indexOf(h)];
  const contact = String(get('Contact') || '').trim();
  let emailed = false;
  if (sendEmail && INQUIRY_EMAIL_RE.test(contact)) {
    try {
      MailApp.sendEmail({
        to: contact,
        subject: `ตอบกลับ: ${get('Topic') || 'คำถามเกี่ยวกับโครงการสิ่งแวดล้อมโรงพยาบาล'}`,
        body: `เรียน คุณ${get('AskerName') || ''}\n\n${reply}\n\n— ${auth.departmentName} (ระบบติดตามโครงการสิ่งแวดล้อมโรงพยาบาล)\n\n` +
          `[คำถามเดิมของคุณ]\n${get('Question') || ''}`
      });
      emailed = true;
    } catch (e) { Logger.log('ส่งอีเมลตอบกลับไม่สำเร็จ: ' + e); }
  }
  const sh = getSheet_(SHEET_INQUIRIES);
  const fields = { Status: INQUIRY_STATUS_DONE, Reply: reply, RepliedBy: auth.departmentName, RepliedAt: new Date() };
  Object.keys(fields).forEach(k => { const col = found.headers.indexOf(k); if (col !== -1) sh.getRange(found.rowIndex, col + 1).setValue(fields[k]); });
  return { ok: true, emailed: emailed };
}

/** รายการหัวข้อให้เลือกในฟอร์มสอบถาม (โครงการทั้งหมด) */
function getInquiryTopics() {
  const names = ['ทั่วไป / เรื่องอื่นๆ', 'การใช้งานระบบ'].concat(ASSISTANT_CAMPAIGNS.map(c => c.name));
  try {
    sheetToObjects_(SHEET_PROJECTS).filter(p => !isTruthyFlag_(p.IsArchived) && !/-LINK$/.test(String(p.ProjectID))).forEach(p => { if (p.ProjectName) names.push(String(p.ProjectName)); });
  } catch (e) { /* ไม่มีโครงการ */ }
  return Array.from(new Set(names));
}


// ============================================================================================
// ภาพกิจกรรมของแต่ละโครงการ (เก็บไฟล์ใน Google Drive, เก็บรายการในชีต ProjectPhotos)
// ============================================================================================
// - เฉพาะผู้ดูแลระบบ (PIN) เพิ่ม/แก้ไข/ลบรูปได้ ทุกคนดูได้
// - ไฟล์เก็บในโฟลเดอร์ Drive ชื่อ "ภาพกิจกรรม-ระบบติดตามโครงการสิ่งแวดล้อม" (สร้างอัตโนมัติ, id เก็บใน Script Properties)
//   แชร์แบบ "ทุกคนที่มีลิงก์ดูได้" เพื่อให้แสดงบนหน้าเว็บได้ (ถ้านโยบายองค์กรห้ามแชร์ จะอัปโหลดไม่สำเร็จ)
// - ต้องอนุญาตสิทธิ์ Drive ครั้งแรก: รันฟังก์ชัน setupPhotos() จาก Apps Script editor 1 ครั้ง
const SHEET_PHOTOS = 'ProjectPhotos';
const PHOTO_HEADERS = ['PhotoID', 'ProjectKey', 'FileID', 'Caption', 'SortOrder', 'UploadedBy', 'CreatedAt'];
const PHOTO_FOLDER_NAME = 'ภาพกิจกรรม-ระบบติดตามโครงการสิ่งแวดล้อม';
const PHOTO_FOLDER_PROP = 'PHOTO_FOLDER_ID';
const PHOTO_MAX_BYTES = 4 * 1024 * 1024;      // จำกัดไฟล์ละ 4 MB (ฝั่งหน้าเว็บย่อรูปก่อนอัปโหลดอยู่แล้ว)
const PHOTO_MAX_PER_PROJECT = 60;
const PHOTO_MIMES = ['image/jpeg', 'image/png', 'image/webp', 'image/gif'];
const PHOTO_CAMPAIGN_KEYS = ['overview', 'ftf', 'paperless', 'cartridge', 'kitchen', 'calendar', 'share', 'em', 'compost'];

/** รันครั้งเดียวจาก editor เพื่อสร้างโฟลเดอร์ + ขอสิทธิ์ Drive */
function setupPhotos() {
  ensureSheet_(SHEET_PHOTOS, PHOTO_HEADERS);
  const f = getPhotoFolder_();
  Logger.log('โฟลเดอร์ภาพกิจกรรม: ' + f.getUrl());
  return f.getUrl();
}

function getPhotoFolder_() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty(PHOTO_FOLDER_PROP);
  if (id) { try { return DriveApp.getFolderById(id); } catch (e) { /* โฟลเดอร์ถูกลบ → สร้างใหม่ */ } }
  const folder = DriveApp.createFolder(PHOTO_FOLDER_NAME);
  props.setProperty(PHOTO_FOLDER_PROP, folder.getId());
  return folder;
}

function isValidPhotoKey_(key) {
  key = String(key || '');
  if (PHOTO_CAMPAIGN_KEYS.indexOf(key) !== -1) return true;
  try { return !!sheetToObjects_(SHEET_PROJECTS).find(p => p.ProjectID === key); } catch (e) { return false; }
}

function photoRowToObj_(r) {
  return { PhotoID: r.PhotoID, ProjectKey: r.ProjectKey, FileID: r.FileID, Caption: r.Caption || '',
           SortOrder: Number(r.SortOrder) || 0, UploadedBy: r.UploadedBy || '', CreatedAt: toIsoDate_(r.CreatedAt) };
}

/** รายการรูปของโครงการ (สาธารณะ) — เรียงตามลำดับ แล้วใหม่สุดก่อน; รูปแรก = รูปปก */
function getProjectPhotos(projectKey) {
  const key = String(projectKey || '');
  return readRowsIfSheet_(SHEET_PHOTOS).filter(r => r.ProjectKey === key && r.FileID)
    .map(photoRowToObj_)
    .sort((a, b) => (a.SortOrder - b.SortOrder) || (a.CreatedAt < b.CreatedAt ? 1 : a.CreatedAt > b.CreatedAt ? -1 : 0));
}

/** รูปของทุกโครงการ ในครั้งเดียว: { key: [photos] } (ลดจำนวนการเรียกเซิร์ฟเวอร์ตอนเปิดหน้า) */
function getAllProjectPhotos() {
  const out = {};
  readRowsIfSheet_(SHEET_PHOTOS).filter(r => r.FileID).map(photoRowToObj_).forEach(p => { (out[p.ProjectKey] = out[p.ProjectKey] || []).push(p); });
  Object.keys(out).forEach(k => out[k].sort((a, b) => (a.SortOrder - b.SortOrder) || (a.CreatedAt < b.CreatedAt ? 1 : a.CreatedAt > b.CreatedAt ? -1 : 0)));
  return out;
}

/** ผู้ดูแลระบบ: อัปโหลดรูป (base64 จากหน้าเว็บ) */
function uploadProjectPhoto(departmentId, pin, projectKey, fileName, mimeType, base64Data, caption) {
  const auth = assertAuthorized_(departmentId, pin);
  if (!isValidPhotoKey_(projectKey)) throw new Error('ไม่พบโครงการที่ต้องการเพิ่มรูป');
  const mime = String(mimeType || '').toLowerCase();
  if (PHOTO_MIMES.indexOf(mime) === -1) throw new Error('รองรับเฉพาะไฟล์รูปภาพ JPG, PNG, WEBP, GIF');
  const b64 = String(base64Data || '').replace(/^data:[^,]*,/, '');
  if (!b64) throw new Error('ไม่พบข้อมูลรูปภาพ');
  if (b64.length * 0.75 > PHOTO_MAX_BYTES) throw new Error('ไฟล์รูปใหญ่เกิน 4 MB');
  ensureSheet_(SHEET_PHOTOS, PHOTO_HEADERS);
  const existing = readRowsIfSheet_(SHEET_PHOTOS).filter(r => r.ProjectKey === projectKey);
  if (existing.length >= PHOTO_MAX_PER_PROJECT) throw new Error('โครงการนี้มีรูปครบ ' + PHOTO_MAX_PER_PROJECT + ' รูปแล้ว กรุณาลบรูปเก่าก่อน');

  const ext = { 'image/jpeg': 'jpg', 'image/png': 'png', 'image/webp': 'webp', 'image/gif': 'gif' }[mime];
  const safeName = String(fileName || 'photo').replace(/[^\wก-๙.\- ]/g, '_').replace(/\.[A-Za-z0-9]+$/, '').slice(0, 60) || 'photo';
  const id = newId_('PHO');
  let file;
  try {
    const blob = Utilities.newBlob(Utilities.base64Decode(b64), mime, projectKey + '_' + id + '_' + safeName + '.' + ext);
    file = getPhotoFolder_().createFile(blob);
  } catch (e) {
    throw new Error('บันทึกรูปลง Google Drive ไม่สำเร็จ (ตรวจสอบว่าได้อนุญาตสิทธิ์ Drive แล้ว): ' + (e.message || e));
  }
  try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); }
  catch (e) {
    try { file.setTrashed(true); } catch (e2) { /* ignore */ }
    throw new Error('ตั้งค่าแชร์รูปไม่สำเร็จ — นโยบายองค์กรอาจไม่อนุญาตการแชร์ลิงก์ ติดต่อผู้ดูแล Google Workspace');
  }
  const maxOrder = existing.reduce((m, r) => Math.max(m, Number(r.SortOrder) || 0), 0);
  appendByHeaders_(SHEET_PHOTOS, {
    PhotoID: id, ProjectKey: projectKey, FileID: file.getId(), Caption: String(caption || '').trim().slice(0, 200),
    SortOrder: maxOrder + 1, UploadedBy: auth.departmentName, CreatedAt: new Date()
  });
  return { ok: true, photo: photoRowToObj_({ PhotoID: id, ProjectKey: projectKey, FileID: file.getId(), Caption: caption, SortOrder: maxOrder + 1, UploadedBy: auth.departmentName, CreatedAt: new Date() }) };
}

function setPhotoCell_(found, field, value) {
  const col = found.headers.indexOf(field);
  if (col !== -1) getSheet_(SHEET_PHOTOS).getRange(found.rowIndex, col + 1).setValue(value);
}

function updatePhotoCaption(departmentId, pin, photoId, caption) {
  assertAuthorized_(departmentId, pin);
  const found = findRowById_(SHEET_PHOTOS, 'PhotoID', photoId);
  if (!found) throw new Error('ไม่พบรูปภาพ');
  setPhotoCell_(found, 'Caption', String(caption || '').trim().slice(0, 200));
  return { ok: true };
}

/** ตั้งรูปนี้เป็นรูปปก (ลำดับแรก) */
function setPhotoCover(departmentId, pin, photoId) {
  assertAuthorized_(departmentId, pin);
  const found = findRowById_(SHEET_PHOTOS, 'PhotoID', photoId);
  if (!found) throw new Error('ไม่พบรูปภาพ');
  const key = found.values[found.headers.indexOf('ProjectKey')];
  const minOrder = readRowsIfSheet_(SHEET_PHOTOS).filter(r => r.ProjectKey === key).reduce((m, r) => Math.min(m, Number(r.SortOrder) || 0), 0);
  setPhotoCell_(found, 'SortOrder', minOrder - 1);
  return { ok: true };
}

function deleteProjectPhoto(departmentId, pin, photoId) {
  assertAuthorized_(departmentId, pin);
  const found = findRowById_(SHEET_PHOTOS, 'PhotoID', photoId);
  if (!found) throw new Error('ไม่พบรูปภาพ');
  const fileId = found.values[found.headers.indexOf('FileID')];
  try { if (fileId) DriveApp.getFileById(fileId).setTrashed(true); } catch (e) { Logger.log('ลบไฟล์ใน Drive ไม่สำเร็จ: ' + e); }
  getSheet_(SHEET_PHOTOS).deleteRow(found.rowIndex);
  return { ok: true };
}


// ============================================================================================
// สวิตช์ "ใช้ข้อมูลหน่วยงาน" รายโครงการ (เปิด/ปิดได้จากหน้าเว็บโดยผู้ดูแลระบบ ไม่ต้องแก้โค้ด)
// ============================================================================================
// - ชีต ProjectSettings: ProjectKey, UseDepartments (TRUE/FALSE), UpdatedBy, UpdatedAt
// - ค่าเริ่มต้น = เปิด (ทำงานเหมือนเดิม) โครงการที่ปิดจะไม่ถามหน่วยงานและไม่แสดงอันดับหน่วยงาน
const SHEET_PROJECT_SETTINGS = 'ProjectSettings';
const PROJECT_SETTINGS_HEADERS = ['ProjectKey', 'UseDepartments', 'UpdatedBy', 'UpdatedAt'];
const DEPT_MODE_PROJECTS = ['paperless', 'cartridge', 'share', 'calendar'];
const NO_DEPT_TEXT = '(ไม่ระบุหน่วยงาน)';
const PL_ALL_ID = 'PLD-ALL';
const PL_ALL_NAME = 'ภาพรวมทั้งโรงพยาบาล';

function getProjectModes_() {
  const modes = {};
  DEPT_MODE_PROJECTS.forEach(k => { modes[k] = { useDepartments: true }; });
  readRowsIfSheet_(SHEET_PROJECT_SETTINGS).forEach(r => {
    if (modes[r.ProjectKey]) modes[r.ProjectKey].useDepartments = !(r.UseDepartments === false || String(r.UseDepartments).toUpperCase() === 'FALSE');
  });
  return modes;
}
function isDeptMode_(key) {
  const m = getProjectModes_()[key];
  return m ? m.useDepartments : true;
}
/** สาธารณะ: ค่าสวิตช์ของทุกโครงการ */
function getProjectModes() { return getProjectModes_(); }

/** ผู้ดูแลระบบ: เปิด/ปิดการใช้ข้อมูลหน่วยงานของโครงการ */
function setProjectUseDepartments(departmentId, pin, projectKey, useDepartments) {
  const auth = assertAuthorized_(departmentId, pin);
  if (DEPT_MODE_PROJECTS.indexOf(projectKey) === -1) throw new Error('โครงการนี้ไม่มีการตั้งค่าหน่วยงาน');
  ensureSheet_(SHEET_PROJECT_SETTINGS, PROJECT_SETTINGS_HEADERS);
  const val = useDepartments === true || useDepartments === 'true';
  const found = findRowById_(SHEET_PROJECT_SETTINGS, 'ProjectKey', projectKey);
  const now = new Date();
  if (found) {
    const sh = getSheet_(SHEET_PROJECT_SETTINGS);
    const set = (h, v) => { const c = found.headers.indexOf(h); if (c !== -1) sh.getRange(found.rowIndex, c + 1).setValue(v); };
    set('UseDepartments', val); set('UpdatedBy', auth.departmentName); set('UpdatedAt', now);
  } else {
    appendByHeaders_(SHEET_PROJECT_SETTINGS, { ProjectKey: projectKey, UseDepartments: val, UpdatedBy: auth.departmentName, UpdatedAt: now });
  }
  if (projectKey === 'paperless' && !val) ensurePaperlessAllDept_();
  return { ok: true, modes: getProjectModes_() };
}

/** Paperless โหมดไม่ใช้หน่วยงาน: ใช้แถว "ภาพรวมทั้งโรงพยาบาล" แถวเดียวเก็บยอดรวมรายเดือน */
function ensurePaperlessAllDept_() {
  ensureSheet_(SHEET_PL_DEPTS, ['DepartmentID', 'DepartmentName', 'PinHash', 'IsActive', 'SortOrder', 'CreatedAt', 'UpdatedAt']);
  if (sheetToObjects_(SHEET_PL_DEPTS).some(d => d.DepartmentID === PL_ALL_ID)) return;
  const now = new Date();
  appendByHeaders_(SHEET_PL_DEPTS, { DepartmentID: PL_ALL_ID, DepartmentName: PL_ALL_NAME, PinHash: sha256Hex_(newId_('X') + Math.random()),
    IsActive: true, SortOrder: 9999, CreatedAt: now, UpdatedAt: now });
}
/** รายการหน่วยงาน Paperless ที่แสดง: โหมดปิด = เฉพาะ "ภาพรวมทั้งโรงพยาบาล"; โหมดเปิด = หน่วยงานจริง (+ภาพรวมถ้ามียอดบันทึกไว้) */
function plVisibleDepts_() {
  const useDepts = isDeptMode_('paperless');
  if (!useDepts) ensurePaperlessAllDept_();
  const logs = sheetToObjects_(SHEET_PL_LOG);
  const allHasLogs = logs.some(l => l.DepartmentID === PL_ALL_ID);
  return sheetToObjects_(SHEET_PL_DEPTS)
    .filter(d => d.IsActive !== false)
    .filter(d => useDepts ? (d.DepartmentID !== PL_ALL_ID || allHasLogs) : d.DepartmentID === PL_ALL_ID)
    .sort((a, b) => (Number(a.SortOrder) || 0) - (Number(b.SortOrder) || 0));
}

// ============================================================================================
// ปฏิทินเก่าเราขอ: ลิงก์/QR ให้ผู้บริจาคแจ้งส่งปฏิทินเข้าโครงการ → ผู้ดูแลตรวจสอบ ชั่งน้ำหนัก แล้วรับเข้าเป็นรายการบริจาค
// ============================================================================================
const SHEET_CAL_SUBS = 'CalendarSubmissions';
const CAL_SUB_HEADERS = ['SubmissionID', 'SubmitDate', 'DonorType', 'DonorName', 'Organization', 'Contact', 'EstimatedQty', 'PlannedDate',
                         'Note', 'Status', 'LinkedDonationID', 'ReviewedBy', 'ReviewedAt', 'CreatedAt'];
const CAL_SUB_PENDING = 'รอตรวจสอบ';
const CAL_SUB_ACCEPTED = 'รับแล้ว';
const CAL_SUB_REJECTED = 'ไม่รับ';

function getCalendarFormLink() {
  return (ScriptApp.getService().getUrl() || '') + '?page=calendar&action=form';
}

/** สาธารณะ (ไม่ต้องมี PIN): แจ้งความประสงค์ส่งปฏิทินเก่าเข้าโครงการ */
function submitCalendarSubmission(data) {
  data = data || {};
  if (String(data.website || '').trim()) throw new Error('ไม่สามารถส่งได้');
  const type = String(data.DonorType || '').trim();
  if (CAL_DONOR_TYPES.indexOf(type) === -1) throw new Error('กรุณาเลือกประเภทผู้บริจาค');
  if (type === 'หน่วยงาน' && !isDeptMode_('calendar')) throw new Error('ขณะนี้ไม่ได้รับแจ้งในนามหน่วยงาน กรุณาเลือกประเภทอื่น');
  const name = String(data.DonorName || '').trim().replace(/\s+/g, ' ').slice(0, 120);
  if (!name) throw new Error(type === 'หน่วยงาน' ? 'กรุณาระบุชื่อหน่วยงาน' : 'กรุณากรอกชื่อผู้บริจาค');
  const org = String(data.Organization || '').trim().slice(0, 150);
  if (type === 'บริษัท/องค์กรภายนอก' && !org && name.length < 2) throw new Error('กรุณาระบุชื่อบริษัท/องค์กร');
  const contact = String(data.Contact || '').trim().slice(0, 120);
  if (!contact) throw new Error('กรุณากรอกช่องทางติดต่อ (เบอร์โทรหรืออีเมล) เพื่อให้เจ้าหน้าที่ประสานงาน');
  if (!INQUIRY_EMAIL_RE.test(contact) && !/^[0-9+\-\s()#ต่อ.]{6,}$/.test(contact)) throw new Error('ช่องทางติดต่อต้องเป็นอีเมลหรือเบอร์โทรที่ถูกต้อง');
  const qty = Math.max(0, Math.min(100000, Number(data.EstimatedQty) || 0));
  let planned = '';
  if (data.PlannedDate) { const d = new Date(data.PlannedDate); if (!isNaN(d.getTime())) planned = d; }
  try {
    if (typeof CacheService !== 'undefined') {
      const cache = CacheService.getScriptCache(), ck = 'cal:' + contact.toLowerCase() + name.toLowerCase();
      if (cache.get(ck)) throw new Error('คุณเพิ่งส่งรายการนี้ไป กรุณารอสักครู่');
      cache.put(ck, '1', 60);
    }
  } catch (e) { if (/เพิ่งส่ง/.test(String(e.message))) throw e; }
  ensureSheet_(SHEET_CAL_SUBS, CAL_SUB_HEADERS);
  const recent = readRowsIfSheet_(SHEET_CAL_SUBS).filter(r => r.CreatedAt && (Date.now() - new Date(r.CreatedAt).getTime()) < 3600 * 1000).length;
  if (recent >= 60) throw new Error('มีการส่งจำนวนมากในขณะนี้ กรุณาลองใหม่ภายหลัง');
  const id = newId_('CSB'), now = new Date();
  appendByHeaders_(SHEET_CAL_SUBS, { SubmissionID: id, SubmitDate: now, DonorType: type, DonorName: name, Organization: org, Contact: contact,
    EstimatedQty: qty, PlannedDate: planned, Note: String(data.Note || '').trim().slice(0, 300), Status: CAL_SUB_PENDING, CreatedAt: now });
  try {
    const to = getInquiryRecipients_();
    if (to.length) MailApp.sendEmail({ to: to.join(','), subject: '[ปฏิทินเก่าเราขอ] มีผู้แจ้งส่งปฏิทินเข้าโครงการ — ' + name,
      body: 'ผู้แจ้ง: ' + name + ' (' + type + ')\nหน่วยงาน/บริษัท: ' + (org || '-') + '\nติดต่อ: ' + contact + '\nจำนวนโดยประมาณ: ' + qty + ' เล่ม\nรหัส: ' + id +
        '\n\nตรวจสอบและรับเข้าได้ที่แท็บ "ปฏิทินเก่าเราขอ" ในระบบ (ผู้ดูแลระบบ)' });
  } catch (e) { Logger.log('ส่งอีเมลแจ้งไม่สำเร็จ: ' + e); }
  return { ok: true, id: id };
}

/** ผู้ดูแลระบบ: รายการที่แจ้งส่งปฏิทิน (ใหม่สุดก่อน) */
function getCalendarSubmissions(departmentId, pin) {
  assertAuthorized_(departmentId, pin);
  return readRowsIfSheet_(SHEET_CAL_SUBS).map(r => ({
    SubmissionID: r.SubmissionID, SubmitDate: toIsoDate_(r.SubmitDate), DonorType: r.DonorType, DonorName: r.DonorName, Organization: r.Organization || '',
    Contact: r.Contact, EstimatedQty: numOr0_(r.EstimatedQty), PlannedDate: toIsoDate_(r.PlannedDate), Note: r.Note || '',
    Status: r.Status || CAL_SUB_PENDING, ReviewedBy: r.ReviewedBy || '', ReviewedAt: toIsoDate_(r.ReviewedAt)
  })).sort((a, b) => (a.Status === CAL_SUB_PENDING ? 0 : 1) - (b.Status === CAL_SUB_PENDING ? 0 : 1) || (b.SubmitDate > a.SubmitDate ? 1 : -1));
}

/** ผู้ดูแลระบบ: รับเข้า (ต้องระบุน้ำหนัก กก.) → สร้างรายการบริจาค / หรือไม่รับ */
function reviewCalendarSubmission(departmentId, pin, submissionId, decision, opts) {
  const auth = assertAuthorized_(departmentId, pin);
  opts = opts || {};
  const found = findRowById_(SHEET_CAL_SUBS, 'SubmissionID', submissionId);
  if (!found) throw new Error('ไม่พบรายการนี้');
  const g = h => found.values[found.headers.indexOf(h)];
  if ((g('Status') || CAL_SUB_PENDING) !== CAL_SUB_PENDING) throw new Error('รายการนี้ตรวจสอบแล้ว');
  const sh = getSheet_(SHEET_CAL_SUBS);
  const set = (h, v) => { const c = found.headers.indexOf(h); if (c !== -1) sh.getRange(found.rowIndex, c + 1).setValue(v); };
  let donationId = '';
  if (decision === 'accept') {
    const n = calendarNormalize_({ DonationDate: opts.DonationDate || new Date(), RoundName: opts.RoundName, DonorType: g('DonorType'),
      DonorName: g('DonorName'), WeightKg: opts.WeightKg, Note: opts.Note || (g('Organization') ? 'บริษัท/หน่วยงาน: ' + g('Organization') : '') });
    ensureSheet_(SHEET_CAL_DON, CAL_HEADERS);
    const now = new Date(); donationId = newId_('CAL');
    appendByHeaders_(SHEET_CAL_DON, { DonationID: donationId, DonationDate: n.date, RoundName: n.round, DonorType: n.type, DonorName: n.name, WeightKg: n.kg,
      Note: n.note, RecordedByName: auth.departmentName, CreatedAt: now, UpdatedAt: now });
    set('Status', CAL_SUB_ACCEPTED); set('LinkedDonationID', donationId);
    syncCalendarLinkedProject_();
  } else if (decision === 'reject') {
    set('Status', CAL_SUB_REJECTED);
  } else throw new Error('การตัดสินใจไม่ถูกต้อง');
  set('ReviewedBy', auth.departmentName); set('ReviewedAt', new Date());
  return { ok: true, donationId: donationId };
}

// ============================================================================================
// เลือกหน่วยงานที่ "ใช้" ในแต่ละโครงการ (ปิดหน่วยงานที่ไม่เกี่ยวข้อง) — ใช้คอลัมน์ IsActive ของชีตรายชื่อหน่วยงานของแต่ละโครงการ
// หน่วยงานที่ปิดจะไม่แสดงในฟอร์ม/ตาราง Paperless ข้อมูลที่บันทึกไว้เดิมไม่ถูกลบ เปิดกลับมาได้ทุกเมื่อ
// ============================================================================================
function projectDeptSheet_(key) {
  const map = { paperless: SHEET_PL_DEPTS, cartridge: SHEET_CTG_DEPTS, share: SHEET_SHARE_DEPTS };
  if (!map[key]) throw new Error('โครงการนี้ไม่มีรายชื่อหน่วยงานให้เลือก');
  return map[key];
}

/** ผู้ดูแลระบบ: รายชื่อหน่วยงานทั้งหมดของโครงการ พร้อมสถานะใช้งาน */
function getProjectDepartments(departmentId, pin, projectKey) {
  assertAuthorized_(departmentId, pin);
  const sheet = projectDeptSheet_(projectKey);
  return sheetToObjects_(sheet)
    .filter(d => d.DepartmentID !== PL_ALL_ID && d.DepartmentName)
    .sort((a, b) => (Number(a.SortOrder) || 0) - (Number(b.SortOrder) || 0))
    .map(d => ({ DepartmentID: d.DepartmentID, DepartmentName: d.DepartmentName, IsActive: shareActive_(d.IsActive) }));
}

/** ผู้ดูแลระบบ: เปิด/ปิดหน่วยงานหลายรายการพร้อมกัน — changes = [{DepartmentID, IsActive}] */
function setProjectDepartmentsActive(departmentId, pin, projectKey, changes) {
  assertAuthorized_(departmentId, pin);
  const sheet = projectDeptSheet_(projectKey);
  const want = {};
  (changes || []).forEach(c => { if (c && c.DepartmentID) want[c.DepartmentID] = (c.IsActive === true || c.IsActive === 'true'); });
  const sh = getSheet_(sheet);
  const values = sh.getDataRange().getValues();
  const headers = values[0], idCol = headers.indexOf('DepartmentID'), actCol = headers.indexOf('IsActive');
  if (idCol === -1 || actCol === -1) throw new Error('ชีตรายชื่อหน่วยงานไม่มีคอลัมน์ที่ต้องใช้ (DepartmentID/IsActive)');
  let n = 0;
  for (let r = 1; r < values.length; r++) {
    const id = values[r][idCol];
    if (Object.prototype.hasOwnProperty.call(want, id) && shareActive_(values[r][actCol]) !== want[id]) {
      sh.getRange(r + 1, actCol + 1).setValue(want[id]);
      n++;
    }
  }
  return { ok: true, changed: n };
}

/**
 * ผู้ดูแลระบบ: เพิ่มหน่วยงานใหม่จากหน้าเว็บ (ไม่ต้องเปิดแก้ชีตเอง)
 *  - projectKey : 'paperless' | 'cartridge' | 'share' (โครงการที่เปิดหน้าต่างอยู่)
 *  - alsoKeys   : โครงการอื่นที่ต้องการเพิ่มชื่อเดียวกันด้วย (ไม่บังคับ)
 *  - deptPin    : PIN ของหน่วยงาน ใช้เฉพาะ Paperless (หน่วยงานต้องเข้าสู่ระบบเพื่อบันทึกยอดเอง) — เก็บเป็นแฮช SHA-256
 * ถ้ามีชื่อนี้อยู่แล้วแต่ถูกปิดไว้ → เปิดใช้งานกลับ; ถ้าเปิดอยู่แล้ว → ข้าม (ไม่สร้างซ้ำ)
 */
function addProjectDepartment(departmentId, pin, projectKey, name, deptPin, alsoKeys) {
  assertAuthorized_(departmentId, pin);
  const clean = String(name || '').replace(/\s+/g, ' ').trim();
  if (!clean) throw new Error('กรุณากรอกชื่อหน่วยงาน');
  if (clean.length > 120) throw new Error('ชื่อหน่วยงานยาวเกินไป (ไม่เกิน 120 ตัวอักษร)');
  const keys = [projectKey].concat(alsoKeys || []).filter((k, i, a) => k && a.indexOf(k) === i);
  keys.forEach(projectDeptSheet_); // ตรวจว่าเป็นโครงการที่รองรับ
  const pinText = String(deptPin || '').trim();
  if (keys.indexOf('paperless') !== -1 && !/^\d{4,8}$/.test(pinText)) {
    throw new Error('หน่วยงาน Paperless ต้องตั้ง PIN เป็นตัวเลข 4–8 หลัก');
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const result = {};
    keys.forEach(key => {
      const sheet = projectDeptSheet_(key);
      const sh = getSheet_(sheet);
      const values = sh.getDataRange().getValues();
      const headers = values[0];
      const nameCol = headers.indexOf('DepartmentName'), actCol = headers.indexOf('IsActive'), sortCol = headers.indexOf('SortOrder');
      const updCol = headers.indexOf('UpdatedAt');
      const norm = s => String(s || '').replace(/\s+/g, ' ').trim().toLowerCase();
      let maxSort = 0, foundRow = -1;
      for (let r = 1; r < values.length; r++) {
        if (values[r][headers.indexOf('DepartmentID')] === PL_ALL_ID) continue;
        maxSort = Math.max(maxSort, Number(values[r][sortCol]) || 0);
        if (norm(values[r][nameCol]) === norm(clean)) foundRow = r;
      }
      if (foundRow !== -1) {
        if (shareActive_(values[foundRow][actCol])) { result[key] = 'exists'; return; }
        sh.getRange(foundRow + 1, actCol + 1).setValue(true);
        if (updCol !== -1) sh.getRange(foundRow + 1, updCol + 1).setValue(new Date());
        result[key] = 'reactivated';
        return;
      }
      const now = new Date();
      const prefix = { paperless: 'PLD', cartridge: 'CTGD', share: 'SHD' }[key];
      appendByHeaders_(sheet, {
        DepartmentID: newId_(prefix), DepartmentName: clean, SortOrder: maxSort + 1, IsActive: true,
        PinHash: key === 'paperless' ? sha256Hex_(pinText) : '', CreatedAt: now, UpdatedAt: now
      });
      result[key] = 'added';
    });
    return { ok: true, name: clean, result: result };
  } finally {
    lock.releaseLock();
  }
}

// ============================================================================================
// รายงานรายเดือนแยกหน่วยงาน (เมทริกซ์ หน่วยงาน × 12 เดือน) — Paperless / ตลับหมึก / คัดแยก แบ่งปัน
// เปิดดูได้ทุกคน (ไม่มีข้อมูลส่วนบุคคล) ใช้ทั้งตาราง กราฟ และรายงาน PDF
// ============================================================================================
function getDeptMonthlyMatrix(projectKey, year) {
  const cfg = {
    paperless: { unit: 'แผ่น', label: 'กระดาษที่ลดลง', title: 'Paperless' },
    cartridge: { unit: 'ตลับ', label: 'จำนวนตลับหมึกที่ส่งเข้าโครงการ', title: 'ตลับหมึกเก่าเราขอ' },
    share:     { unit: 'บาท', label: 'มูลค่าที่รับแบ่งปัน', title: 'คัดแยก แบ่งปัน' }
  }[projectKey];
  if (!cfg) throw new Error('โครงการนี้ไม่มีรายงานแยกหน่วยงาน');

  const entries = []; // {name, year, month(0-11), value, count}
  let inactive = {};
  if (projectKey === 'paperless') {
    const visible = {};
    plVisibleDepts_().forEach(d => { visible[d.DepartmentID] = d.DepartmentName; });
    sheetToObjects_(SHEET_PL_LOG).forEach(l => {
      const nm = visible[l.DepartmentID];
      if (!nm) return;
      entries.push({ name: nm, year: Number(l.Year), month: Number(l.Month) - 1, value: Number(l.SheetsReduced) || 0, count: Number(l.ItemCount) || 0 });
    });
  } else {
    const sheet = projectKey === 'cartridge' ? SHEET_CTG_DEPTS : SHEET_SHARE_DEPTS;
    try { sheetToObjects_(sheet).forEach(d => { if (d.DepartmentName && !shareActive_(d.IsActive)) inactive[d.DepartmentName] = true; }); } catch (e) {}
    if (projectKey === 'cartridge') {
      sheetToObjects_(SHEET_CTG_CONTRIB).forEach(c => {
        if (!isVerified_(c.VerifyStatus)) return;
        if (!c.ContributionDate || !c.DepartmentName || inactive[c.DepartmentName]) return;
        const d = new Date(c.ContributionDate);
        if (isNaN(d)) return;
        entries.push({ name: c.DepartmentName, year: d.getFullYear(), month: d.getMonth(), value: Number(c.Quantity) || 0, count: 1 });
      });
    } else {
      shareRows_(SHEET_SHARE_WD).forEach(w => {
        if (!w.WithdrawDate) return;
        const nm = w.DepartmentName || '(ไม่ระบุหน่วยงาน)';
        if (inactive[nm]) return;
        const d = new Date(w.WithdrawDate);
        if (isNaN(d)) return;
        const price = shareHasNum_(w.UnitPrice) ? Number(w.UnitPrice) : 0;
        entries.push({ name: nm, year: d.getFullYear(), month: d.getMonth(), value: shareNum_(w.Quantity) * price, count: 1 });
      });
    }
  }

  const years = Array.from(new Set(entries.map(e => e.year).filter(y => y > 2000))).sort((a, b) => b - a);
  const targetYear = Number(year) || (years[0] || new Date().getFullYear());
  const map = {};
  entries.filter(e => e.year === targetYear && e.month >= 0 && e.month < 12).forEach(e => {
    const r = map[e.name] = map[e.name] || { name: e.name, monthly: new Array(12).fill(0), counts: new Array(12).fill(0), total: 0, count: 0 };
    r.monthly[e.month] += e.value; r.counts[e.month] += e.count; r.total += e.value; r.count += e.count;
  });
  const rows = Object.keys(map).map(k => {
    const r = map[k];
    r.monthly = r.monthly.map(round2_); r.total = round2_(r.total);
    return r;
  }).filter(r => r.total > 0 || r.count > 0).sort((a, b) => b.total - a.total || a.name.localeCompare(b.name, 'th'));
  rows.forEach((r, i) => { r.rank = i + 1; });
  const monthTotals = new Array(12).fill(0), monthCounts = new Array(12).fill(0);
  rows.forEach(r => r.monthly.forEach((v, m) => { monthTotals[m] += v; monthCounts[m] += r.counts[m]; }));
  return {
    projectKey: projectKey, title: cfg.title, unit: cfg.unit, label: cfg.label,
    year: targetYear, years: years.length ? years : [targetYear],
    monthNames: PL_MONTH_NAMES, rows: rows,
    monthTotals: monthTotals.map(round2_), monthCounts: monthCounts,
    grandTotal: round2_(monthTotals.reduce((s, v) => s + v, 0))
  };
}


// ============================================================================================
// ยืนยันรายการที่ส่งเข้าโครงการ (ทอดไม่ทิ้ง / ตลับหมึกเก่าเราขอ)
//  - รายการที่บันทึกผ่านฟอร์มสาธารณะจะเป็น "รอยืนยัน" และยังไม่นับในยอด/อันดับ/แต้ม จนกว่าผู้ดูแลระบบกดยืนยัน
//  - รายการเดิมก่อนอัปเดตนี้ (คอลัมน์ VerifyStatus ว่าง) ถือว่ายืนยันแล้ว ยอดเดิมไม่เปลี่ยน
// ============================================================================================
var VERIFY_PENDING = 'รอยืนยัน';
var VERIFY_OK = 'ยืนยันแล้ว';
var VERIFY_REJECTED = 'ไม่อนุมัติ';
var VERIFY_COLS = ['VerifyStatus', 'VerifiedBy', 'VerifiedAt'];

function isVerified_(v) {
  const s = String(v || '').trim();
  return s === '' || s === VERIFY_OK;
}

var VERIFY_COLS_CHECKED_ = {};
function ensureVerifyColumns_(sheetName) {
  if (VERIFY_COLS_CHECKED_[sheetName]) return;
  VERIFY_COLS_CHECKED_[sheetName] = true;
  const sh = findSheet_(sheetName);
  if (!sh) return;
  const lastCol = sh.getLastColumn();
  const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  const need = VERIFY_COLS.filter(h => headers.indexOf(h) === -1);
  need.forEach((h, i) => sh.getRange(1, lastCol + i + 1).setValue(h).setFontWeight('bold'));
  if (need.length) invalidateReadCache_();
}

function verifyCfg_(key) {
  if (key === 'ftf') return { sheet: SHEET_FTF_LOG, id: 'LogID', sync: syncFtfLinkedProject_ };
  if (key === 'cartridge') return { sheet: SHEET_CTG_CONTRIB, id: 'ContributionID', sync: syncCartridgeLinkedProject_ };
  if (key === 'emuse') return { sheet: SHEET_EM_USAGE, id: 'UsageID', sync: syncEmLinkedProject_ };
  if (key === 'emlog') return { sheet: SHEET_EM_LOGS, id: 'LogID', sync: syncEmLinkedProject_ };
  if (key === 'cpuse') return { sheet: SHEET_CP_USAGE, id: 'UsageID', sync: syncCpLinkedProject_ };
  if (key === 'cplog') return { sheet: SHEET_CP_LOGS, id: 'LogID', sync: syncCpLinkedProject_ };
  throw new Error('โครงการนี้ไม่มีขั้นตอนยืนยันรายการ');
}

/** ผู้ดูแลระบบ: รายการที่รอยืนยัน */
function getPendingVerifications(departmentId, pin, key) {
  assertAuthorized_(departmentId, pin);
  const cfg = verifyCfg_(key);
  ensureVerifyColumns_(cfg.sheet);
  return sheetToObjects_(cfg.sheet)
    .filter(r => r[cfg.id] && String(r.VerifyStatus || '') === VERIFY_PENDING)
    .map(r => key === 'cplog' ? cpLogPendingRow_(r)
      : key === 'cpuse' ? { id: r.UsageID, at: r.CreatedAt ? toIsoDateTime_(r.CreatedAt) : '', date: toIsoDate_(r.UsageDate), name: r.UsedBy || '', place: r.Place || '', qty: Number(r.Kg) || 0, purpose: r.Purpose || '', note: r.Note || '' }
      : key === 'emlog' ? emLogPendingRow_(r)
      : key === 'emuse'
      ? { id: r.UsageID, at: r.CreatedAt ? toIsoDateTime_(r.CreatedAt) : '', date: toIsoDate_(r.UsageDate), name: r.UsedBy || '', place: r.Place || '', qty: Number(r.Liters) || 0, purpose: r.Purpose || '', note: r.Note || '' }
      : key === 'ftf'
      ? { id: r.LogID, at: r.Timestamp ? toIsoDateTime_(r.Timestamp) : '', staffCode: String(r.StaffCode || ''), name: r.StaffName || '',
          phone: r.Phone || '', action: r.ActionType || '', qty: Number(r.Quantity) || 0, note: r.Note || '' }
      : { id: r.ContributionID, at: r.CreatedAt ? toIsoDateTime_(r.CreatedAt) : '', date: toIsoDate_(r.ContributionDate),
          name: r.DepartmentName || '', qty: Number(r.Quantity) || 0, note: r.Note || '' })
    .sort((a, b) => String(a.at).localeCompare(String(b.at)));
}

/** ผู้ดูแลระบบ: ยืนยัน (approve=true) หรือไม่อนุมัติ (approve=false) หลายรายการพร้อมกัน */
function verifySubmissions(departmentId, pin, key, ids, approve) {
  const auth = assertAuthorized_(departmentId, pin);
  const cfg = verifyCfg_(key);
  ensureVerifyColumns_(cfg.sheet);
  const want = {};
  (ids || []).forEach(id => { want[id] = true; });
  if (key === 'cplog' && approve !== false) return verifyCpLogs_(auth, ids);
  if (key === 'cpuse' && approve !== false) {
    const wantC = readRowsIfSheet_(SHEET_CP_USAGE).filter(u => want[u.UsageID] && String(u.VerifyStatus || '') === VERIFY_PENDING).reduce((s, u) => s + (Number(u.Kg) || 0), 0);
    const stockC = cpStockKg_();
    if (wantC > stockC + 1e-9) throw new Error('ปุ๋ยหมักคงเหลือไม่พอสำหรับรายการที่ยืนยัน (คงเหลือ ' + stockC + ' กก. แต่รายการรวม ' + round2_(wantC) + ' กก.) — ยืนยันทีละรายการหรือบันทึกผลผลิตเพิ่มก่อน');
  }
  if (key === 'emlog' && approve !== false) return verifyEmLogs_(auth, ids);
  if (key === 'emuse' && approve !== false) {
    const want2 = readRowsIfSheet_(SHEET_EM_USAGE).filter(u => want[u.UsageID] && String(u.VerifyStatus || '') === VERIFY_PENDING).reduce((s, u) => s + (Number(u.Liters) || 0), 0);
    const stock = emStockLiters_();
    if (want2 > stock + 1e-9) throw new Error('น้ำยา EM คงเหลือไม่พอสำหรับรายการที่ยืนยัน (คงเหลือ ' + stock + ' ลิตร แต่รายการรวม ' + round2_(want2) + ' ลิตร) — ยืนยันทีละรายการหรือบันทึกผลผลิตเพิ่มก่อน');
  }
  const sh = getSheet_(cfg.sheet);
  const values = sh.getDataRange().getValues();
  const h = values[0];
  const idCol = h.indexOf(cfg.id), stCol = h.indexOf('VerifyStatus'), byCol = h.indexOf('VerifiedBy'), atCol = h.indexOf('VerifiedAt');
  const now = new Date();
  let n = 0;
  for (let r = 1; r < values.length; r++) {
    if (!want[values[r][idCol]] || String(values[r][stCol] || '') !== VERIFY_PENDING) continue;
    sh.getRange(r + 1, stCol + 1).setValue(approve === false ? VERIFY_REJECTED : VERIFY_OK);
    if (byCol !== -1) sh.getRange(r + 1, byCol + 1).setValue(auth.departmentName);
    if (atCol !== -1) sh.getRange(r + 1, atCol + 1).setValue(now);
    n++;
  }
  invalidateReadCache_();
  if (n) { try { cfg.sync(); } catch (e) {} }
  return { ok: true, changed: n };
}

function toIsoDateTime_(v) {
  const d = v instanceof Date ? v : new Date(v);
  if (isNaN(d)) return '';
  return Utilities.formatDate(d, Session.getScriptTimeZone(), "yyyy-MM-dd'T'HH:mm:ss");
}


// ============================================================================================
// เป้าหมายรายเดือนของแต่ละโครงการ → ร้อยละความสำเร็จ (แสดงที่หน้าภาพรวม + แท็บโครงการ)
//  - ผู้ดูแลระบบตั้งเป้าหมายรายเดือนรายปี (ชีต ProjectTargets) ผลจริงดึงจากข้อมูลที่บันทึกในแต่ละโครงการอัตโนมัติ (เฉพาะรายการที่ยืนยันแล้ว)
//  - % ตามเป้าปี = ผลจริงสะสมปีนี้ ÷ เป้าหมายรวมทั้งปี  |  % เทียบเป้าถึงเดือนนี้ = ผลจริงสะสม ÷ เป้าสะสมถึงเดือนปัจจุบัน (บอกว่าเร็วหรือช้ากว่าแผน)
//  - โครงการที่ตั้งเป้าไว้และอยู่ในโหมดสถานะอัตโนมัติ จะใช้ % ตามเป้าปีเป็น "ความคืบหน้า" ของโครงการ
// ============================================================================================
const SHEET_PROJECT_TARGETS = 'ProjectTargets';
const TARGET_KEYS = ['ftf', 'paperless', 'cartridge', 'kitchen', 'calendar', 'share', 'em'];
const TARGET_METRICS = {
  ftf:       { col: 'gallons',    label: 'น้ำมันที่รับบริจาค',        unit: 'แกลลอน', dec: 1 },
  paperless: { col: 'sheets',     label: 'กระดาษที่ลดได้',           unit: 'แผ่น',   dec: 0 },
  cartridge: { col: 'contributed', label: 'ตลับหมึกที่ส่งเข้าโครงการ', unit: 'ตลับ',   dec: 0 },
  kitchen:   { col: 'liters',     label: 'น้ำมันที่ส่งออก',           unit: 'ลิตร',   dec: 1 },
  calendar:  { col: 'kg',         label: 'ปฏิทินที่รับบริจาค',        unit: 'กก.',    dec: 1 },
  share:     { col: 'value',      label: 'มูลค่าที่แบ่งปัน (ประหยัดงบ)', unit: 'บาท',   dec: 0 },
  em:        { col: 'emLiters',   label: 'น้ำยา EM ที่ผลิตได้',       unit: 'ลิตร',   dec: 1 },
  compost:   { col: 'cpKg',       label: 'ปุ๋ยหมักที่ผลิตได้',         unit: 'กก.',    dec: 1 }
};
const TARGET_PROJECT_ID = {
  ftf: LINKED_PROJECT_FTF_ID, paperless: LINKED_PROJECT_PL_ID, cartridge: LINKED_PROJECT_CTG_ID,
  kitchen: LINKED_PROJECT_KIT_ID, calendar: LINKED_PROJECT_CAL_ID, share: LINKED_PROJECT_SHARE_ID, em: LINKED_PROJECT_EM_ID, compost: LINKED_PROJECT_CP_ID
};

/** ProjectID ของแถวโครงการเชื่อมโยง → คีย์เป้าหมาย (ftf/paperless/…) */
function targetKeyOf_(projectId) {
  return TARGET_KEYS.filter(k => TARGET_PROJECT_ID[k] === projectId)[0] || '';
}

function targetRows_(key, year) {
  return readRowsIfSheet_(SHEET_PROJECT_TARGETS).filter(r => r.ProjectKey === key && Number(r.Year) === Number(year));
}

/** เดือนปัจจุบันของปีที่ดู: ปีนี้ = เดือนปัจจุบัน, ปีที่ผ่านมา = 12, ปีอนาคต = 0 */
function targetCurMonth_(year) {
  const now = new Date();
  if (year < now.getFullYear()) return 12;
  if (year > now.getFullYear()) return 0;
  return now.getMonth() + 1;
}

/** เป้าหมาย + ผลจริงรายเดือน + ร้อยละ ของโครงการหนึ่งในปีที่ระบุ */
function getProjectTargets(projectKey, year) {
  const metric = TARGET_METRICS[projectKey];
  if (!metric) throw new Error('โครงการนี้ยังไม่รองรับการตั้งเป้าหมาย');
  const y = Number(year) || new Date().getFullYear();
  const tmap = {};
  targetRows_(projectKey, y).forEach(r => { tmap[Number(r.Month)] = Number(r.TargetValue) || 0; });
  const ps = getPeriodSummary(projectKey, y);
  const cur = targetCurMonth_(y);
  let cumT = 0, cumA = 0;
  const months = ps.monthly.map((m, i) => {
    const target = tmap[i + 1] || 0, actual = Number(m[metric.col]) || 0;
    cumT += target; cumA += actual;
    return { month: i + 1, target: round2_(target), actual: round2_(actual), pct: target > 0 ? round2_(actual / target * 100) : null,
             cumTarget: round2_(cumT), cumActual: round2_(cumA), cumPct: cumT > 0 ? round2_(cumA / cumT * 100) : null };
  });
  const annualTarget = cumT;
  const ytd = months.slice(0, cur);
  const ytdTarget = ytd.reduce((s, m) => s + m.target, 0), ytdActual = ytd.reduce((s, m) => s + m.actual, 0);
  const yearActual = months.reduce((s, m) => s + m.actual, 0);
  const allYears = Array.from(new Set(readRowsIfSheet_(SHEET_PROJECT_TARGETS).filter(r => r.ProjectKey === projectKey).map(r => Number(r.Year)))).sort((a, b) => b - a);
  return {
    key: projectKey, year: y, metric: metric, curMonth: cur, months: months, targetYears: allYears, years: ps.years,
    hasTarget: annualTarget > 0,
    annualTarget: round2_(annualTarget), yearActual: round2_(yearActual),
    annualPct: annualTarget > 0 ? round2_(yearActual / annualTarget * 100) : null,
    ytdTarget: round2_(ytdTarget), ytdActual: round2_(ytdActual),
    ytdPct: ytdTarget > 0 ? round2_(ytdActual / ytdTarget * 100) : null,
    monthsWithTarget: months.filter(m => m.target > 0).length
  };
}

/** ฉบับย่อสำหรับหน้าภาพรวม (ปีปัจจุบัน) — คืน null ถ้ายังไม่ตั้งเป้าหมาย */
function targetProgressFor_(projectKey, year) {
  if (!TARGET_METRICS[projectKey]) return null;
  const y = Number(year) || new Date().getFullYear();
  if (!targetRows_(projectKey, y).some(r => Number(r.TargetValue) > 0)) return null;
  const t = getProjectTargets(projectKey, y);
  return { year: t.year, hasTarget: t.hasTarget, annualPct: t.annualPct, ytdPct: t.ytdPct, annualTarget: t.annualTarget, yearActual: t.yearActual,
           ytdTarget: t.ytdTarget, ytdActual: t.ytdActual, unit: t.metric.unit, label: t.metric.label, dec: t.metric.dec, curMonth: t.curMonth };
}

/** ผู้ดูแลระบบ: บันทึกเป้าหมาย 12 เดือนของโครงการ/ปี  targets = [เป้าเดือน ม.ค. … ธ.ค.] (ว่าง/0 = ไม่ตั้งเป้าเดือนนั้น) */
function setProjectTargets(departmentId, pin, projectKey, year, targets) {
  const auth = assertAuthorized_(departmentId, pin);
  if (!TARGET_METRICS[projectKey]) throw new Error('โครงการนี้ยังไม่รองรับการตั้งเป้าหมาย');
  const y = Number(year);
  if (!y || y < 2000 || y > 2200) throw new Error('ปีไม่ถูกต้อง');
  if (!Array.isArray(targets) || targets.length !== 12) throw new Error('ต้องระบุเป้าหมายครบ 12 เดือน');
  const vals = targets.map(v => {
    const n = Number(v);
    if (v !== '' && v !== null && v !== undefined && (isNaN(n) || n < 0)) throw new Error('เป้าหมายต้องเป็นตัวเลขที่ไม่ติดลบ');
    return isNaN(n) ? 0 : n;
  });
  const headers = ['TargetKey', 'ProjectKey', 'Year', 'Month', 'TargetValue', 'UpdatedBy', 'UpdatedAt'];
  const sh = ensureSheet_(SHEET_PROJECT_TARGETS, headers);
  const data = sh.getDataRange().getValues();
  const h = data[0];
  const col = n => h.indexOf(n);
  const now = new Date();
  const existing = {};
  for (let r = 1; r < data.length; r++) existing[data[r][col('TargetKey')]] = r + 1;
  vals.forEach((v, i) => {
    const k = projectKey + '|' + y + '|' + (i + 1);
    if (existing[k]) {
      sh.getRange(existing[k], col('TargetValue') + 1).setValue(v);
      sh.getRange(existing[k], col('UpdatedBy') + 1).setValue(auth.departmentName);
      sh.getRange(existing[k], col('UpdatedAt') + 1).setValue(now);
    } else if (v > 0) {
      const obj = { TargetKey: k, ProjectKey: projectKey, Year: y, Month: i + 1, TargetValue: v, UpdatedBy: auth.departmentName, UpdatedAt: now };
      sh.appendRow(h.map(x => (x in obj ? obj[x] : '')));
    }
  });
  invalidateReadCache_();
  // อัปเดตความคืบหน้าของโครงการเชื่อมโยง (โหมดอัตโนมัติ) ให้ตรงกับเป้าหมายใหม่
  try {
    const f = { ftf: syncFtfLinkedProject_, paperless: syncPaperlessLinkedProject_, cartridge: syncCartridgeLinkedProject_,
                share: syncShareLinkedProject_, kitchen: syncKitchenLinkedProject_, calendar: syncCalendarLinkedProject_ }[projectKey];
    if (f) f();
  } catch (e) { /* ข้าม */ }
  return getProjectTargets(projectKey, y);
}

// ============================================================================================
// โครงการ "น้ำยา EM จากเศษอาหาร"
//  INPUT   : เศษอาหาร / ผักและผลไม้ / เนื้อสัตว์ (กก.) + กากน้ำตาล/น้ำ
//  PROCESS : หมักเป็นรอบ (หลายถังต่อรอบ) ติดตามอายุการหมัก 7 / 15 / 45 วัน และตรวจทุก 7 วัน
//  OUTPUT  : เมื่อครบกำหนด (ReadyDays) เก็บเป็นน้ำยา EM กี่ลิตร
//  USE     : บันทึกการนำไปใช้ (สถานที่ ผู้ใช้ ปริมาณ) → คงเหลือ = ผลิตได้ − ใช้ไป
//  ชีต: EMBatches (รอบการหมัก) / EMChecks (ผลตรวจทุก 7 วัน) / EMUsage (การนำไปใช้) / EMSettings (ค่าตั้งต้น ปรับได้ในชีตโดยไม่ต้องแก้โค้ด)
//  การบันทึก/แก้ไข: ผู้ดูแลระบบ (IsAdmin) เท่านั้น • การดูรายงาน: ทุกคน
// ============================================================================================
const SHEET_EM_BATCHES  = 'EMBatches';
const SHEET_EM_CHECKS   = 'EMChecks';
const SHEET_EM_USAGE    = 'EMUsage';
const SHEET_EM_SETTINGS = 'EMSettings';
const SHEET_EM_LOGS     = 'EMDailyLogs';
const EM_LOG_HEADERS = ['LogID', 'LogDate', 'BatchID', 'LogType', 'FoodKg', 'FruitVegKg', 'MeatKg', 'Barrels', 'Location', 'PH', 'Smell', 'Appearance', 'Action', 'Note', 'Operator', 'CreatedAt', 'VerifyStatus', 'VerifiedBy', 'VerifiedAt'];
const EM_LOG_NEW = 'เริ่มหมักถังใหม่', EM_LOG_ADD = 'เติมวัตถุดิบเพิ่ม', EM_LOG_CHECK = 'ตรวจ/ดูแลประจำวัน';
const EM_LOG_TYPES = [EM_LOG_NEW, EM_LOG_ADD, EM_LOG_CHECK];
const EM_BATCH_HEADERS = ['BatchID', 'BatchNo', 'StartDate', 'FoodKg', 'FruitVegKg', 'MeatKg', 'MolassesKg', 'WaterL', 'BarrelCount', 'BarrelLiters',
                          'ExpectedLiters', 'ReadyDays', 'Location', 'Note', 'Status', 'HarvestDate', 'YieldLiters', 'HarvestNote', 'RecordedBy', 'CreatedAt', 'UpdatedAt'];
const EM_CHECK_HEADERS = ['CheckID', 'BatchID', 'CheckDate', 'Week', 'DayNo', 'PH', 'Smell', 'Appearance', 'Action', 'Note', 'RecordedBy', 'CreatedAt'];
const EM_USAGE_HEADERS = ['UsageID', 'BatchID', 'UsageDate', 'Liters', 'Place', 'UsedBy', 'Purpose', 'Note', 'RecordedBy', 'CreatedAt', 'VerifyStatus', 'VerifiedBy', 'VerifiedAt'];
const EM_STATUS_FERMENT = 'กำลังหมัก';
const EM_STATUS_DONE    = 'เก็บเป็นน้ำยา EM แล้ว';
const EM_STATUS_CANCEL  = 'ยกเลิก';
const EM_DEFAULTS = { Stage1Days: 7, Stage2Days: 15, ReadyDays: 45, CheckIntervalDays: 7, DefaultBarrelLiters: 200, YieldPercent: 60, SoonDays: 7 };

function ensureEmSheets_() {
  ensureSheet_(SHEET_EM_BATCHES, EM_BATCH_HEADERS);
  ensureSheet_(SHEET_EM_CHECKS, EM_CHECK_HEADERS);
  ensureSheet_(SHEET_EM_USAGE, EM_USAGE_HEADERS);
  ensureVerifyColumns_(SHEET_EM_USAGE);
  ensureSheet_(SHEET_EM_LOGS, EM_LOG_HEADERS);
  ensureEmSettingsSeeded_();
}

function ensureEmSettingsSeeded_() {
  const sh = ensureSheet_(SHEET_EM_SETTINGS, ['Key', 'Value']);
  const have = sheetToObjects_(SHEET_EM_SETTINGS).map(r => r.Key);
  Object.keys(EM_DEFAULTS).forEach(k => { if (have.indexOf(k) === -1) sh.appendRow([k, EM_DEFAULTS[k]]); });
}

/** ค่าตั้งต้นของโครงการ (ชีต EMSettings): ช่วงวันที่สำคัญของการหมัก ขนาดถัง และ % น้ำยาที่คาดว่าจะได้ */
function getEmSettings_() {
  const s = Object.assign({}, EM_DEFAULTS);
  readRowsIfSheet_(SHEET_EM_SETTINGS).forEach(r => {
    if (Object.prototype.hasOwnProperty.call(EM_DEFAULTS, r.Key) && r.Value !== '' && !isNaN(Number(r.Value)) && Number(r.Value) > 0) s[r.Key] = Number(r.Value);
  });
  if (s.Stage2Days <= s.Stage1Days) s.Stage2Days = s.Stage1Days + 1;
  if (s.ReadyDays <= s.Stage2Days) s.ReadyDays = s.Stage2Days + 1;
  return s;
}

/** วันที่ (ปฏิทินท้องถิ่น) → เลขวันนับจาก epoch (ตัดเวลาทิ้ง) */
function emDayNum_(v) {
  if (v === '' || v === null || v === undefined) return null;
  const s = (v instanceof Date || typeof v === 'number') ? toIsoDate_(v) : String(v);
  const m = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (!m) return null;
  return Math.floor(Date.UTC(Number(m[1]), Number(m[2]) - 1, Number(m[3])) / 86400000);
}
function emIso_(n) { return new Date(n * 86400000).toISOString().slice(0, 10); }
function emToday_() { return emDayNum_(new Date()); }
function emInputKg_(b) { return round2_((Number(b.FoodKg) || 0) + (Number(b.FruitVegKg) || 0) + (Number(b.MeatKg) || 0)); }

/** ข้อมูลรอบการหมัก 1 รอบ ณ วันที่ asOfN: อายุ ระยะ ความคืบหน้า ผลตรวจรายสัปดาห์ */
function emLogView_(l) {
  const kg = (Number(l.FoodKg) || 0) + (Number(l.FruitVegKg) || 0) + (Number(l.MeatKg) || 0);
  return { LogID: l.LogID, LogDate: toIsoDate_(l.LogDate), BatchID: l.BatchID || '', LogType: l.LogType || '', Operator: l.Operator || '', Kg: round2_(kg), FoodKg: Number(l.FoodKg) || 0, FruitVegKg: Number(l.FruitVegKg) || 0, MeatKg: Number(l.MeatKg) || 0,
           Barrels: Number(l.Barrels) || 0, Location: l.Location || '', PH: l.PH === '' || l.PH === null || l.PH === undefined ? null : Number(l.PH), Smell: l.Smell || '', Appearance: l.Appearance || '', Action: l.Action || '', Note: l.Note || '' };
}
function emBatchView_(b, st, asOfN, checks, usedL) {
  const startN = emDayNum_(b.StartDate);
  const ready = Number(b.ReadyDays) > 0 ? Number(b.ReadyDays) : st.ReadyDays;
  const barrels = Number(b.BarrelCount) || 0;
  const barrelL = Number(b.BarrelLiters) > 0 ? Number(b.BarrelLiters) : st.DefaultBarrelLiters;
  const status = b.Status || EM_STATUS_FERMENT;
  const harvestN = emDayNum_(b.HarvestDate);
  const harvested = status === EM_STATUS_DONE && (harvestN === null || harvestN <= asOfN);
  const cancelled = status === EM_STATUS_CANCEL;
  const age = startN === null ? 0 : Math.max(0, asOfN - startN);
  const expected = Number(b.ExpectedLiters) > 0 ? Number(b.ExpectedLiters) : round2_(barrels * barrelL * st.YieldPercent / 100);
  let stage, stageLabel;
  if (cancelled) { stage = 'cancel'; stageLabel = 'ยกเลิก'; }
  else if (harvested) { stage = 'done'; stageLabel = 'เก็บเป็นน้ำยา EM แล้ว'; }
  else if (age >= ready) { stage = 'ready'; stageLabel = 'ครบกำหนด — ควรเก็บเป็นน้ำยา EM'; }
  else if (age >= st.Stage2Days) { stage = 'late'; stageLabel = 'หมักช่วง ' + st.Stage2Days + '–' + (ready - 1) + ' วัน'; }
  else if (age >= st.Stage1Days) { stage = 'mid'; stageLabel = 'หมักช่วง ' + st.Stage1Days + '–' + (st.Stage2Days - 1) + ' วัน'; }
  else { stage = 'early'; stageLabel = 'หมักช่วงต้น (ไม่ถึง ' + st.Stage1Days + ' วัน)'; }
  const fermenting = stage === 'early' || stage === 'mid' || stage === 'late' || stage === 'ready';
  const progress = harvested ? 100 : (cancelled ? 0 : Math.min(100, Math.round(age / ready * 100)));

  // เช็กพอยต์ตรวจทุก N วัน (7, 14, 21, …) ก่อนครบกำหนด
  const interval = st.CheckIntervalDays;
  const myChecks = (checks || []).filter(c => c.BatchID === b.BatchID).map(c => ({
    CheckID: c.CheckID, CheckDate: toIsoDate_(c.CheckDate), Week: Number(c.Week) || 0, DayNo: Number(c.DayNo) || 0,
    PH: c.PH === '' || c.PH === null || c.PH === undefined ? null : Number(c.PH), Smell: c.Smell || '', Appearance: c.Appearance || '',
    Action: c.Action || '', Note: c.Note || '', RecordedBy: c.RecordedBy || ''
  })).filter(c => { const cn = emDayNum_(c.CheckDate); return cn === null || cn <= asOfN; }).sort((a, b2) => String(a.CheckDate).localeCompare(String(b2.CheckDate)));
  const weeks = [];
  const nW = Math.max(1, Math.floor(ready / interval));
  for (let k = 1; k <= nW; k++) {
    const dueDay = interval * k, chk = myChecks.find(c => c.Week === k);
    let wst = 'upcoming';
    if (chk) wst = 'done';
    else if (cancelled || harvested && harvestN !== null && harvestN < (startN || 0) + dueDay) wst = 'skip';
    else if (age > dueDay) wst = 'overdue';
    else if (age === dueDay) wst = 'due';
    weeks.push({ week: k, dueDay: dueDay, dueDate: startN === null ? null : emIso_(startN + dueDay), status: wst, check: chk || null });
  }
  const pending = fermenting ? weeks.filter(w => w.status === 'due' || w.status === 'overdue') : [];
  const nextWeek = fermenting ? weeks.find(w => w.status === 'upcoming') : null;
  // หมุดหมายถัดไปของการหมัก
  let nextMilestone = null;
  if (fermenting && age < ready) {
    const cand = [st.Stage1Days, st.Stage2Days, ready].filter(d => d > age);
    const d = cand[0];
    nextMilestone = { day: d, date: startN === null ? null : emIso_(startN + d), inDays: d - age, label: d === ready ? 'ครบกำหนดเป็นน้ำยา EM' : 'ครบ ' + d + ' วัน' };
  }
  const used = usedL || 0, yieldL = harvested ? (Number(b.YieldLiters) || 0) : 0;
  return {
    BatchID: b.BatchID, BatchNo: Number(b.BatchNo) || 0, StartDate: toIsoDate_(b.StartDate), age: age, readyDays: ready,
    readyDate: startN === null ? null : emIso_(startN + ready), daysToReady: ready - age,
    progress: progress, stage: stage, stageLabel: stageLabel, fermenting: fermenting,
    barrels: barrels, barrelLiters: barrelL, expectedLiters: expected,
    foodKg: Number(b.FoodKg) || 0, fruitVegKg: Number(b.FruitVegKg) || 0, meatKg: Number(b.MeatKg) || 0, inputKg: emInputKg_(b),
    molassesKg: Number(b.MolassesKg) || 0, waterL: Number(b.WaterL) || 0,
    location: b.Location || '', note: b.Note || '', status: status,
    harvestDate: b.HarvestDate ? toIsoDate_(b.HarvestDate) : null, yieldLiters: yieldL, harvestNote: b.HarvestNote || '',
    usedLiters: round2_(used), remainingLiters: round2_(yieldL - used),
    checks: myChecks, weeks: weeks, pendingChecks: pending.length, pendingWeeks: pending.map(w => w.week), nextMilestone: nextMilestone,
    nextCheckDate: nextWeek ? nextWeek.dueDate : null, nextCheckWeek: nextWeek ? nextWeek.week : null,
    recordedBy: b.RecordedBy || ''
  };
}

/** ภาพรวมทั้งหมดของโครงการ (เรียกครั้งเดียวต่อการเปิดแท็บ) */
function getEmOverview() {
  const st = getEmSettings_();
  const todayN = emToday_();
  const rawB = readRowsIfSheet_(SHEET_EM_BATCHES).filter(b => b.BatchID);
  const rawC = readRowsIfSheet_(SHEET_EM_CHECKS).filter(c => c.CheckID);
  const allU = readRowsIfSheet_(SHEET_EM_USAGE).filter(u => u.UsageID);
  const rawU = allU.filter(u => isVerified_(u.VerifyStatus));
  const allL = readRowsIfSheet_(SHEET_EM_LOGS).filter(l => l.LogID);
  const pendL = allL.filter(l => String(l.VerifyStatus || '') === VERIFY_PENDING);
  const okL = allL.filter(l => isVerified_(l.VerifyStatus));
  const pendU = allU.filter(u => String(u.VerifyStatus || '') === VERIFY_PENDING);
  const usedBy = {};
  rawU.forEach(u => { if (u.BatchID) usedBy[u.BatchID] = (usedBy[u.BatchID] || 0) + (Number(u.Liters) || 0); });
  const batches = rawB.map(b => emBatchView_(b, st, todayN, rawC, usedBy[b.BatchID] || 0))
    .sort((a, b) => String(b.StartDate).localeCompare(String(a.StartDate)) || b.BatchNo - a.BatchNo);
  const logsOf = {};
  okL.forEach(l => { const bid = l.BatchID; if (bid) (logsOf[bid] = logsOf[bid] || []).push(emLogView_(l)); });
  batches.forEach(b => { b.logs = (logsOf[b.BatchID] || []).sort((x, y) => String(y.LogDate).localeCompare(String(x.LogDate))); b.logCount = b.logs.length; b.lastLogDate = b.logs.length ? b.logs[0].LogDate : null; });
  const live = batches.filter(b => b.stage !== 'cancel');
  const byNo = {}; batches.forEach(b => { byNo[b.BatchID] = b.BatchNo; });

  const sum = (list, f) => list.reduce((s, i) => s + (Number(f(i)) || 0), 0);
  const stageDef = [
    { key: 'early', label: 'ไม่ถึง ' + st.Stage1Days + ' วัน', range: '0–' + (st.Stage1Days - 1) + ' วัน' },
    { key: 'mid', label: st.Stage1Days + '–' + (st.Stage2Days - 1) + ' วัน', range: st.Stage1Days + '–' + (st.Stage2Days - 1) + ' วัน' },
    { key: 'late', label: st.Stage2Days + '–' + (st.ReadyDays - 1) + ' วัน', range: st.Stage2Days + '–' + (st.ReadyDays - 1) + ' วัน' },
    { key: 'ready', label: 'ครบ ' + st.ReadyDays + ' วัน (พร้อมเก็บ)', range: '≥ ' + st.ReadyDays + ' วัน' }
  ];
  const byStage = stageDef.map(d => {
    const L = live.filter(b => b.stage === d.key);
    return { key: d.key, label: d.label, range: d.range, batches: L.length, barrels: sum(L, b => b.barrels), expectedLiters: round2_(sum(L, b => b.expectedLiters)), inputKg: round2_(sum(L, b => b.inputKg)) };
  });
  // ถังที่ "ครบ" 7 / 15 / 45 วันแล้ว (สะสม) และที่จะครบในอีก SoonDays วัน
  const milestones = [st.Stage1Days, st.Stage2Days, st.ReadyDays].map(M => {
    const reached = live.filter(b => (b.stage === 'done' && b.harvestDate && emDayNum_(b.harvestDate) - emDayNum_(b.StartDate) < M) ? false : b.age >= M || b.stage === 'done');
    const soon = live.filter(b => b.fermenting && b.age < M && M - b.age <= st.SoonDays);
    const waiting = live.filter(b => b.fermenting && b.age < M);
    return { days: M, reachedBatches: reached.length, reachedBarrels: sum(reached, b => b.barrels), reachedExpectedLiters: round2_(sum(reached, b => b.expectedLiters)),
             soonBatches: soon.length, soonBarrels: sum(soon, b => b.barrels), waitingBarrels: sum(waiting, b => b.barrels) };
  });
  // ถังที่เริ่มหมักรายวัน 30 วันล่าสุด
  const daily = [];
  for (let i = 29; i >= 0; i--) {
    const n = todayN - i, L = live.filter(b => emDayNum_(b.StartDate) === n);
    daily.push({ date: emIso_(n), batches: L.length, barrels: sum(L, b => b.barrels), kg: round2_(sum(L, b => b.inputKg)) });
  }
  // คาดการณ์น้ำยาที่จะพร้อมเก็บในแต่ละสัปดาห์ถัดไป (8 สัปดาห์)
  const forecast = [];
  for (let w = 0; w < 8; w++) {
    const from = todayN + w * 7, to = from + 6;
    const L = live.filter(b => b.fermenting && (() => { const rn = emDayNum_(b.readyDate); return w === 0 ? rn <= to : (rn >= from && rn <= to); })());
    forecast.push({ from: emIso_(from), to: emIso_(to), batches: L.length, barrels: sum(L, b => b.barrels), expectedLiters: round2_(sum(L, b => b.expectedLiters)), overdue: w === 0 ? L.filter(b => b.stage === 'ready').length : 0 });
  }
  const harvestedL = round2_(sum(live.filter(b => b.stage === 'done'), b => b.yieldLiters));
  const usedAll = round2_(sum(rawU, u => u.Liters));
  const usage = rawU.map(u => ({ UsageID: u.UsageID, BatchID: u.BatchID || '', BatchNo: u.BatchID ? (byNo[u.BatchID] || 0) : 0, UsageDate: toIsoDate_(u.UsageDate),
    Liters: Number(u.Liters) || 0, Place: u.Place || '', UsedBy: u.UsedBy || '', Purpose: u.Purpose || '', Note: u.Note || '', RecordedBy: u.RecordedBy || '' }))
    .sort((a, b) => String(b.UsageDate).localeCompare(String(a.UsageDate)));
  const rank = (field) => {
    const m = {}; usage.forEach(u => { const k = u[field] || '(ไม่ระบุ)'; const r = m[k] = m[k] || { name: k, liters: 0, times: 0, last: '' }; r.liters += u.Liters; r.times++; if (u.UsageDate > r.last) r.last = u.UsageDate; });
    return Object.keys(m).map(k => Object.assign(m[k], { liters: round2_(m[k].liters) })).sort((a, b) => b.liters - a.liters).map((r, i) => Object.assign(r, { rank: i + 1 }));
  };
  const ready = live.filter(b => b.stage === 'ready');
  const soonReady = live.filter(b => b.fermenting && b.stage !== 'ready' && b.daysToReady <= st.SoonDays).sort((a, b) => a.daysToReady - b.daysToReady);
  const pendingChecks = live.filter(b => b.pendingChecks > 0).map(b => ({ BatchID: b.BatchID, BatchNo: b.BatchNo, age: b.age, weeks: b.pendingWeeks }));
  const fermenting = live.filter(b => b.fermenting);
  return {
    settings: st, today: emIso_(todayN), batches: batches, usage: usage,
    totals: {
      batches: live.length, barrels: sum(live, b => b.barrels),
      foodKg: round2_(sum(live, b => b.foodKg)), fruitVegKg: round2_(sum(live, b => b.fruitVegKg)), meatKg: round2_(sum(live, b => b.meatKg)),
      inputKg: round2_(sum(live, b => b.inputKg)), molassesKg: round2_(sum(live, b => b.molassesKg)), waterL: round2_(sum(live, b => b.waterL)),
      fermentingBatches: fermenting.length, fermentingBarrels: sum(fermenting, b => b.barrels), pipelineLiters: round2_(sum(fermenting, b => b.expectedLiters)),
      readyBatches: ready.length, readyBarrels: sum(ready, b => b.barrels), readyLiters: round2_(sum(ready, b => b.expectedLiters)),
      harvestedBatches: live.filter(b => b.stage === 'done').length, harvestedLiters: harvestedL, usedLiters: usedAll, stockLiters: round2_(harvestedL - usedAll),
      pendingLogs: pendL.length, dailyLogs: okL.length,
      pendingUsage: pendU.length, pendingUsageLiters: round2_(sum(pendU, u => u.Liters)),
      usageTimes: usage.length, usagePlaces: new Set(usage.map(u => u.Place).filter(Boolean)).size, usageUsers: new Set(usage.map(u => u.UsedBy).filter(Boolean)).size,
      pendingChecks: pendingChecks.reduce((s, b) => s + b.weeks.length, 0)
    },
    byStage: byStage, milestones: milestones, daily: daily, forecast: forecast,
    readyList: ready.map(b => b.BatchID), soonList: soonReady.map(b => b.BatchID), pendingChecks: pendingChecks,
    byPlace: rank('Place'), byUser: rank('UsedBy'),
    recentLogs: okL.map(l => Object.assign(emLogView_(l), { BatchNo: l.BatchID ? (byNo[l.BatchID] || 0) : 0 })).sort((a, b) => String(b.LogDate).localeCompare(String(a.LogDate))).slice(0, 40),
    stockByBatch: live.filter(b => b.stage === 'done').map(b => ({ BatchID: b.BatchID, BatchNo: b.BatchNo, harvestDate: b.harvestDate, yieldLiters: b.yieldLiters, usedLiters: b.usedLiters, remainingLiters: b.remainingLiters }))
      .sort((a, b) => String(a.harvestDate).localeCompare(String(b.harvestDate)))
  };
}

/**
 * รายงานความคืบหน้าทุก 7 วัน: ช่วงละ 7 วัน นับถอยหลังจากวันที่สิ้นสุด (endDate, ค่าเริ่มต้น = วันนี้) ย้อนหลัง `periods` ช่วง
 * แต่ละช่วงสรุป: รอบ/ถัง/กก.ที่เริ่มหมัก • ถังที่ครบ 7/15/ครบกำหนด • น้ำยาที่เก็บได้ • ที่ใช้ไป • คงเหลือ ณ สิ้นช่วง • การตรวจที่ถึงกำหนด/บันทึกแล้ว
 * และ "สถานะแต่ละรอบ ณ วันสิ้นสุด" สำหรับช่วงล่าสุด
 */
function getEmWeeklyReport(endDate, periods) {
  const st = getEmSettings_();
  let endN = emDayNum_(endDate); if (endN === null) endN = emToday_();
  const P = Math.min(26, Math.max(1, Number(periods) || 8));
  const rawB = readRowsIfSheet_(SHEET_EM_BATCHES).filter(b => b.BatchID && (b.Status || EM_STATUS_FERMENT) !== EM_STATUS_CANCEL);
  const rawC = readRowsIfSheet_(SHEET_EM_CHECKS).filter(c => c.CheckID);
  const rawU = readRowsIfSheet_(SHEET_EM_USAGE).filter(u => u.UsageID && isVerified_(u.VerifyStatus));
  const barrelsOf = b => Number(b.BarrelCount) || 0;
  const sum = (list, f) => list.reduce((s, i) => s + (Number(f(i)) || 0), 0);
  const stockAt = n => round2_(sum(rawB.filter(b => b.Status === EM_STATUS_DONE && emDayNum_(b.HarvestDate) !== null && emDayNum_(b.HarvestDate) <= n), b => b.YieldLiters)
                             - sum(rawU.filter(u => emDayNum_(u.UsageDate) !== null && emDayNum_(u.UsageDate) <= n), u => u.Liters));
  const interval = st.CheckIntervalDays;
  const series = [];
  for (let i = 0; i < P; i++) {
    const to = endN - 7 * (P - 1 - i), from = to - 6;
    const inP = n => n !== null && n >= from && n <= to;
    const started = rawB.filter(b => inP(emDayNum_(b.StartDate)));
    const har = rawB.filter(b => b.Status === EM_STATUS_DONE && inP(emDayNum_(b.HarvestDate)));
    const reach = M => rawB.filter(b => { const sn = emDayNum_(b.StartDate); return sn !== null && inP(sn + M); });
    const ferment = rawB.filter(b => { const sn = emDayNum_(b.StartDate); const hn = b.Status === EM_STATUS_DONE ? emDayNum_(b.HarvestDate) : null; return sn !== null && sn <= to && (hn === null || hn > to); });
    const readyWait = ferment.filter(b => (Number(b.ReadyDays) > 0 ? Number(b.ReadyDays) : st.ReadyDays) <= to - emDayNum_(b.StartDate));
    let due = 0, done = 0;
    ferment.forEach(b => {
      const sn = emDayNum_(b.StartDate), rd = Number(b.ReadyDays) > 0 ? Number(b.ReadyDays) : st.ReadyDays, nW = Math.max(1, Math.floor(rd / interval));
      for (let k = 1; k <= nW; k++) {
        const dd = sn + interval * k;
        if (dd < from || dd > to) continue;
        due++;
        if (rawC.some(c => c.BatchID === b.BatchID && Number(c.Week) === k && emDayNum_(c.CheckDate) !== null && emDayNum_(c.CheckDate) <= to)) done++;
      }
    });
    series.push({
      from: emIso_(from), to: emIso_(to), startedBatches: started.length, startedBarrels: sum(started, barrelsOf), startedKg: round2_(sum(started, emInputKg_)),
      reach1Barrels: sum(reach(st.Stage1Days), barrelsOf), reach2Barrels: sum(reach(st.Stage2Days), barrelsOf), reachReadyBarrels: sum(reach(st.ReadyDays), barrelsOf),
      harvestedBatches: har.length, harvestedLiters: round2_(sum(har, b => b.YieldLiters)),
      usedLiters: round2_(sum(rawU.filter(u => inP(emDayNum_(u.UsageDate))), u => u.Liters)), usageTimes: rawU.filter(u => inP(emDayNum_(u.UsageDate))).length,
      stockEnd: stockAt(to), fermentingBarrels: sum(ferment, barrelsOf), fermentingBatches: ferment.length,
      readyWaitingBatches: readyWait.length, readyWaitingBarrels: sum(readyWait, barrelsOf),
      checksDue: due, checksDone: done
    });
  }
  const usedBy = {};
  rawU.forEach(u => { if (u.BatchID && emDayNum_(u.UsageDate) !== null && emDayNum_(u.UsageDate) <= endN) usedBy[u.BatchID] = (usedBy[u.BatchID] || 0) + (Number(u.Liters) || 0); });
  const snapshot = rawB.filter(b => emDayNum_(b.StartDate) !== null && emDayNum_(b.StartDate) <= endN)
    .map(b => emBatchView_(b, st, endN, rawC, usedBy[b.BatchID] || 0))
    .filter(b => b.fermenting || (b.stage === 'done' && b.harvestDate && emDayNum_(b.harvestDate) > endN - 7))
    .sort((a, b) => b.age - a.age);
  return { settings: st, endDate: emIso_(endN), periods: P, series: series, current: series[series.length - 1], previous: series.length > 1 ? series[series.length - 2] : null, snapshot: snapshot };
}

// ----------------------------- WRITE (ผู้ดูแลระบบ) -----------------------------
function emNum_(v, label, allowZero) {
  if (v === '' || v === null || v === undefined) return 0;
  const n = Number(v);
  if (isNaN(n) || n < 0) throw new Error(label + ' ต้องเป็นตัวเลขที่ไม่ติดลบ');
  return n;
}
function emDate_(v, label) {
  const n = emDayNum_(v);
  if (n === null) throw new Error('กรุณาระบุ' + label + 'ให้ถูกต้อง');
  return new Date(n * 86400000 + 12 * 3600000);
}
function emSetFields_(sheetName, idField, id, fields) {
  const found = findRowById_(sheetName, idField, id);
  if (!found) throw new Error('ไม่พบรายการ');
  const sh = getSheet_(sheetName);
  Object.keys(fields).forEach(k => { const col = found.headers.indexOf(k); if (col !== -1) sh.getRange(found.rowIndex, col + 1).setValue(fields[k]); });
  return found;
}
function emUsedForBatch_(batchId) {
  return readRowsIfSheet_(SHEET_EM_USAGE).filter(u => u.BatchID === batchId && isVerified_(u.VerifyStatus)).reduce((s, u) => s + (Number(u.Liters) || 0), 0);
}
function emStockLiters_() {
  const h = readRowsIfSheet_(SHEET_EM_BATCHES).filter(b => b.Status === EM_STATUS_DONE).reduce((s, b) => s + (Number(b.YieldLiters) || 0), 0);
  const u = readRowsIfSheet_(SHEET_EM_USAGE).filter(x => isVerified_(x.VerifyStatus)).reduce((s, x) => s + (Number(x.Liters) || 0), 0);
  return round2_(h - u);
}

/** เพิ่มรอบการหมัก (ไม่ล็อก/ไม่เช็กสิทธิ์ — ผู้เรียกต้องจัดการเอง) */
function emInsertBatch_(data, recordedBy) {
  data = data || {};
  const food = emNum_(data.FoodKg, 'เศษอาหาร'), fv = emNum_(data.FruitVegKg, 'ผักและผลไม้'), meat = emNum_(data.MeatKg, 'เนื้อสัตว์');
  if (food + fv + meat <= 0) throw new Error('กรุณาระบุปริมาณเศษอาหาร ผักและผลไม้ หรือเนื้อสัตว์ อย่างน้อย 1 ช่อง (กก.)');
  const barrels = Math.round(emNum_(data.BarrelCount, 'จำนวนถัง'));
  if (barrels <= 0) throw new Error('กรุณาระบุจำนวนถังที่หมัก');
  const start = emDate_(data.StartDate || new Date(), 'วันที่เริ่มหมัก');
  const st = getEmSettings_();
  ensureEmSheets_();
  const no = readRowsIfSheet_(SHEET_EM_BATCHES).reduce((m, b) => Math.max(m, Number(b.BatchNo) || 0), 0) + 1;
  const now = new Date(), id = newId_('EMB');
  const barrelL = emNum_(data.BarrelLiters, 'ขนาดถัง') || st.DefaultBarrelLiters;
  appendByHeaders_(SHEET_EM_BATCHES, {
    BatchID: id, BatchNo: no, StartDate: start, FoodKg: food, FruitVegKg: fv, MeatKg: meat, MolassesKg: emNum_(data.MolassesKg, 'กากน้ำตาล'), WaterL: emNum_(data.WaterL, 'น้ำ'),
    BarrelCount: barrels, BarrelLiters: barrelL, ExpectedLiters: emNum_(data.ExpectedLiters, 'ปริมาณน้ำยาที่คาดว่าจะได้') || round2_(barrels * barrelL * st.YieldPercent / 100),
    ReadyDays: emNum_(data.ReadyDays, 'จำนวนวันที่ใช้หมัก') || '', Location: String(data.Location || '').trim(), Note: String(data.Note || '').trim(),
    Status: EM_STATUS_FERMENT, RecordedBy: recordedBy, CreatedAt: now, UpdatedAt: now
  });
  return { ok: true, batchId: id, batchNo: no };
}

function addEmBatch(departmentId, pin, data) {
  const auth = assertAuthorized_(departmentId, pin);
  data = data || {};
  // ตรวจรูปแบบข้อมูลก่อนล็อก (ให้ข้อความผิดพลาดชัดเจน)
  emNum_(data.FoodKg, 'เศษอาหาร'); emNum_(data.FruitVegKg, 'ผักและผลไม้'); emNum_(data.MeatKg, 'เนื้อสัตว์');
  return withShareLock_(() => {
    const r = emInsertBatch_(data, auth.departmentName);
    syncEmLinkedProject_();
    return r;
  });
}

function updateEmBatch(departmentId, pin, batchId, data) {
  assertAuthorized_(departmentId, pin);
  data = data || {};
  return withShareLock_(() => {
    const found = findRowById_(SHEET_EM_BATCHES, 'BatchID', batchId);
    if (!found) throw new Error('ไม่พบรอบการหมัก');
    const f = {};
    if (data.StartDate !== undefined) f.StartDate = emDate_(data.StartDate, 'วันที่เริ่มหมัก');
    ['FoodKg', 'FruitVegKg', 'MeatKg', 'MolassesKg', 'WaterL', 'BarrelLiters', 'ExpectedLiters'].forEach(k => { if (data[k] !== undefined) f[k] = emNum_(data[k], k); });
    if (data.BarrelCount !== undefined) { const b = Math.round(emNum_(data.BarrelCount, 'จำนวนถัง')); if (b <= 0) throw new Error('จำนวนถังต้องมากกว่า 0'); f.BarrelCount = b; }
    if (data.ReadyDays !== undefined) f.ReadyDays = emNum_(data.ReadyDays, 'จำนวนวันที่ใช้หมัก') || '';
    ['Location', 'Note'].forEach(k => { if (data[k] !== undefined) f[k] = String(data[k] || '').trim(); });
    if (data.Status !== undefined) {
      if ([EM_STATUS_FERMENT, EM_STATUS_CANCEL].indexOf(data.Status) === -1) throw new Error('สถานะไม่ถูกต้อง');
      const cur = found.values[found.headers.indexOf('Status')] || EM_STATUS_FERMENT;
      if (cur === EM_STATUS_DONE) {
        if (emUsedForBatch_(batchId) > 0) throw new Error('รอบนี้มีการนำน้ำยาไปใช้แล้ว ไม่สามารถยกเลิกการเก็บน้ำยาได้ (ลบรายการใช้งานก่อน)');
        f.HarvestDate = ''; f.YieldLiters = ''; f.HarvestNote = '';
      }
      f.Status = data.Status;
    }
    f.UpdatedAt = new Date();
    emSetFields_(SHEET_EM_BATCHES, 'BatchID', batchId, f);
    syncEmLinkedProject_();
    return { ok: true };
  });
}

function deleteEmBatch(departmentId, pin, batchId) {
  assertAuthorized_(departmentId, pin);
  return withShareLock_(() => {
    const found = findRowById_(SHEET_EM_BATCHES, 'BatchID', batchId);
    if (!found) throw new Error('ไม่พบรอบการหมัก');
    if (emUsedForBatch_(batchId) > 0) throw new Error('รอบนี้มีการนำน้ำยาไปใช้แล้ว ลบไม่ได้ (ถ้าบันทึกผิด ให้ลบรายการใช้งานก่อน หรือเปลี่ยนสถานะเป็นยกเลิก)');
    // ลบผลตรวจของรอบนี้ แล้วลบรอบ (ลบจากท้ายขึ้นมาเพื่อไม่ให้ลำดับแถวเลื่อน)
    const csh = getSheet_(SHEET_EM_CHECKS), cv = csh.getDataRange().getValues(), cb = cv[0].indexOf('BatchID');
    for (let r = cv.length - 1; r >= 1; r--) if (cv[r][cb] === batchId) csh.deleteRow(r + 1);
    const lsh = findSheet_(SHEET_EM_LOGS);
    if (lsh) { const lv = lsh.getDataRange().getValues(), lb = lv[0].indexOf('BatchID'); if (lb !== -1) for (let r = lv.length - 1; r >= 1; r--) if (lv[r][lb] === batchId) getSheet_(SHEET_EM_LOGS).deleteRow(r + 1); }
    getSheet_(SHEET_EM_BATCHES).deleteRow(findRowById_(SHEET_EM_BATCHES, 'BatchID', batchId).rowIndex);
    syncEmLinkedProject_();
    return { ok: true };
  });
}

/** เก็บเป็นน้ำยา EM: บันทึกวันที่และปริมาณน้ำยาที่ได้จริง (ลิตร) ของรอบนั้น */
function harvestEmBatch(departmentId, pin, batchId, data) {
  assertAuthorized_(departmentId, pin);
  data = data || {};
  const liters = emNum_(data.YieldLiters, 'ปริมาณน้ำยา EM ที่ได้');
  if (liters <= 0) throw new Error('กรุณาระบุปริมาณน้ำยา EM ที่ได้ (ลิตร)');
  const hd = emDate_(data.HarvestDate || new Date(), 'วันที่เก็บน้ำยา EM');
  return withShareLock_(() => {
    const found = findRowById_(SHEET_EM_BATCHES, 'BatchID', batchId);
    if (!found) throw new Error('ไม่พบรอบการหมัก');
    const cur = found.values[found.headers.indexOf('Status')] || EM_STATUS_FERMENT;
    if (cur === EM_STATUS_CANCEL) throw new Error('รอบนี้ถูกยกเลิกแล้ว');
    const startN = emDayNum_(found.values[found.headers.indexOf('StartDate')]);
    if (startN !== null && emDayNum_(hd) < startN) throw new Error('วันที่เก็บน้ำยาต้องไม่ก่อนวันที่เริ่มหมัก');
    if (cur === EM_STATUS_DONE && liters < emUsedForBatch_(batchId)) throw new Error('ปริมาณที่ได้ต้องไม่น้อยกว่าที่นำไปใช้แล้ว (' + round2_(emUsedForBatch_(batchId)) + ' ลิตร)');
    emSetFields_(SHEET_EM_BATCHES, 'BatchID', batchId, { Status: EM_STATUS_DONE, HarvestDate: hd, YieldLiters: liters, HarvestNote: String(data.Note || '').trim(), UpdatedAt: new Date() });
    syncEmLinkedProject_();
    return { ok: true };
  });
}

/** เพิ่มผลตรวจ (ไม่ล็อก/ไม่เช็กสิทธิ์) */
function emInsertCheck_(batchId, data, recordedBy) {
  data = data || {};
  const cd = emDate_(data.CheckDate || new Date(), 'วันที่ตรวจ');
  const ph = (data.PH === '' || data.PH === null || data.PH === undefined) ? '' : Number(data.PH);
  if (ph !== '' && (isNaN(ph) || ph < 0 || ph > 14)) throw new Error('ค่า pH ต้องอยู่ระหว่าง 0–14');
  ensureEmSheets_();
  const b = findRowById_(SHEET_EM_BATCHES, 'BatchID', batchId);
  if (!b) throw new Error('ไม่พบรอบการหมัก');
  const startN = emDayNum_(b.values[b.headers.indexOf('StartDate')]);
  const dayNo = startN === null ? 0 : emDayNum_(cd) - startN;
  if (dayNo < 0) throw new Error('วันที่ตรวจต้องไม่ก่อนวันที่เริ่มหมัก');
  const week = Math.round(Number(data.Week) || 0);
  if (week > 0 && readRowsIfSheet_(SHEET_EM_CHECKS).some(c => c.BatchID === batchId && Number(c.Week) === week)) throw new Error('บันทึกผลตรวจของสัปดาห์ที่ ' + week + ' ของรอบนี้ไว้แล้ว (ลบรายการเดิมก่อนถ้าต้องการบันทึกใหม่)');
  const id = newId_('EMC');
  appendByHeaders_(SHEET_EM_CHECKS, { CheckID: id, BatchID: batchId, CheckDate: cd, Week: week, DayNo: dayNo, PH: ph, Smell: String(data.Smell || '').trim(),
    Appearance: String(data.Appearance || '').trim(), Action: String(data.Action || '').trim(), Note: String(data.Note || '').trim(), RecordedBy: recordedBy, CreatedAt: new Date() });
  return { ok: true, checkId: id };
}

function addEmCheck(departmentId, pin, batchId, data) {
  const auth = assertAuthorized_(departmentId, pin);
  data = data || {};
  return withShareLock_(() => {
    const r = emInsertCheck_(batchId, data, auth.departmentName);
    syncEmLinkedProject_();
    return r;
  });
}

function deleteEmCheck(departmentId, pin, checkId) {
  assertAuthorized_(departmentId, pin);
  const f = findRowById_(SHEET_EM_CHECKS, 'CheckID', checkId);
  if (!f) throw new Error('ไม่พบรายการ');
  getSheet_(SHEET_EM_CHECKS).deleteRow(f.rowIndex);
  return { ok: true };
}

/** บันทึกการนำน้ำยา EM ไปใช้: วันที่ ปริมาณ สถานที่ ผู้ใช้ วัตถุประสงค์ (ไม่เกินปริมาณคงเหลือ) */
function addEmUsage(departmentId, pin, data) {
  const auth = assertAuthorized_(departmentId, pin);
  data = data || {};
  const liters = emNum_(data.Liters, 'ปริมาณที่นำไปใช้');
  if (liters <= 0) throw new Error('กรุณาระบุปริมาณที่นำไปใช้ (ลิตร)');
  const place = String(data.Place || '').trim(), who = String(data.UsedBy || '').trim();
  if (!place) throw new Error('กรุณาระบุสถานที่ที่นำไปใช้');
  if (!who) throw new Error('กรุณาระบุชื่อผู้นำไปใช้');
  const ud = emDate_(data.UsageDate || new Date(), 'วันที่นำไปใช้');
  return withShareLock_(() => {
    ensureEmSheets_();
    const stock = emStockLiters_();
    if (liters > stock + 1e-9) throw new Error('น้ำยา EM คงเหลือไม่พอ (คงเหลือ ' + stock + ' ลิตร)');
    const batchId = String(data.BatchID || '').trim();
    if (batchId) {
      const b = findRowById_(SHEET_EM_BATCHES, 'BatchID', batchId);
      if (!b) throw new Error('ไม่พบรอบการหมักที่ระบุ');
      if (b.values[b.headers.indexOf('Status')] !== EM_STATUS_DONE) throw new Error('รอบที่เลือกยังไม่ได้เก็บเป็นน้ำยา EM');
      const remain = Number(b.values[b.headers.indexOf('YieldLiters')] || 0) - emUsedForBatch_(batchId);
      if (liters > remain + 1e-9) throw new Error('น้ำยาของรอบนี้คงเหลือไม่พอ (คงเหลือ ' + round2_(remain) + ' ลิตร)');
    }
    const id = newId_('EMU');
    appendByHeaders_(SHEET_EM_USAGE, { UsageID: id, BatchID: batchId, UsageDate: ud, Liters: liters, Place: place, UsedBy: who, Purpose: String(data.Purpose || '').trim(),
      Note: String(data.Note || '').trim(), RecordedBy: auth.departmentName, CreatedAt: new Date() });
    syncEmLinkedProject_();
    return { ok: true, usageId: id, stockLiters: round2_(stock - liters) };
  });
}

function deleteEmUsage(departmentId, pin, usageId) {
  assertAuthorized_(departmentId, pin);
  const f = findRowById_(SHEET_EM_USAGE, 'UsageID', usageId);
  if (!f) throw new Error('ไม่พบรายการ');
  getSheet_(SHEET_EM_USAGE).deleteRow(f.rowIndex);
  syncEmLinkedProject_();
  return { ok: true };
}

/** ซิงก์สถานะแถวโครงการในหน้าภาพรวม */
function syncEmLinkedProject_() {
  try {
    const t = getEmOverview().totals;
    updateLinkedProjectFields_(LINKED_PROJECT_EM_ID, {
      CurrentStatus: t.batches
        ? `หมักแล้ว ${t.batches} รอบ (${t.barrels} ถัง) • กำลังหมัก ${t.fermentingBarrels} ถัง${t.readyBarrels ? ' (ครบกำหนดเก็บ ' + t.readyBarrels + ' ถัง)' : ''} • ได้น้ำยา EM ${t.harvestedLiters.toLocaleString()} ลิตร • ใช้แล้ว ${t.usedLiters.toLocaleString()} • คงเหลือ ${t.stockLiters.toLocaleString()} ลิตร`
        : 'เริ่มโครงการ — ยังไม่มีรอบการหมัก (ดูรายละเอียดที่แท็บ "น้ำยา EM")',
      Progress: yearProgressPct_()
    });
  } catch (e) { /* ยังไม่มีชีต/ข้อมูล ก็ข้ามไปเงียบๆ */ }
}


// ----------------------------- ฟอร์มสาธารณะ: บันทึกการนำน้ำยา EM ไปใช้ (ลิงก์ / QR Code) -----------------------------
/** ลิงก์ฟอร์ม "บันทึกการนำน้ำยา EM ไปใช้" สำหรับแชร์/สร้าง QR Code (ไม่ต้องมี PIN) */
function getEmUsageFormLink() {
  return (ScriptApp.getService().getUrl() || '') + '?page=em&action=use';
}

/**
 * ผู้ใช้น้ำยา EM บันทึกการนำไปใช้เองผ่านลิงก์/QR Code — รายการจะเป็น "รอยืนยัน" ยังไม่หักจากคงเหลือ
 * จนกว่าผู้ดูแลระบบจะยืนยัน (verifySubmissions key = 'emuse')
 */
function submitEmUsage(data) {
  data = data || {};
  if (String(data.website || '').trim()) return { ok: true, pending: true }; // ช่องล่อบอท (honeypot)
  const liters = emNum_(data.Liters, 'ปริมาณที่นำไปใช้');
  if (liters <= 0) throw new Error('กรุณาระบุปริมาณที่นำไปใช้ (ลิตร)');
  if (liters > 100000) throw new Error('ปริมาณมากผิดปกติ กรุณาตรวจสอบอีกครั้ง');
  const place = String(data.Place || '').trim().slice(0, 150), who = String(data.UsedBy || '').trim().slice(0, 120);
  if (!place) throw new Error('กรุณาระบุสถานที่ที่นำไปใช้');
  if (!who) throw new Error('กรุณาระบุชื่อผู้นำไปใช้');
  const ud = emDate_(data.UsageDate || new Date(), 'วันที่นำไปใช้');
  const todayN = emToday_(), un = emDayNum_(ud);
  if (un > todayN) throw new Error('วันที่นำไปใช้ต้องไม่เป็นวันในอนาคต');
  if (un < todayN - 90) throw new Error('วันที่นำไปใช้ย้อนหลังได้ไม่เกิน 90 วัน (หากเก่ากว่านี้ กรุณาแจ้งผู้ดูแลระบบ)');
  return withShareLock_(() => {
    ensureEmSheets_();
    const rows = readRowsIfSheet_(SHEET_EM_USAGE);
    if (rows.filter(u => String(u.VerifyStatus || '') === VERIFY_PENDING).length >= 200) throw new Error('มีรายการรอยืนยันจำนวนมากแล้ว กรุณาแจ้งผู้ดูแลระบบ');
    const stock = emStockLiters_();
    if (liters > stock + 1e-9) throw new Error('ขณะนี้น้ำยา EM คงเหลือ ' + stock + ' ลิตร ไม่พอตามที่ระบุ กรุณาตรวจสอบปริมาณ');
    const id = newId_('EMU');
    appendByHeaders_(SHEET_EM_USAGE, { UsageID: id, BatchID: '', UsageDate: ud, Liters: liters, Place: place, UsedBy: who, Purpose: String(data.Purpose || '').trim().slice(0, 200),
      Note: String(data.Note || '').trim().slice(0, 300), RecordedBy: 'ผู้ใช้บันทึกเองผ่านลิงก์', CreatedAt: new Date(), VerifyStatus: VERIFY_PENDING });
    return { ok: true, usageId: id, pending: true };
  });
}


// ----------------------------- บันทึกการหมักประจำวัน (ผู้ปฏิบัติกรอกผ่านลิงก์/QR → ผู้ดูแลยืนยัน → ผูกกับรอบการหมัก) -----------------------------
/**
 * ประเภทบันทึก:
 *  - เริ่มหมักถังใหม่   : เมื่อยืนยัน → สร้างรอบการหมักใหม่ (วันที่ วัตถุดิบ จำนวนถัง สถานที่)
 *  - เติมวัตถุดิบเพิ่ม  : เมื่อยืนยัน → บวกวัตถุดิบ (กก.) เข้ารอบที่เลือก (และจำนวนถังเพิ่ม ถ้าระบุ)
 *  - ตรวจ/ดูแลประจำวัน : เมื่อยืนยัน → บันทึกเป็นผลตรวจของรอบที่เลือก (นับในผลตรวจทุก 7 วัน)
 * รายการที่ยังไม่ยืนยันจะไม่นับในข้อมูลใดๆ ทั้งสิ้น
 */
function getEmDailyLogFormLink() {
  return (ScriptApp.getService().getUrl() || '') + '?page=em&action=log';
}

function emLogPendingRow_(r) {
  const b = r.BatchID ? findRowById_(SHEET_EM_BATCHES, 'BatchID', r.BatchID) : null;
  const no = b ? Number(b.values[b.headers.indexOf('BatchNo')]) || 0 : 0;
  const kg = (Number(r.FoodKg) || 0) + (Number(r.FruitVegKg) || 0) + (Number(r.MeatKg) || 0);
  let detail = '';
  if (r.LogType === EM_LOG_NEW) detail = 'วัตถุดิบ ' + round2_(kg) + ' กก. (เศษอาหาร ' + (Number(r.FoodKg) || 0) + ' • ผักผลไม้ ' + (Number(r.FruitVegKg) || 0) + ' • เนื้อสัตว์ ' + (Number(r.MeatKg) || 0) + ') • ' + (Number(r.Barrels) || 0) + ' ถัง' + (r.Location ? ' • ' + r.Location : '');
  else if (r.LogType === EM_LOG_ADD) detail = 'เติมวัตถุดิบ ' + round2_(kg) + ' กก.' + (Number(r.Barrels) > 0 ? ' • เพิ่ม ' + Number(r.Barrels) + ' ถัง' : '');
  else detail = [r.PH !== '' && r.PH !== null && r.PH !== undefined ? 'pH ' + r.PH : '', r.Smell, r.Appearance, r.Action].filter(Boolean).join(' • ');
  return { id: r.LogID, at: r.CreatedAt ? toIsoDateTime_(r.CreatedAt) : '', date: toIsoDate_(r.LogDate), name: r.Operator || '', type: r.LogType, batchNo: no, detail: detail, note: r.Note || '' };
}

function submitEmDailyLog(data) {
  data = data || {};
  if (String(data.website || '').trim()) return { ok: true, pending: true }; // ช่องล่อบอท
  const type = String(data.LogType || '').trim();
  if (EM_LOG_TYPES.indexOf(type) === -1) throw new Error('กรุณาเลือกประเภทบันทึก');
  const who = String(data.Operator || '').trim().slice(0, 120);
  if (!who) throw new Error('กรุณาระบุชื่อผู้ปฏิบัติงาน');
  const ld = emDate_(data.LogDate || new Date(), 'วันที่ปฏิบัติงาน');
  const todayN = emToday_(), ln = emDayNum_(ld);
  if (ln > todayN) throw new Error('วันที่ปฏิบัติงานต้องไม่เป็นวันในอนาคต');
  if (ln < todayN - 90) throw new Error('วันที่ปฏิบัติงานย้อนหลังได้ไม่เกิน 90 วัน (หากเก่ากว่านี้ กรุณาแจ้งผู้ดูแลระบบ)');
  const food = emNum_(data.FoodKg, 'เศษอาหาร'), fv = emNum_(data.FruitVegKg, 'ผักและผลไม้'), meat = emNum_(data.MeatKg, 'เนื้อสัตว์');
  if (food > 100000 || fv > 100000 || meat > 100000) throw new Error('ปริมาณมากผิดปกติ กรุณาตรวจสอบอีกครั้ง');
  const barrels = Math.round(emNum_(data.Barrels, 'จำนวนถัง'));
  const ph = (data.PH === '' || data.PH === null || data.PH === undefined) ? '' : Number(data.PH);
  if (ph !== '' && (isNaN(ph) || ph < 0 || ph > 14)) throw new Error('ค่า pH ต้องอยู่ระหว่าง 0–14');
  const batchId = String(data.BatchID || '').trim();
  if (type === EM_LOG_NEW) {
    if (food + fv + meat <= 0) throw new Error('กรุณาระบุปริมาณเศษอาหาร ผักและผลไม้ หรือเนื้อสัตว์ อย่างน้อย 1 ช่อง (กก.)');
    if (barrels <= 0) throw new Error('กรุณาระบุจำนวนถังที่หมัก');
  } else {
    if (!batchId) throw new Error('กรุณาเลือกรอบการหมัก');
    if (type === EM_LOG_ADD && food + fv + meat <= 0) throw new Error('กรุณาระบุปริมาณวัตถุดิบที่เติม (กก.)');
    if (type === EM_LOG_CHECK && ph === '' && !String(data.Smell || '').trim() && !String(data.Appearance || '').trim() && !String(data.Action || '').trim() && !String(data.Note || '').trim()) throw new Error('กรุณาบันทึกผลที่ตรวจ/สิ่งที่ทำอย่างน้อย 1 ช่อง');
  }
  return withShareLock_(() => {
    ensureEmSheets_();
    if (batchId) {
      const b = findRowById_(SHEET_EM_BATCHES, 'BatchID', batchId);
      if (!b) throw new Error('ไม่พบรอบการหมักที่เลือก');
      const status = b.values[b.headers.indexOf('Status')] || EM_STATUS_FERMENT;
      if (status !== EM_STATUS_FERMENT) throw new Error('รอบที่เลือกไม่ได้อยู่ระหว่างการหมักแล้ว');
      const sn = emDayNum_(b.values[b.headers.indexOf('StartDate')]);
      if (sn !== null && ln < sn) throw new Error('วันที่ปฏิบัติงานต้องไม่ก่อนวันที่เริ่มหมักของรอบนี้');
    }
    if (readRowsIfSheet_(SHEET_EM_LOGS).filter(l => String(l.VerifyStatus || '') === VERIFY_PENDING).length >= 200) throw new Error('มีรายการรอยืนยันจำนวนมากแล้ว กรุณาแจ้งผู้ดูแลระบบ');
    const id = newId_('EML');
    appendByHeaders_(SHEET_EM_LOGS, { LogID: id, LogDate: ld, BatchID: batchId, LogType: type, FoodKg: food, FruitVegKg: fv, MeatKg: meat, Barrels: barrels,
      Location: String(data.Location || '').trim().slice(0, 120), PH: ph, Smell: String(data.Smell || '').trim().slice(0, 100), Appearance: String(data.Appearance || '').trim().slice(0, 200),
      Action: String(data.Action || '').trim().slice(0, 200), Note: String(data.Note || '').trim().slice(0, 300), Operator: who, CreatedAt: new Date(), VerifyStatus: VERIFY_PENDING });
    return { ok: true, logId: id, pending: true };
  });
}

/** ผู้ดูแลระบบยืนยันบันทึกประจำวัน: ทำผลตามประเภทและผูกกับรอบการหมัก (ยืนยันได้ทีละรายการ; รายการที่ทำผลไม่ได้จะถูกข้ามและแจ้งเหตุผล) */
function verifyEmLogs_(auth, ids) {
  ensureVerifyColumns_(SHEET_EM_LOGS);
  const want = {}; (ids || []).forEach(id => { want[id] = true; });
  return withShareLock_(() => {
    const sh = getSheet_(SHEET_EM_LOGS);
    const values = sh.getDataRange().getValues(), h = values[0];
    const col = n => h.indexOf(n), now = new Date();
    let n = 0; const errors = [];
    for (let r = 1; r < values.length; r++) {
      if (!want[values[r][col('LogID')]] || String(values[r][col('VerifyStatus')] || '') !== VERIFY_PENDING) continue;
      const o = {}; h.forEach((k, i) => { o[k] = values[r][i]; });
      try {
        const batchId = emApplyLog_(o, auth.departmentName);
        if (batchId && !o.BatchID) sh.getRange(r + 1, col('BatchID') + 1).setValue(batchId);
        sh.getRange(r + 1, col('VerifyStatus') + 1).setValue(VERIFY_OK);
        if (col('VerifiedBy') !== -1) sh.getRange(r + 1, col('VerifiedBy') + 1).setValue(auth.departmentName);
        if (col('VerifiedAt') !== -1) sh.getRange(r + 1, col('VerifiedAt') + 1).setValue(now);
        n++;
      } catch (e) { errors.push((o.Operator || '') + ' (' + toIsoDate_(o.LogDate) + '): ' + (e && e.message || e)); }
    }
    invalidateReadCache_();
    if (n) { try { syncEmLinkedProject_(); } catch (e) {} }
    if (!n && errors.length) throw new Error('ยืนยันไม่ได้: ' + errors.join(' | '));
    return { ok: true, changed: n, errors: errors };
  });
}

/** ทำผลของบันทึกประจำวันหนึ่งรายการ คืนค่า BatchID ที่เกี่ยวข้อง */
function emApplyLog_(o, by) {
  if (o.LogType === EM_LOG_NEW) {
    const r = emInsertBatch_({ StartDate: o.LogDate, FoodKg: o.FoodKg, FruitVegKg: o.FruitVegKg, MeatKg: o.MeatKg, BarrelCount: o.Barrels, Location: o.Location,
      Note: 'จากบันทึกประจำวันโดย ' + (o.Operator || '') + (o.Note ? ' • ' + o.Note : '') }, by);
    return r.batchId;
  }
  const f = findRowById_(SHEET_EM_BATCHES, 'BatchID', o.BatchID);
  if (!f) throw new Error('ไม่พบรอบการหมักที่เกี่ยวข้อง (อาจถูกลบแล้ว)');
  const val = k => f.values[f.headers.indexOf(k)];
  if ((val('Status') || EM_STATUS_FERMENT) !== EM_STATUS_FERMENT) throw new Error('รอบนี้ไม่ได้อยู่ระหว่างการหมักแล้ว');
  if (o.LogType === EM_LOG_ADD) {
    const upd = { FoodKg: (Number(val('FoodKg')) || 0) + (Number(o.FoodKg) || 0), FruitVegKg: (Number(val('FruitVegKg')) || 0) + (Number(o.FruitVegKg) || 0),
      MeatKg: (Number(val('MeatKg')) || 0) + (Number(o.MeatKg) || 0), UpdatedAt: new Date() };
    if (Number(o.Barrels) > 0) upd.BarrelCount = (Number(val('BarrelCount')) || 0) + Number(o.Barrels);
    emSetFields_(SHEET_EM_BATCHES, 'BatchID', o.BatchID, upd);
    return o.BatchID;
  }
  // ตรวจ/ดูแล → ผลตรวจ: เลือกครั้งที่ (สัปดาห์) ที่ถึงกำหนดและยังไม่ได้บันทึกให้อัตโนมัติ
  const st = getEmSettings_(), interval = st.CheckIntervalDays;
  const sn = emDayNum_(val('StartDate')), ln = emDayNum_(o.LogDate), age = sn === null ? 0 : ln - sn;
  const ready = Number(val('ReadyDays')) > 0 ? Number(val('ReadyDays')) : st.ReadyDays, nW = Math.max(1, Math.floor(ready / interval));
  const done = readRowsIfSheet_(SHEET_EM_CHECKS).filter(c => c.BatchID === o.BatchID).map(c => Number(c.Week));
  let week = 0;
  for (let k = 1; k <= nW; k++) { if (interval * k <= age && done.indexOf(k) === -1) { week = k; break; } }
  emInsertCheck_(o.BatchID, { CheckDate: o.LogDate, Week: week, PH: o.PH, Smell: o.Smell, Appearance: o.Appearance, Action: o.Action,
    Note: (o.Note ? o.Note + ' • ' : '') + 'จากบันทึกประจำวันโดย ' + (o.Operator || '') }, by);
  return o.BatchID;
}


// ============================================================================================
// เชื่อมต่อระบบอื่น (Google Sheets / Apps Script ที่มีข้อมูลสัมพันธ์กัน)
//  ① ส่งออก  : แอปอื่น/แดชบอร์ดกลางเรียก  <WebApp URL>?api=summary&key=<API KEY>  → JSON ตัวเลขสรุปของโครงการที่เปิดให้ (ไม่มีข้อมูลส่วนบุคคล)
//  ② ส่งเข้า : แอปอื่นเรียก doPost (JSON) ส่งปริมาณเศษอาหาร/ผักผลไม้/เนื้อสัตว์/น้ำมันพืชรายวัน และรายชื่อหน่วยงาน → เก็บในชีต ExternalWasteDaily / ExternalDepartments
//  ③ ดึงเอง  : ตั้งค่า Spreadsheet ID + ชื่อชีต + ชื่อคอลัมน์ของระบบอื่น แล้วกด "ดึงข้อมูลตอนนี้" (หรือตั้ง Trigger รายวันให้ฟังก์ชัน syncExternalWasteDaily)
//  API KEY เก็บใน Script Properties (ไม่อยู่ในชีต) — สร้าง/เปลี่ยนได้ในหน้า "เชื่อมต่อระบบอื่น" (ผู้ดูแลระบบ)
// ============================================================================================
const SHEET_INTEG = 'IntegrationSettings';
const SHEET_EXT_WASTE = 'ExternalWasteDaily';
const SHEET_EXT_DEPTS = 'ExternalDepartments';
const EXT_WASTE_HEADERS = ['WasteKey', 'WasteDate', 'Source', 'FoodKg', 'FruitVegKg', 'MeatKg', 'OilL', 'Note', 'Origin', 'UpdatedAt'];
const EXT_DEPT_HEADERS = ['Name', 'Source', 'UpdatedAt'];
const INTEG_KEY_PROP = 'INTEGRATION_API_KEY';
const API_PROJECT_LABELS = { ftf: 'ทอดไม่ทิ้ง', kitchen: 'จากครัวสู่ท้องฟ้า', paperless: 'Paperless', cartridge: 'ตลับหมึกเก่าเราขอ', calendar: 'ปฏิทินเก่าเราขอ', share: 'คัดแยก แบ่งปัน', em: 'น้ำยา EM จากเศษอาหาร', compost: 'ปุ๋ยหมักจากขยะใบไม้' };
const INTEG_DEFAULTS = { ExposeProjects: 'ftf,kitchen,paperless,cartridge,calendar,share,em,compost', SourceSpreadsheetId: '', SourceSheetName: '', ColDate: 'วันที่', ColSource: 'หน่วยงาน',
  ColFood: 'เศษอาหาร (กก.)', ColFruitVeg: 'ผักและผลไม้ (กก.)', ColMeat: 'เนื้อสัตว์ (กก.)', ColOil: 'น้ำมันพืช (ลิตร)', LastPullAt: '', LastPullRows: '', LastPullStatus: '', LastPushAt: '', LastPushRows: '' };

function getIntegSettings_() {
  const s = Object.assign({}, INTEG_DEFAULTS);
  readRowsIfSheet_(SHEET_INTEG).forEach(r => { if (Object.prototype.hasOwnProperty.call(INTEG_DEFAULTS, r.Key)) s[r.Key] = r.Value === undefined || r.Value === null ? '' : String(r.Value); });
  return s;
}
function setIntegSettings_(obj) {
  const sh = ensureSheet_(SHEET_INTEG, ['Key', 'Value']);
  const vals = sh.getDataRange().getValues();
  Object.keys(obj).forEach(k => {
    if (!Object.prototype.hasOwnProperty.call(INTEG_DEFAULTS, k)) return;
    let r = -1; for (let i = 1; i < vals.length; i++) if (vals[i][0] === k) { r = i + 1; break; }
    if (r === -1) { sh.appendRow([k, obj[k]]); vals.push([k, obj[k]]); } else { sh.getRange(r, 2).setValue(obj[k]); vals[r - 1][1] = obj[k]; }
  });
}
function integKey_() { return PropertiesService.getScriptProperties().getProperty(INTEG_KEY_PROP) || ''; }
function integKeyOk_(key) {
  const real = integKey_(), k = String(key || '');
  if (!real || real.length !== k.length) return false;
  let d = 0; for (let i = 0; i < real.length; i++) d |= real.charCodeAt(i) ^ k.charCodeAt(i);
  return d === 0;
}
// ============================================================================================
// หน้าเว็บ GitHub Pages (Portal): บันทึกการเข้าใช้งานด้วยรหัสเจ้าหน้าที่ 7 หลัก + ชื่อ-นามสกุล
// และสรุปว่าบุคคลนั้นร่วมโครงการอะไรบ้าง (แสดงเฉพาะชื่อโครงการ จำนวนครั้ง และยอดรวม — ไม่ส่งเบอร์โทร/ช่องทางติดต่อ)
// ============================================================================================
const SHEET_PORTAL_VISITS = 'PortalVisits';
const PORTAL_VISIT_HEADERS = ['VisitID', 'Timestamp', 'StaffCode', 'StaffName', 'Page', 'Device'];

/** ตัดคำนำหน้าและช่องว่างส่วนเกิน เพื่อเทียบชื่อ-นามสกุล */
function portalNormName_(v) {
  return String(v || '')
    .replace(/^\s*(นางสาว|นาง|นาย|น\.ส\.|ด\.ช\.|ด\.ญ\.|ดร\.|Mr\.?|Mrs\.?|Ms\.?|Miss)\s*/i, '')
    .replace(/\s+/g, ' ').trim().toLowerCase();
}

/** สรุปการร่วมโครงการของบุคคลจากรหัส (ถ้ามี) หรือชื่อ-นามสกุลที่ตรงกัน */
function portalParticipation_(code, name) {
  const nn = portalNormName_(name);
  const byName = v => !!nn && portalNormName_(v) === nn;
  const out = [];
  const add = (key, label, rows, qtyField, unit, dateField) => {
    if (!rows.length) return;
    const qty = qtyField ? round2_(rows.reduce((s, r) => s + (Number(r[qtyField]) || 0), 0)) : null;
    const last = rows.map(r => toIsoDate_(r[dateField])).filter(Boolean).sort().pop() || '';
    out.push({ key: key, label: label, count: rows.length, amount: qty, unit: unit || '', lastDate: last });
  };

  const ftf = readRowsIfSheet_(SHEET_FTF_LOG)
    .filter(r => (code && String(r.StaffCode || '').trim() === code) || byName(r.StaffName));
  add('frytofly', 'ทอดไม่ทิ้ง', ftf, 'Quantity', 'แกลลอน', 'Timestamp');

  const calDon = readRowsIfSheet_('CalendarDonations').filter(r => byName(r.DonorName));
  const calSub = readRowsIfSheet_('CalendarSubmissions').filter(r => byName(r.DonorName));
  add('calendar', 'ปฏิทินเก่าเราขอ', calDon.concat(calSub.map(r => ({ DonationDate: r.SubmitDate }))), 'WeightKg', 'กก.', 'DonationDate');

  const shr = readRowsIfSheet_('ShareSubmissions').filter(r => byName(r.SubmitterName));
  add('share', 'คัดแยก แบ่งปัน', shr, 'Quantity', 'รายการ', 'SubmitDate');

  const rq = readRowsIfSheet_(SHEET_PRODUCT_REQ).filter(r => r.Status !== PREQ_CANCEL && ((code && String(r.StaffCode || '').trim() === code) || byName(r.FullName)));
  add('em', 'ขอรับน้ำยา EM', rq.filter(r => r.Project === 'em'), 'Quantity', 'ลิตร', 'CreatedAt');
  add('compost', 'ขอรับปุ๋ยหมักใบไม้', rq.filter(r => r.Project === 'compost'), 'Quantity', 'กก.', 'CreatedAt');

  return out;
}

function portalCheckin_(params) {
  const code = String(params.code || '').trim();
  const name = String(params.name || '').replace(/\s+/g, ' ').trim().slice(0, 100);
  if (!/^\d{7}$/.test(code)) return { ok: false, error: 'รหัสเจ้าหน้าที่ต้องเป็นตัวเลข 7 หลัก' };
  if (name.length < 3) return { ok: false, error: 'กรุณากรอกชื่อ-นามสกุล' };

  ensureSheet_(SHEET_PORTAL_VISITS, PORTAL_VISIT_HEADERS);
  const before = readRowsIfSheet_(SHEET_PORTAL_VISITS).filter(r => String(r.StaffCode || '').trim() === code);
  if (String(params.log || '1') !== '0') {
    appendByHeaders_(SHEET_PORTAL_VISITS, {
      VisitID: newId_('PV'), Timestamp: new Date(), StaffCode: code, StaffName: name,
      Page: String(params.page || '').slice(0, 40), Device: String(params.device || '').slice(0, 20)
    });
  }
  const times = before.map(r => toIsoDate_(r.Timestamp)).filter(Boolean).sort();
  return {
    ok: true, staffCode: code, name: name,
    visits: { count: before.length + 1, first: times[0] || toIsoDate_(new Date()) },
    projects: portalParticipation_(code, name)
  };
}

/** ตัวเลขสรุปรายงานสิ่งแวดล้อมสำหรับหน้าแรก (ไม่มีข้อมูลส่วนบุคคล) */
function portalImpact_() {
  const s = getEnvironmentalImpactSummary();
  return {
    ok: true, generatedAt: new Date().toISOString(),
    overview: s.overview,
    paperless: { sheets: s.paperless.totalSheetsReduced, trees: s.paperless.treesSaved, co2Kg: s.paperless.co2SavedKg },
    fryToFly: { liters: s.fryToFly.totalLiters, participants: s.fryToFly.totalParticipants, co2Kg: s.fryToFly.co2SavedKg },
    cartridge: { contributed: s.cartridge.totalContributed, co2Kg: s.cartridge.co2AvoidedKg },
    share: { withdrawals: s.share.withdrawalCount, value: s.share.withdrawalValue, co2Kg: s.share.co2SavedKg },
    calendar: { kg: s.calendar.totalKg, trees: s.calendar.treesSaved, donors: s.calendar.donors }
  };
}

function jsonOut_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
function exposedProjects_() {
  const set = String(getIntegSettings_().ExposeProjects || '').split(',').map(x => x.trim()).filter(Boolean);
  return Object.keys(API_PROJECT_LABELS).filter(k => set.indexOf(k) !== -1);
}

/** รายชื่อหน่วยงานที่ใช้อยู่ในโครงการต่างๆ (เฉพาะชื่อ) */
function integDepartments_() {
  const m = {};
  [['paperless', SHEET_PL_DEPTS], ['cartridge', SHEET_CTG_DEPTS], ['share', SHEET_SHARE_DEPTS]].forEach(p => {
    readRowsIfSheet_(p[1]).forEach(d => {
      if (!d.DepartmentName || d.DepartmentID === PL_ALL_ID || !shareActive_(d.IsActive)) return;
      (m[d.DepartmentName] = m[d.DepartmentName] || { name: d.DepartmentName, projects: [] }).projects.push(p[0]);
    });
  });
  return Object.keys(m).sort((a, b) => a.localeCompare(b, 'th')).map(k => m[k]);
}

function handleApiGet_(params) {
  try {
    const api = String(params.api || '');
    if (api === 'ping') return jsonOut_({ ok: true, system: 'environment-project-tracker', time: new Date().toISOString(), keyConfigured: !!integKey_() });
    // หน้าเว็บ GitHub Pages (สาธารณะ ไม่ต้องใช้คีย์)
    if (api === 'portalImpact') return jsonOut_(portalImpact_());
    if (api === 'portalCheckin') return jsonOut_(portalCheckin_(params));
    if (!integKeyOk_(params.key)) return jsonOut_({ ok: false, error: 'unauthorized' });
    if (api === 'summary') {
      const exp = exposedProjects_();
      const want = String(params.projects || '').split(',').map(x => x.trim()).filter(Boolean);
      const list = want.length ? exp.filter(k => want.indexOf(k) !== -1) : exp;
      const year = Number(params.year) || new Date().getFullYear();
      const out = {};
      list.forEach(k => {
        const ps = getPeriodSummary(k, year);
        out[k] = { label: API_PROJECT_LABELS[k], columns: ps.columns.map(c => ({ key: c.key, label: c.label, unit: c.unit })), total: ps.total, yearTotal: ps.yearTotal, monthly: ps.monthly };
      });
      if (out.em) { const t = getEmOverview().totals; out.em.stock = { harvestedLiters: t.harvestedLiters, usedLiters: t.usedLiters, stockLiters: t.stockLiters, fermentingBarrels: t.fermentingBarrels, readyBarrels: t.readyBarrels, pipelineLiters: t.pipelineLiters }; }
      if (out.compost) { const tc = getCpOverview().totals; out.compost.stock = { harvestedKg: tc.harvestedKg, usedKg: tc.usedKg, stockKg: tc.stockKg, fermentingPiles: tc.fermentingPiles, readyPiles: tc.readyPiles, pipelineKg: tc.pipelineKg }; }
      return jsonOut_({ ok: true, generatedAt: new Date().toISOString(), year: year, projects: out });
    }
    if (api === 'departments') return jsonOut_({ ok: true, generatedAt: new Date().toISOString(), departments: integDepartments_() });
    if (api === 'emDaily') {
      const days = Math.min(366, Math.max(1, Number(params.days) || 60)), todayN = emToday_();
      const bs = readRowsIfSheet_(SHEET_EM_BATCHES).filter(b => b.BatchID && (b.Status || EM_STATUS_FERMENT) !== EM_STATUS_CANCEL);
      const rows = [];
      for (let i = days - 1; i >= 0; i--) {
        const n = todayN - i, L = bs.filter(b => emDayNum_(b.StartDate) === n);
        rows.push({ date: emIso_(n), foodKg: round2_(L.reduce((s, b) => s + (Number(b.FoodKg) || 0), 0)), fruitVegKg: round2_(L.reduce((s, b) => s + (Number(b.FruitVegKg) || 0), 0)),
          meatKg: round2_(L.reduce((s, b) => s + (Number(b.MeatKg) || 0), 0)), barrels: L.reduce((s, b) => s + (Number(b.BarrelCount) || 0), 0), batches: L.length });
      }
      return jsonOut_({ ok: true, generatedAt: new Date().toISOString(), days: days, daily: rows });
    }
    return jsonOut_({ ok: false, error: 'unknown api' });
  } catch (err) { return jsonOut_({ ok: false, error: String(err && err.message || err) }); }
}

function doPost(e) {
  try {
    let body = {};
    try { body = JSON.parse((e && e.postData && e.postData.contents) || '{}'); } catch (x) { return jsonOut_({ ok: false, error: 'invalid json' }); }
    const key = body.key || (e && e.parameter && e.parameter.key);
    if (!integKeyOk_(key)) return jsonOut_({ ok: false, error: 'unauthorized' });
    const type = String(body.type || '');
    if (type === 'wasteDaily') {
      const r = upsertExternalWaste_(body.rows, 'push', String(body.source || '').trim());
      setIntegSettings_({ LastPushAt: new Date().toISOString(), LastPushRows: String(r.upserted) });
      return jsonOut_(Object.assign({ ok: true }, r));
    }
    if (type === 'departments') {
      const r = upsertExternalDepartments_(body.names, String(body.source || '').trim());
      return jsonOut_(Object.assign({ ok: true }, r));
    }
    return jsonOut_({ ok: false, error: 'unknown type' });
  } catch (err) { return jsonOut_({ ok: false, error: String(err && err.message || err) }); }
}

/** วันที่จากระบบอื่น: Date / yyyy-mm-dd / d/m/yyyy (รองรับ พ.ศ.) → ISO */
function parseExtDate_(v) {
  if (v instanceof Date && !isNaN(v)) return toIsoDate_(v);
  const s = String(v === null || v === undefined ? '' : v).trim();
  let m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
  let y, mo, d;
  if (m) { y = +m[1]; mo = +m[2]; d = +m[3]; }
  else if ((m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/))) { d = +m[1]; mo = +m[2]; y = +m[3]; }
  else return '';
  if (y > 2400) y -= 543;
  if (y < 2000 || y > 2100 || mo < 1 || mo > 12 || d < 1 || d > 31) return '';
  return y + '-' + ('0' + mo).slice(-2) + '-' + ('0' + d).slice(-2);
}
function extNum_(v) {
  const n = Number(String(v === null || v === undefined ? '' : v).replace(/,/g, '').trim());
  return isNaN(n) || n < 0 ? 0 : Math.min(n, 1000000);
}

/** เพิ่ม/แก้ไข (upsert) ข้อมูลรายวันจากระบบอื่น — คีย์ = วันที่ + แหล่งที่มา/หน่วยงาน */
function upsertExternalWaste_(rows, origin, defaultSource) {
  if (!Array.isArray(rows)) throw new Error('rows ต้องเป็นอาร์เรย์');
  if (rows.length > 2000) throw new Error('ส่งได้ครั้งละไม่เกิน 2000 แถว');
  return withShareLock_(() => {
    const sh = ensureSheet_(SHEET_EXT_WASTE, EXT_WASTE_HEADERS);
    const vals = sh.getDataRange().getValues(), idx = {};
    for (let i = 1; i < vals.length; i++) idx[vals[i][0]] = i + 1;
    let up = 0, skipped = 0; const now = new Date(), pend = {};
    rows.forEach(r => {
      r = r || {};
      const date = parseExtDate_(r.date !== undefined ? r.date : r.Date);
      if (!date) { skipped++; return; }
      const src = String(r.source !== undefined ? r.source : (r.department !== undefined ? r.department : (defaultSource || 'ไม่ระบุ'))).trim().slice(0, 120) || 'ไม่ระบุ';
      const food = extNum_(r.foodKg), fv = extNum_(r.fruitVegKg), meat = extNum_(r.meatKg), oil = extNum_(r.oilL);
      const k = date + '|' + src, row = [k, date, src, food, fv, meat, oil, String(r.note || '').slice(0, 200), origin, now];
      if (idx[k]) { sh.getRange(idx[k], 1, 1, row.length).setValues([row]); }
      else if (pend[k] !== undefined) { pend[k] = row; }
      else { pend[k] = row; }
      up++;
    });
    Object.keys(pend).forEach(k => { if (!idx[k]) { sh.appendRow(pend[k]); idx[k] = -1; } });
    invalidateReadCache_();
    return { upserted: up, skipped: skipped };
  });
}

function upsertExternalDepartments_(names, source) {
  if (!Array.isArray(names)) throw new Error('names ต้องเป็นอาร์เรย์');
  return withShareLock_(() => {
    const sh = ensureSheet_(SHEET_EXT_DEPTS, EXT_DEPT_HEADERS);
    const have = {}; sheetToObjects_(SHEET_EXT_DEPTS).forEach(r => { have[shareKey_(r.Name)] = true; });
    let added = 0;
    names.slice(0, 1000).forEach(n => { const nm = String(n || '').trim().slice(0, 150); if (!nm || have[shareKey_(nm)]) return; have[shareKey_(nm)] = true; sh.appendRow([nm, source || '', new Date()]); added++; });
    return { added: added, total: Object.keys(have).length };
  });
}

/** ดึงข้อมูลจากชีตของระบบอื่นตามที่ตั้งค่าไว้ */
function pullExternalWasteCore_() {
  const s = getIntegSettings_();
  if (!s.SourceSpreadsheetId || !s.SourceSheetName) throw new Error('ยังไม่ได้ตั้งค่า Spreadsheet ID และชื่อชีตของระบบต้นทาง');
  let ss; try { ss = SpreadsheetApp.openById(s.SourceSpreadsheetId); } catch (e) { throw new Error('เปิดสเปรดชีตต้นทางไม่ได้ — ตรวจ Spreadsheet ID และสิทธิ์การเข้าถึงของบัญชีเจ้าของระบบนี้'); }
  const sh = ss.getSheetByName(s.SourceSheetName);
  if (!sh) throw new Error('ไม่พบชีต "' + s.SourceSheetName + '" ในสเปรดชีตต้นทาง');
  const vals = sh.getDataRange().getValues();
  if (vals.length < 2) return { upserted: 0, skipped: 0 };
  const h = vals[0].map(x => String(x).trim()), col = n => n ? h.indexOf(String(n).trim()) : -1;
  const cD = col(s.ColDate);
  if (cD === -1) throw new Error('ไม่พบคอลัมน์วันที่ "' + s.ColDate + '" ในชีตต้นทาง (หัวคอลัมน์ที่พบ: ' + h.join(', ') + ')');
  const cS = col(s.ColSource), cF = col(s.ColFood), cV = col(s.ColFruitVeg), cM = col(s.ColMeat), cO = col(s.ColOil);
  if (cF === -1 && cV === -1 && cM === -1 && cO === -1) throw new Error('ไม่พบคอลัมน์ปริมาณใดเลย (เศษอาหาร/ผักผลไม้/เนื้อสัตว์/น้ำมัน) — ตรวจชื่อคอลัมน์ที่ตั้งค่า');
  // รวมหลายแถวของวันเดียวกัน+แหล่งเดียวกันก่อน
  const agg = {};
  for (let r = 1; r < vals.length; r++) {
    const row = vals[r], date = parseExtDate_(row[cD]);
    if (!date) continue;
    const src = cS === -1 ? 'ระบบต้นทาง' : (String(row[cS] || '').trim() || 'ไม่ระบุ');
    const k = date + '|' + src, a = agg[k] = agg[k] || { date: date, source: src, foodKg: 0, fruitVegKg: 0, meatKg: 0, oilL: 0 };
    a.foodKg += cF === -1 ? 0 : extNum_(row[cF]); a.fruitVegKg += cV === -1 ? 0 : extNum_(row[cV]); a.meatKg += cM === -1 ? 0 : extNum_(row[cM]); a.oilL += cO === -1 ? 0 : extNum_(row[cO]);
  }
  return upsertExternalWaste_(Object.keys(agg).map(k => agg[k]), 'pull', '');
}

function pullExternalWaste(departmentId, pin) {
  assertAuthorized_(departmentId, pin);
  try {
    const r = pullExternalWasteCore_();
    setIntegSettings_({ LastPullAt: new Date().toISOString(), LastPullRows: String(r.upserted), LastPullStatus: 'สำเร็จ' + (r.skipped ? ' (ข้าม ' + r.skipped + ' แถวที่วันที่ไม่ถูกต้อง)' : '') });
    return Object.assign({ ok: true }, r);
  } catch (e) {
    setIntegSettings_({ LastPullAt: new Date().toISOString(), LastPullStatus: 'ผิดพลาด: ' + (e && e.message || e) });
    throw e;
  }
}

/** สำหรับตั้ง Trigger รายวัน (Time-driven) — ดึงข้อมูลตามที่ตั้งค่าไว้ ไม่ต้องล็อกอิน (จำกัดไม่ให้ทำซ้ำภายใน 5 นาที) */
function syncExternalWasteDaily() {
  const s = getIntegSettings_();
  if (s.LastPullAt && Date.now() - new Date(s.LastPullAt).getTime() < 5 * 60 * 1000) return { ok: true, skipped: 'recent' };
  try {
    const r = pullExternalWasteCore_();
    setIntegSettings_({ LastPullAt: new Date().toISOString(), LastPullRows: String(r.upserted), LastPullStatus: 'สำเร็จ (อัตโนมัติ)' });
    return { ok: true, upserted: r.upserted };
  } catch (e) { setIntegSettings_({ LastPullAt: new Date().toISOString(), LastPullStatus: 'ผิดพลาด: ' + (e && e.message || e) }); return { ok: false, error: String(e && e.message || e) }; }
}

// ----------------------------- ผู้ดูแลระบบ: หน้าตั้งค่า "เชื่อมต่อระบบอื่น" -----------------------------
function getIntegrationInfo(departmentId, pin) {
  assertAuthorized_(departmentId, pin);
  const s = getIntegSettings_(), ext = readRowsIfSheet_(SHEET_EXT_WASTE).filter(r => r.WasteKey), depts = readRowsIfSheet_(SHEET_EXT_DEPTS).filter(r => r.Name);
  const have = {}; ['cartridge', 'share'].forEach(k => { have[k] = {}; readRowsIfSheet_(k === 'cartridge' ? SHEET_CTG_DEPTS : SHEET_SHARE_DEPTS).forEach(d => { have[k][shareKey_(d.DepartmentName)] = true; }); });
  return {
    apiKey: integKey_(), webAppUrl: ScriptApp.getService().getUrl() || '',
    projects: Object.keys(API_PROJECT_LABELS).map(k => ({ key: k, label: API_PROJECT_LABELS[k], exposed: exposedProjects_().indexOf(k) !== -1 })),
    source: { id: s.SourceSpreadsheetId, sheet: s.SourceSheetName, colDate: s.ColDate, colSource: s.ColSource, colFood: s.ColFood, colFruitVeg: s.ColFruitVeg, colMeat: s.ColMeat, colOil: s.ColOil },
    status: { lastPullAt: s.LastPullAt, lastPullRows: s.LastPullRows, lastPullStatus: s.LastPullStatus, lastPushAt: s.LastPushAt, lastPushRows: s.LastPushRows },
    waste: { rows: ext.length, sources: Array.from(new Set(ext.map(r => r.Source))).length, from: ext.length ? ext.map(r => toIsoDate_(r.WasteDate)).sort()[0] : null, to: ext.length ? ext.map(r => toIsoDate_(r.WasteDate)).sort().pop() : null },
    departments: depts.map(d => ({ name: d.Name, source: d.Source || '', inCartridge: !!have.cartridge[shareKey_(d.Name)], inShare: !!have.share[shareKey_(d.Name)] }))
  };
}
function generateIntegrationKey(departmentId, pin) {
  assertAuthorized_(departmentId, pin);
  const key = (Utilities.getUuid() + Utilities.getUuid()).replace(/-/g, '');
  PropertiesService.getScriptProperties().setProperty(INTEG_KEY_PROP, key);
  return { ok: true, apiKey: key };
}
function saveIntegrationSettings(departmentId, pin, data) {
  assertAuthorized_(departmentId, pin);
  data = data || {};
  const o = {};
  if (Array.isArray(data.expose)) o.ExposeProjects = data.expose.filter(k => API_PROJECT_LABELS[k]).join(',');
  if (data.source) {
    const s = data.source, t = v => String(v === undefined || v === null ? '' : v).trim().slice(0, 200);
    let id = t(s.id); const m = id.match(/\/d\/([a-zA-Z0-9_-]+)/); if (m) id = m[1]; // วางลิงก์ทั้งเส้นได้
    o.SourceSpreadsheetId = id; o.SourceSheetName = t(s.sheet); o.ColDate = t(s.colDate); o.ColSource = t(s.colSource); o.ColFood = t(s.colFood); o.ColFruitVeg = t(s.colFruitVeg); o.ColMeat = t(s.colMeat); o.ColOil = t(s.colOil);
  }
  setIntegSettings_(o);
  return { ok: true };
}
/** นำรายชื่อหน่วยงานที่ได้รับจากระบบอื่น เข้าไปในรายชื่อของโครงการ (ตลับหมึก / คัดแยก แบ่งปัน) เฉพาะชื่อที่ยังไม่มี */
function importExternalDepartments(departmentId, pin, projectKeys) {
  assertAuthorized_(departmentId, pin);
  const keys = (projectKeys || []).filter(k => k === 'cartridge' || k === 'share');
  if (!keys.length) throw new Error('กรุณาเลือกโครงการที่จะนำเข้า');
  const ext = readRowsIfSheet_(SHEET_EXT_DEPTS).filter(r => r.Name).map(r => String(r.Name).trim());
  const res = {};
  keys.forEach(k => {
    const sheet = projectDeptSheet_(k), sh = getSheet_(sheet), rows = sheetToObjects_(sheet);
    const have = {}; rows.forEach(d => { have[shareKey_(d.DepartmentName)] = true; });
    let sort = rows.reduce((m, d) => Math.max(m, Number(d.SortOrder) || 0), 0), added = 0;
    ext.forEach(n => { if (have[shareKey_(n)]) return; have[shareKey_(n)] = true; sort++; appendByHeaders_(sheet, { DepartmentID: newId_(k === 'cartridge' ? 'CTD' : 'SHD'), DepartmentName: n, SortOrder: sort, IsActive: true }); added++; });
    res[k] = added;
  });
  return { ok: true, added: res };
}

// ----------------------------- แสดงผลในแท็บน้ำยา EM: วัตถุดิบ (เศษอาหาร/ผักผลไม้/เนื้อสัตว์/น้ำมัน) ที่ได้รับจากระบบอื่นเทียบกับที่นำมาหมัก -----------------------------
function getEmExternalWaste(days) {
  const D = Math.min(180, Math.max(7, Number(days) || 30)), todayN = emToday_();
  const rows = readRowsIfSheet_(SHEET_EXT_WASTE).filter(r => r.WasteKey);
  const bs = readRowsIfSheet_(SHEET_EM_BATCHES).filter(b => b.BatchID && (b.Status || EM_STATUS_FERMENT) !== EM_STATUS_CANCEL);
  const daily = [];
  let tf = 0, tv = 0, tm = 0, to = 0, tb = 0;
  for (let i = D - 1; i >= 0; i--) {
    const n = todayN - i, iso = emIso_(n);
    const L = rows.filter(r => toIsoDate_(r.WasteDate) === iso), B = bs.filter(b => emDayNum_(b.StartDate) === n);
    const f = L.reduce((s, r) => s + (Number(r.FoodKg) || 0), 0), v = L.reduce((s, r) => s + (Number(r.FruitVegKg) || 0), 0), m = L.reduce((s, r) => s + (Number(r.MeatKg) || 0), 0), o = L.reduce((s, r) => s + (Number(r.OilL) || 0), 0);
    const bk = B.reduce((s, b) => s + emInputKg_(b), 0);
    tf += f; tv += v; tm += m; to += o; tb += bk;
    daily.push({ date: iso, foodKg: round2_(f), fruitVegKg: round2_(v), meatKg: round2_(m), oilL: round2_(o), batchKg: round2_(bk), sources: L.length });
  }
  const bySource = {};
  rows.filter(r => { const n = emDayNum_(r.WasteDate); return n !== null && n > todayN - D && n <= todayN; }).forEach(r => {
    const a = bySource[r.Source] = bySource[r.Source] || { name: r.Source, kg: 0, oilL: 0 };
    a.kg += (Number(r.FoodKg) || 0) + (Number(r.FruitVegKg) || 0) + (Number(r.MeatKg) || 0); a.oilL += Number(r.OilL) || 0;
  });
  const wasteKg = tf + tv + tm;
  return {
    days: D, hasData: rows.length > 0, daily: daily,
    totals: { foodKg: round2_(tf), fruitVegKg: round2_(tv), meatKg: round2_(tm), wasteKg: round2_(wasteKg), oilL: round2_(to), batchKg: round2_(tb), utilizationPct: wasteKg > 0 ? round2_(Math.min(999, tb / wasteKg * 100)) : null },
    bySource: Object.keys(bySource).map(k => ({ name: k, kg: round2_(bySource[k].kg), oilL: round2_(bySource[k].oilL) })).sort((a, b) => b.kg - a.kg).slice(0, 10),
    lastDate: rows.length ? rows.map(r => toIsoDate_(r.WasteDate)).sort().pop() : null
  };
}

// ==== COMPOST-BEGIN (สร้างจากโมดูล EM — โครงการปุ๋ยหมักจากขยะใบไม้) ====
// ============================================================================================
// โครงการ "ปุ๋ยหมักจากขยะใบไม้"
//  INPUT   : ใบไม้แห้ง / ใบไม้สดและเศษหญ้า / กิ่งไม้และเศษไม้ (กก.) + สารเร่ง/มูลสัตว์ + น้ำ
//  PROCESS : หมักเป็นรอบ (หลายกองต่อรอบ) ติดตามอายุการหมัก 7 / 15 / 45 วัน และตรวจทุก 7 วัน
//  OUTPUT  : เมื่อครบกำหนด (ReadyDays) เก็บเป็นปุ๋ยหมักกี่กก.
//  USE     : บันทึกการนำไปใช้ (สถานที่ ผู้ใช้ ปริมาณ) → คงเหลือ = ผลิตได้ − ใช้ไป
//  ชีต: CompostBatches (รอบการหมัก) / CompostChecks (ผลตรวจทุก 7 วัน) / CompostUsage (การนำไปใช้) / CompostSettings (ค่าตั้งต้น ปรับได้ในชีตโดยไม่ต้องแก้โค้ด)
//  การบันทึก/แก้ไข: ผู้ดูแลระบบ (IsAdmin) เท่านั้น • การดูรายงาน: ทุกคน
// ============================================================================================
const SHEET_CP_BATCHES  = 'CompostBatches';
const SHEET_CP_CHECKS   = 'CompostChecks';
const SHEET_CP_USAGE    = 'CompostUsage';
const SHEET_CP_SETTINGS = 'CompostSettings';
const SHEET_CP_LOGS     = 'CompostDailyLogs';
const CP_LOG_HEADERS = ['LogID', 'LogDate', 'BatchID', 'LogType', 'DryKg', 'FreshKg', 'BranchKg', 'Piles', 'Location', 'TempC', 'Smell', 'Appearance', 'Action', 'Note', 'Operator', 'CreatedAt', 'VerifyStatus', 'VerifiedBy', 'VerifiedAt'];
const CP_LOG_NEW = 'เริ่มหมักกองใหม่', CP_LOG_ADD = 'เติมวัตถุดิบเพิ่ม', CP_LOG_CHECK = 'ตรวจ/ดูแลประจำวัน';
const CP_LOG_TYPES = [CP_LOG_NEW, CP_LOG_ADD, CP_LOG_CHECK];
const CP_BATCH_HEADERS = ['BatchID', 'BatchNo', 'StartDate', 'DryKg', 'FreshKg', 'BranchKg', 'ActivatorKg', 'WaterL', 'PileCount',
                          'ExpectedKg', 'ReadyDays', 'Location', 'Note', 'Status', 'HarvestDate', 'YieldKg', 'HarvestNote', 'RecordedBy', 'CreatedAt', 'UpdatedAt'];
const CP_CHECK_HEADERS = ['CheckID', 'BatchID', 'CheckDate', 'Week', 'DayNo', 'TempC', 'Smell', 'Appearance', 'Action', 'Note', 'RecordedBy', 'CreatedAt'];
const CP_USAGE_HEADERS = ['UsageID', 'BatchID', 'UsageDate', 'Kg', 'Place', 'UsedBy', 'Purpose', 'Note', 'RecordedBy', 'CreatedAt', 'VerifyStatus', 'VerifiedBy', 'VerifiedAt'];
const CP_STATUS_FERMENT = 'กำลังหมัก';
const CP_STATUS_DONE    = 'เก็บเป็นปุ๋ยหมักแล้ว';
const CP_STATUS_CANCEL  = 'ยกเลิก';
const CP_DEFAULTS = { Stage1Days: 7, Stage2Days: 15, ReadyDays: 45, CheckIntervalDays: 7, YieldPercent: 40, SoonDays: 7 };

function ensureCpSheets_() {
  ensureSheet_(SHEET_CP_BATCHES, CP_BATCH_HEADERS);
  ensureSheet_(SHEET_CP_CHECKS, CP_CHECK_HEADERS);
  ensureSheet_(SHEET_CP_USAGE, CP_USAGE_HEADERS);
  ensureVerifyColumns_(SHEET_CP_USAGE);
  ensureSheet_(SHEET_CP_LOGS, CP_LOG_HEADERS);
  ensureCpSettingsSeeded_();
}

function ensureCpSettingsSeeded_() {
  const sh = ensureSheet_(SHEET_CP_SETTINGS, ['Key', 'Value']);
  const have = sheetToObjects_(SHEET_CP_SETTINGS).map(r => r.Key);
  Object.keys(CP_DEFAULTS).forEach(k => { if (have.indexOf(k) === -1) sh.appendRow([k, CP_DEFAULTS[k]]); });
}

/** ค่าตั้งต้นของโครงการ (ชีต CompostSettings): ช่วงวันที่สำคัญของการหมัก ขนาดกอง และ % ปุ๋ยหมักที่คาดว่าจะได้ */
function getCpSettings_() {
  const s = Object.assign({}, CP_DEFAULTS);
  readRowsIfSheet_(SHEET_CP_SETTINGS).forEach(r => {
    if (Object.prototype.hasOwnProperty.call(CP_DEFAULTS, r.Key) && r.Value !== '' && !isNaN(Number(r.Value)) && Number(r.Value) > 0) s[r.Key] = Number(r.Value);
  });
  if (s.Stage2Days <= s.Stage1Days) s.Stage2Days = s.Stage1Days + 1;
  if (s.ReadyDays <= s.Stage2Days) s.ReadyDays = s.Stage2Days + 1;
  return s;
}

/** วันที่ (ปฏิทินท้องถิ่น) → เลขวันนับจาก epoch (ตัดเวลาทิ้ง) */
function cpDayNum_(v) {
  if (v === '' || v === null || v === undefined) return null;
  const s = (v instanceof Date || typeof v === 'number') ? toIsoDate_(v) : String(v);
  const m = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (!m) return null;
  return Math.floor(Date.UTC(Number(m[1]), Number(m[2]) - 1, Number(m[3])) / 86400000);
}
function cpIso_(n) { return new Date(n * 86400000).toISOString().slice(0, 10); }
function cpToday_() { return cpDayNum_(new Date()); }
function cpInputKg_(b) { return round2_((Number(b.DryKg) || 0) + (Number(b.FreshKg) || 0) + (Number(b.BranchKg) || 0)); }

/** ข้อมูลรอบการหมัก 1 รอบ ณ วันที่ asOfN: อายุ ระยะ ความคืบหน้า ผลตรวจรายสัปดาห์ */
function cpLogView_(l) {
  const kg = (Number(l.DryKg) || 0) + (Number(l.FreshKg) || 0) + (Number(l.BranchKg) || 0);
  return { LogID: l.LogID, LogDate: toIsoDate_(l.LogDate), BatchID: l.BatchID || '', LogType: l.LogType || '', Operator: l.Operator || '', Kg: round2_(kg), DryKg: Number(l.DryKg) || 0, FreshKg: Number(l.FreshKg) || 0, BranchKg: Number(l.BranchKg) || 0,
           Piles: Number(l.Piles) || 0, Location: l.Location || '', TempC: l.TempC === '' || l.TempC === null || l.TempC === undefined ? null : Number(l.TempC), Smell: l.Smell || '', Appearance: l.Appearance || '', Action: l.Action || '', Note: l.Note || '' };
}
function cpBatchView_(b, st, asOfN, checks, usedL) {
  const startN = cpDayNum_(b.StartDate);
  const ready = Number(b.ReadyDays) > 0 ? Number(b.ReadyDays) : st.ReadyDays;
  const piles = Number(b.PileCount) || 0;
  const status = b.Status || CP_STATUS_FERMENT;
  const harvestN = cpDayNum_(b.HarvestDate);
  const harvested = status === CP_STATUS_DONE && (harvestN === null || harvestN <= asOfN);
  const cancelled = status === CP_STATUS_CANCEL;
  const age = startN === null ? 0 : Math.max(0, asOfN - startN);
  const expected = Number(b.ExpectedKg) > 0 ? Number(b.ExpectedKg) : round2_(cpInputKg_(b) * st.YieldPercent / 100);
  let stage, stageLabel;
  if (cancelled) { stage = 'cancel'; stageLabel = 'ยกเลิก'; }
  else if (harvested) { stage = 'done'; stageLabel = 'เก็บเป็นปุ๋ยหมักแล้ว'; }
  else if (age >= ready) { stage = 'ready'; stageLabel = 'ครบกำหนด — ควรเก็บเป็นปุ๋ยหมัก'; }
  else if (age >= st.Stage2Days) { stage = 'late'; stageLabel = 'หมักช่วง ' + st.Stage2Days + '–' + (ready - 1) + ' วัน'; }
  else if (age >= st.Stage1Days) { stage = 'mid'; stageLabel = 'หมักช่วง ' + st.Stage1Days + '–' + (st.Stage2Days - 1) + ' วัน'; }
  else { stage = 'early'; stageLabel = 'หมักช่วงต้น (ไม่ถึง ' + st.Stage1Days + ' วัน)'; }
  const fermenting = stage === 'early' || stage === 'mid' || stage === 'late' || stage === 'ready';
  const progress = harvested ? 100 : (cancelled ? 0 : Math.min(100, Math.round(age / ready * 100)));

  // เช็กพอยต์ตรวจทุก N วัน (7, 14, 21, …) ก่อนครบกำหนด
  const interval = st.CheckIntervalDays;
  const myChecks = (checks || []).filter(c => c.BatchID === b.BatchID).map(c => ({
    CheckID: c.CheckID, CheckDate: toIsoDate_(c.CheckDate), Week: Number(c.Week) || 0, DayNo: Number(c.DayNo) || 0,
    TempC: c.TempC === '' || c.TempC === null || c.TempC === undefined ? null : Number(c.TempC), Smell: c.Smell || '', Appearance: c.Appearance || '',
    Action: c.Action || '', Note: c.Note || '', RecordedBy: c.RecordedBy || ''
  })).filter(c => { const cn = cpDayNum_(c.CheckDate); return cn === null || cn <= asOfN; }).sort((a, b2) => String(a.CheckDate).localeCompare(String(b2.CheckDate)));
  const weeks = [];
  const nW = Math.max(1, Math.floor(ready / interval));
  for (let k = 1; k <= nW; k++) {
    const dueDay = interval * k, chk = myChecks.find(c => c.Week === k);
    let wst = 'upcoming';
    if (chk) wst = 'done';
    else if (cancelled || harvested && harvestN !== null && harvestN < (startN || 0) + dueDay) wst = 'skip';
    else if (age > dueDay) wst = 'overdue';
    else if (age === dueDay) wst = 'due';
    weeks.push({ week: k, dueDay: dueDay, dueDate: startN === null ? null : cpIso_(startN + dueDay), status: wst, check: chk || null });
  }
  const pending = fermenting ? weeks.filter(w => w.status === 'due' || w.status === 'overdue') : [];
  const nextWeek = fermenting ? weeks.find(w => w.status === 'upcoming') : null;
  // หมุดหมายถัดไปของการหมัก
  let nextMilestone = null;
  if (fermenting && age < ready) {
    const cand = [st.Stage1Days, st.Stage2Days, ready].filter(d => d > age);
    const d = cand[0];
    nextMilestone = { day: d, date: startN === null ? null : cpIso_(startN + d), inDays: d - age, label: d === ready ? 'ครบกำหนดเป็นปุ๋ยหมัก' : 'ครบ ' + d + ' วัน' };
  }
  const used = usedL || 0, yieldL = harvested ? (Number(b.YieldKg) || 0) : 0;
  return {
    BatchID: b.BatchID, BatchNo: Number(b.BatchNo) || 0, StartDate: toIsoDate_(b.StartDate), age: age, readyDays: ready,
    readyDate: startN === null ? null : cpIso_(startN + ready), daysToReady: ready - age,
    progress: progress, stage: stage, stageLabel: stageLabel, fermenting: fermenting,
    piles: piles, expectedKg: expected,
    dryKg: Number(b.DryKg) || 0, freshKg: Number(b.FreshKg) || 0, branchKg: Number(b.BranchKg) || 0, inputKg: cpInputKg_(b),
    activatorKg: Number(b.ActivatorKg) || 0, waterL: Number(b.WaterL) || 0,
    location: b.Location || '', note: b.Note || '', status: status,
    harvestDate: b.HarvestDate ? toIsoDate_(b.HarvestDate) : null, yieldKg: yieldL, harvestNote: b.HarvestNote || '',
    usedKg: round2_(used), remainingKg: round2_(yieldL - used),
    checks: myChecks, weeks: weeks, pendingChecks: pending.length, pendingWeeks: pending.map(w => w.week), nextMilestone: nextMilestone,
    nextCheckDate: nextWeek ? nextWeek.dueDate : null, nextCheckWeek: nextWeek ? nextWeek.week : null,
    recordedBy: b.RecordedBy || ''
  };
}

/** ภาพรวมทั้งหมดของโครงการ (เรียกครั้งเดียวต่อการเปิดแท็บ) */
function getCpOverview() {
  const st = getCpSettings_();
  const todayN = cpToday_();
  const rawB = readRowsIfSheet_(SHEET_CP_BATCHES).filter(b => b.BatchID);
  const rawC = readRowsIfSheet_(SHEET_CP_CHECKS).filter(c => c.CheckID);
  const allU = readRowsIfSheet_(SHEET_CP_USAGE).filter(u => u.UsageID);
  const rawU = allU.filter(u => isVerified_(u.VerifyStatus));
  const allL = readRowsIfSheet_(SHEET_CP_LOGS).filter(l => l.LogID);
  const pendL = allL.filter(l => String(l.VerifyStatus || '') === VERIFY_PENDING);
  const okL = allL.filter(l => isVerified_(l.VerifyStatus));
  const pendU = allU.filter(u => String(u.VerifyStatus || '') === VERIFY_PENDING);
  const usedBy = {};
  rawU.forEach(u => { if (u.BatchID) usedBy[u.BatchID] = (usedBy[u.BatchID] || 0) + (Number(u.Kg) || 0); });
  const batches = rawB.map(b => cpBatchView_(b, st, todayN, rawC, usedBy[b.BatchID] || 0))
    .sort((a, b) => String(b.StartDate).localeCompare(String(a.StartDate)) || b.BatchNo - a.BatchNo);
  const logsOf = {};
  okL.forEach(l => { const bid = l.BatchID; if (bid) (logsOf[bid] = logsOf[bid] || []).push(cpLogView_(l)); });
  batches.forEach(b => { b.logs = (logsOf[b.BatchID] || []).sort((x, y) => String(y.LogDate).localeCompare(String(x.LogDate))); b.logCount = b.logs.length; b.lastLogDate = b.logs.length ? b.logs[0].LogDate : null; });
  const live = batches.filter(b => b.stage !== 'cancel');
  const byNo = {}; batches.forEach(b => { byNo[b.BatchID] = b.BatchNo; });

  const sum = (list, f) => list.reduce((s, i) => s + (Number(f(i)) || 0), 0);
  const stageDef = [
    { key: 'early', label: 'ไม่ถึง ' + st.Stage1Days + ' วัน', range: '0–' + (st.Stage1Days - 1) + ' วัน' },
    { key: 'mid', label: st.Stage1Days + '–' + (st.Stage2Days - 1) + ' วัน', range: st.Stage1Days + '–' + (st.Stage2Days - 1) + ' วัน' },
    { key: 'late', label: st.Stage2Days + '–' + (st.ReadyDays - 1) + ' วัน', range: st.Stage2Days + '–' + (st.ReadyDays - 1) + ' วัน' },
    { key: 'ready', label: 'ครบ ' + st.ReadyDays + ' วัน (พร้อมเก็บ)', range: '≥ ' + st.ReadyDays + ' วัน' }
  ];
  const byStage = stageDef.map(d => {
    const L = live.filter(b => b.stage === d.key);
    return { key: d.key, label: d.label, range: d.range, batches: L.length, piles: sum(L, b => b.piles), expectedKg: round2_(sum(L, b => b.expectedKg)), inputKg: round2_(sum(L, b => b.inputKg)) };
  });
  // กองที่ "ครบ" 7 / 15 / 45 วันแล้ว (สะสม) และที่จะครบในอีก SoonDays วัน
  const milestones = [st.Stage1Days, st.Stage2Days, st.ReadyDays].map(M => {
    const reached = live.filter(b => (b.stage === 'done' && b.harvestDate && cpDayNum_(b.harvestDate) - cpDayNum_(b.StartDate) < M) ? false : b.age >= M || b.stage === 'done');
    const soon = live.filter(b => b.fermenting && b.age < M && M - b.age <= st.SoonDays);
    const waiting = live.filter(b => b.fermenting && b.age < M);
    return { days: M, reachedBatches: reached.length, reachedPiles: sum(reached, b => b.piles), reachedExpectedKg: round2_(sum(reached, b => b.expectedKg)),
             soonBatches: soon.length, soonPiles: sum(soon, b => b.piles), waitingPiles: sum(waiting, b => b.piles) };
  });
  // กองที่เริ่มหมักรายวัน 30 วันล่าสุด
  const daily = [];
  for (let i = 29; i >= 0; i--) {
    const n = todayN - i, L = live.filter(b => cpDayNum_(b.StartDate) === n);
    daily.push({ date: cpIso_(n), batches: L.length, piles: sum(L, b => b.piles), kg: round2_(sum(L, b => b.inputKg)) });
  }
  // คาดการณ์ปุ๋ยหมักที่จะพร้อมเก็บในแต่ละสัปดาห์ถัดไป (8 สัปดาห์)
  const forecast = [];
  for (let w = 0; w < 8; w++) {
    const from = todayN + w * 7, to = from + 6;
    const L = live.filter(b => b.fermenting && (() => { const rn = cpDayNum_(b.readyDate); return w === 0 ? rn <= to : (rn >= from && rn <= to); })());
    forecast.push({ from: cpIso_(from), to: cpIso_(to), batches: L.length, piles: sum(L, b => b.piles), expectedKg: round2_(sum(L, b => b.expectedKg)), overdue: w === 0 ? L.filter(b => b.stage === 'ready').length : 0 });
  }
  const harvestedL = round2_(sum(live.filter(b => b.stage === 'done'), b => b.yieldKg));
  const usedAll = round2_(sum(rawU, u => u.Kg));
  const usage = rawU.map(u => ({ UsageID: u.UsageID, BatchID: u.BatchID || '', BatchNo: u.BatchID ? (byNo[u.BatchID] || 0) : 0, UsageDate: toIsoDate_(u.UsageDate),
    Kg: Number(u.Kg) || 0, Place: u.Place || '', UsedBy: u.UsedBy || '', Purpose: u.Purpose || '', Note: u.Note || '', RecordedBy: u.RecordedBy || '' }))
    .sort((a, b) => String(b.UsageDate).localeCompare(String(a.UsageDate)));
  const rank = (field) => {
    const m = {}; usage.forEach(u => { const k = u[field] || '(ไม่ระบุ)'; const r = m[k] = m[k] || { name: k, amount: 0, times: 0, last: '' }; r.amount += u.Kg; r.times++; if (u.UsageDate > r.last) r.last = u.UsageDate; });
    return Object.keys(m).map(k => Object.assign(m[k], { amount: round2_(m[k].amount) })).sort((a, b) => b.amount - a.amount).map((r, i) => Object.assign(r, { rank: i + 1 }));
  };
  const ready = live.filter(b => b.stage === 'ready');
  const soonReady = live.filter(b => b.fermenting && b.stage !== 'ready' && b.daysToReady <= st.SoonDays).sort((a, b) => a.daysToReady - b.daysToReady);
  const pendingChecks = live.filter(b => b.pendingChecks > 0).map(b => ({ BatchID: b.BatchID, BatchNo: b.BatchNo, age: b.age, weeks: b.pendingWeeks }));
  const fermenting = live.filter(b => b.fermenting);
  return {
    settings: st, today: cpIso_(todayN), batches: batches, usage: usage,
    totals: {
      batches: live.length, piles: sum(live, b => b.piles),
      dryKg: round2_(sum(live, b => b.dryKg)), freshKg: round2_(sum(live, b => b.freshKg)), branchKg: round2_(sum(live, b => b.branchKg)),
      inputKg: round2_(sum(live, b => b.inputKg)), activatorKg: round2_(sum(live, b => b.activatorKg)), waterL: round2_(sum(live, b => b.waterL)),
      fermentingBatches: fermenting.length, fermentingPiles: sum(fermenting, b => b.piles), pipelineKg: round2_(sum(fermenting, b => b.expectedKg)),
      readyBatches: ready.length, readyPiles: sum(ready, b => b.piles), readyKg: round2_(sum(ready, b => b.expectedKg)),
      harvestedBatches: live.filter(b => b.stage === 'done').length, harvestedKg: harvestedL, usedKg: usedAll, stockKg: round2_(harvestedL - usedAll),
      pendingLogs: pendL.length, dailyLogs: okL.length,
      pendingUsage: pendU.length, pendingUsageKg: round2_(sum(pendU, u => u.Kg)),
      usageTimes: usage.length, usagePlaces: new Set(usage.map(u => u.Place).filter(Boolean)).size, usageUsers: new Set(usage.map(u => u.UsedBy).filter(Boolean)).size,
      pendingChecks: pendingChecks.reduce((s, b) => s + b.weeks.length, 0)
    },
    byStage: byStage, milestones: milestones, daily: daily, forecast: forecast,
    readyList: ready.map(b => b.BatchID), soonList: soonReady.map(b => b.BatchID), pendingChecks: pendingChecks,
    byPlace: rank('Place'), byUser: rank('UsedBy'),
    recentLogs: okL.map(l => Object.assign(cpLogView_(l), { BatchNo: l.BatchID ? (byNo[l.BatchID] || 0) : 0 })).sort((a, b) => String(b.LogDate).localeCompare(String(a.LogDate))).slice(0, 40),
    stockByBatch: live.filter(b => b.stage === 'done').map(b => ({ BatchID: b.BatchID, BatchNo: b.BatchNo, harvestDate: b.harvestDate, yieldKg: b.yieldKg, usedKg: b.usedKg, remainingKg: b.remainingKg }))
      .sort((a, b) => String(a.harvestDate).localeCompare(String(b.harvestDate)))
  };
}

/**
 * รายงานความคืบหน้าทุก 7 วัน: ช่วงละ 7 วัน นับถอยหลังจากวันที่สิ้นสุด (endDate, ค่าเริ่มต้น = วันนี้) ย้อนหลัง `periods` ช่วง
 * แต่ละช่วงสรุป: รอบ/กอง/กก.ที่เริ่มหมัก • กองที่ครบ 7/15/ครบกำหนด • ปุ๋ยหมักที่เก็บได้ • ที่ใช้ไป • คงเหลือ ณ สิ้นช่วง • การตรวจที่ถึงกำหนด/บันทึกแล้ว
 * และ "สถานะแต่ละรอบ ณ วันสิ้นสุด" สำหรับช่วงล่าสุด
 */
function getCpWeeklyReport(endDate, periods) {
  const st = getCpSettings_();
  let endN = cpDayNum_(endDate); if (endN === null) endN = cpToday_();
  const P = Math.min(26, Math.max(1, Number(periods) || 8));
  const rawB = readRowsIfSheet_(SHEET_CP_BATCHES).filter(b => b.BatchID && (b.Status || CP_STATUS_FERMENT) !== CP_STATUS_CANCEL);
  const rawC = readRowsIfSheet_(SHEET_CP_CHECKS).filter(c => c.CheckID);
  const rawU = readRowsIfSheet_(SHEET_CP_USAGE).filter(u => u.UsageID && isVerified_(u.VerifyStatus));
  const pilesOf = b => Number(b.PileCount) || 0;
  const sum = (list, f) => list.reduce((s, i) => s + (Number(f(i)) || 0), 0);
  const stockAt = n => round2_(sum(rawB.filter(b => b.Status === CP_STATUS_DONE && cpDayNum_(b.HarvestDate) !== null && cpDayNum_(b.HarvestDate) <= n), b => b.YieldKg)
                             - sum(rawU.filter(u => cpDayNum_(u.UsageDate) !== null && cpDayNum_(u.UsageDate) <= n), u => u.Kg));
  const interval = st.CheckIntervalDays;
  const series = [];
  for (let i = 0; i < P; i++) {
    const to = endN - 7 * (P - 1 - i), from = to - 6;
    const inP = n => n !== null && n >= from && n <= to;
    const started = rawB.filter(b => inP(cpDayNum_(b.StartDate)));
    const har = rawB.filter(b => b.Status === CP_STATUS_DONE && inP(cpDayNum_(b.HarvestDate)));
    const reach = M => rawB.filter(b => { const sn = cpDayNum_(b.StartDate); return sn !== null && inP(sn + M); });
    const ferment = rawB.filter(b => { const sn = cpDayNum_(b.StartDate); const hn = b.Status === CP_STATUS_DONE ? cpDayNum_(b.HarvestDate) : null; return sn !== null && sn <= to && (hn === null || hn > to); });
    const readyWait = ferment.filter(b => (Number(b.ReadyDays) > 0 ? Number(b.ReadyDays) : st.ReadyDays) <= to - cpDayNum_(b.StartDate));
    let due = 0, done = 0;
    ferment.forEach(b => {
      const sn = cpDayNum_(b.StartDate), rd = Number(b.ReadyDays) > 0 ? Number(b.ReadyDays) : st.ReadyDays, nW = Math.max(1, Math.floor(rd / interval));
      for (let k = 1; k <= nW; k++) {
        const dd = sn + interval * k;
        if (dd < from || dd > to) continue;
        due++;
        if (rawC.some(c => c.BatchID === b.BatchID && Number(c.Week) === k && cpDayNum_(c.CheckDate) !== null && cpDayNum_(c.CheckDate) <= to)) done++;
      }
    });
    series.push({
      from: cpIso_(from), to: cpIso_(to), startedBatches: started.length, startedPiles: sum(started, pilesOf), startedKg: round2_(sum(started, cpInputKg_)),
      reach1Piles: sum(reach(st.Stage1Days), pilesOf), reach2Piles: sum(reach(st.Stage2Days), pilesOf), reachReadyPiles: sum(reach(st.ReadyDays), pilesOf),
      harvestedBatches: har.length, harvestedKg: round2_(sum(har, b => b.YieldKg)),
      usedKg: round2_(sum(rawU.filter(u => inP(cpDayNum_(u.UsageDate))), u => u.Kg)), usageTimes: rawU.filter(u => inP(cpDayNum_(u.UsageDate))).length,
      stockEnd: stockAt(to), fermentingPiles: sum(ferment, pilesOf), fermentingBatches: ferment.length,
      readyWaitingBatches: readyWait.length, readyWaitingPiles: sum(readyWait, pilesOf),
      checksDue: due, checksDone: done
    });
  }
  const usedBy = {};
  rawU.forEach(u => { if (u.BatchID && cpDayNum_(u.UsageDate) !== null && cpDayNum_(u.UsageDate) <= endN) usedBy[u.BatchID] = (usedBy[u.BatchID] || 0) + (Number(u.Kg) || 0); });
  const snapshot = rawB.filter(b => cpDayNum_(b.StartDate) !== null && cpDayNum_(b.StartDate) <= endN)
    .map(b => cpBatchView_(b, st, endN, rawC, usedBy[b.BatchID] || 0))
    .filter(b => b.fermenting || (b.stage === 'done' && b.harvestDate && cpDayNum_(b.harvestDate) > endN - 7))
    .sort((a, b) => b.age - a.age);
  return { settings: st, endDate: cpIso_(endN), periods: P, series: series, current: series[series.length - 1], previous: series.length > 1 ? series[series.length - 2] : null, snapshot: snapshot };
}

// ----------------------------- WRITE (ผู้ดูแลระบบ) -----------------------------
function cpNum_(v, label, allowZero) {
  if (v === '' || v === null || v === undefined) return 0;
  const n = Number(v);
  if (isNaN(n) || n < 0) throw new Error(label + ' ต้องเป็นตัวเลขที่ไม่ติดลบ');
  return n;
}
function cpDate_(v, label) {
  const n = cpDayNum_(v);
  if (n === null) throw new Error('กรุณาระบุ' + label + 'ให้ถูกต้อง');
  return new Date(n * 86400000 + 12 * 3600000);
}
function cpSetFields_(sheetName, idField, id, fields) {
  const found = findRowById_(sheetName, idField, id);
  if (!found) throw new Error('ไม่พบรายการ');
  const sh = getSheet_(sheetName);
  Object.keys(fields).forEach(k => { const col = found.headers.indexOf(k); if (col !== -1) sh.getRange(found.rowIndex, col + 1).setValue(fields[k]); });
  return found;
}
function cpUsedForBatch_(batchId) {
  return readRowsIfSheet_(SHEET_CP_USAGE).filter(u => u.BatchID === batchId && isVerified_(u.VerifyStatus)).reduce((s, u) => s + (Number(u.Kg) || 0), 0);
}
function cpStockKg_() {
  const h = readRowsIfSheet_(SHEET_CP_BATCHES).filter(b => b.Status === CP_STATUS_DONE).reduce((s, b) => s + (Number(b.YieldKg) || 0), 0);
  const u = readRowsIfSheet_(SHEET_CP_USAGE).filter(x => isVerified_(x.VerifyStatus)).reduce((s, x) => s + (Number(x.Kg) || 0), 0);
  return round2_(h - u);
}

/** เพิ่มรอบการหมัก (ไม่ล็อก/ไม่เช็กสิทธิ์ — ผู้เรียกต้องจัดการเอง) */
function cpInsertBatch_(data, recordedBy) {
  data = data || {};
  const food = cpNum_(data.DryKg, 'ใบไม้แห้ง'), fv = cpNum_(data.FreshKg, 'ใบไม้สด/เศษหญ้า'), meat = cpNum_(data.BranchKg, 'กิ่งไม้/เศษไม้');
  if (food + fv + meat <= 0) throw new Error('กรุณาระบุปริมาณใบไม้แห้ง ใบไม้สด/เศษหญ้า หรือกิ่งไม้/เศษไม้ อย่างน้อย 1 ช่อง (กก.)');
  const piles = Math.round(cpNum_(data.PileCount, 'จำนวนกอง'));
  if (piles <= 0) throw new Error('กรุณาระบุจำนวนกองที่หมัก');
  const start = cpDate_(data.StartDate || new Date(), 'วันที่เริ่มหมัก');
  const st = getCpSettings_();
  ensureCpSheets_();
  const no = readRowsIfSheet_(SHEET_CP_BATCHES).reduce((m, b) => Math.max(m, Number(b.BatchNo) || 0), 0) + 1;
  const now = new Date(), id = newId_('CPB');
  appendByHeaders_(SHEET_CP_BATCHES, {
    BatchID: id, BatchNo: no, StartDate: start, DryKg: food, FreshKg: fv, BranchKg: meat, ActivatorKg: cpNum_(data.ActivatorKg, 'สารเร่ง/มูลสัตว์'), WaterL: cpNum_(data.WaterL, 'น้ำ'),
    PileCount: piles, ExpectedKg: cpNum_(data.ExpectedKg, 'ปริมาณปุ๋ยหมักที่คาดว่าจะได้') || round2_((food + fv + meat) * st.YieldPercent / 100),
    ReadyDays: cpNum_(data.ReadyDays, 'จำนวนวันที่ใช้หมัก') || '', Location: String(data.Location || '').trim(), Note: String(data.Note || '').trim(),
    Status: CP_STATUS_FERMENT, RecordedBy: recordedBy, CreatedAt: now, UpdatedAt: now
  });
  return { ok: true, batchId: id, batchNo: no };
}

function addCpBatch(departmentId, pin, data) {
  const auth = assertAuthorized_(departmentId, pin);
  data = data || {};
  // ตรวจรูปแบบข้อมูลก่อนล็อก (ให้ข้อความผิดพลาดชัดเจน)
  cpNum_(data.DryKg, 'ใบไม้แห้ง'); cpNum_(data.FreshKg, 'ใบไม้สด/เศษหญ้า'); cpNum_(data.BranchKg, 'กิ่งไม้/เศษไม้');
  return withShareLock_(() => {
    const r = cpInsertBatch_(data, auth.departmentName);
    syncCpLinkedProject_();
    return r;
  });
}

function updateCpBatch(departmentId, pin, batchId, data) {
  assertAuthorized_(departmentId, pin);
  data = data || {};
  return withShareLock_(() => {
    const found = findRowById_(SHEET_CP_BATCHES, 'BatchID', batchId);
    if (!found) throw new Error('ไม่พบรอบการหมัก');
    const f = {};
    if (data.StartDate !== undefined) f.StartDate = cpDate_(data.StartDate, 'วันที่เริ่มหมัก');
    ['DryKg', 'FreshKg', 'BranchKg', 'ActivatorKg', 'WaterL', 'ExpectedKg'].forEach(k => { if (data[k] !== undefined) f[k] = cpNum_(data[k], k); });
    if (data.PileCount !== undefined) { const b = Math.round(cpNum_(data.PileCount, 'จำนวนกอง')); if (b <= 0) throw new Error('จำนวนกองต้องมากกว่า 0'); f.PileCount = b; }
    if (data.ReadyDays !== undefined) f.ReadyDays = cpNum_(data.ReadyDays, 'จำนวนวันที่ใช้หมัก') || '';
    ['Location', 'Note'].forEach(k => { if (data[k] !== undefined) f[k] = String(data[k] || '').trim(); });
    if (data.Status !== undefined) {
      if ([CP_STATUS_FERMENT, CP_STATUS_CANCEL].indexOf(data.Status) === -1) throw new Error('สถานะไม่ถูกต้อง');
      const cur = found.values[found.headers.indexOf('Status')] || CP_STATUS_FERMENT;
      if (cur === CP_STATUS_DONE) {
        if (cpUsedForBatch_(batchId) > 0) throw new Error('รอบนี้มีการนำปุ๋ยหมักไปใช้แล้ว ไม่สามารถยกเลิกการเก็บปุ๋ยหมักได้ (ลบรายการใช้งานก่อน)');
        f.HarvestDate = ''; f.YieldKg = ''; f.HarvestNote = '';
      }
      f.Status = data.Status;
    }
    f.UpdatedAt = new Date();
    cpSetFields_(SHEET_CP_BATCHES, 'BatchID', batchId, f);
    syncCpLinkedProject_();
    return { ok: true };
  });
}

function deleteCpBatch(departmentId, pin, batchId) {
  assertAuthorized_(departmentId, pin);
  return withShareLock_(() => {
    const found = findRowById_(SHEET_CP_BATCHES, 'BatchID', batchId);
    if (!found) throw new Error('ไม่พบรอบการหมัก');
    if (cpUsedForBatch_(batchId) > 0) throw new Error('รอบนี้มีการนำปุ๋ยหมักไปใช้แล้ว ลบไม่ได้ (ถ้าบันทึกผิด ให้ลบรายการใช้งานก่อน หรือเปลี่ยนสถานะเป็นยกเลิก)');
    // ลบผลตรวจของรอบนี้ แล้วลบรอบ (ลบจากท้ายขึ้นมาเพื่อไม่ให้ลำดับแถวเลื่อน)
    const csh = getSheet_(SHEET_CP_CHECKS), cv = csh.getDataRange().getValues(), cb = cv[0].indexOf('BatchID');
    for (let r = cv.length - 1; r >= 1; r--) if (cv[r][cb] === batchId) csh.deleteRow(r + 1);
    const lsh = findSheet_(SHEET_CP_LOGS);
    if (lsh) { const lv = lsh.getDataRange().getValues(), lb = lv[0].indexOf('BatchID'); if (lb !== -1) for (let r = lv.length - 1; r >= 1; r--) if (lv[r][lb] === batchId) getSheet_(SHEET_CP_LOGS).deleteRow(r + 1); }
    getSheet_(SHEET_CP_BATCHES).deleteRow(findRowById_(SHEET_CP_BATCHES, 'BatchID', batchId).rowIndex);
    syncCpLinkedProject_();
    return { ok: true };
  });
}

/** เก็บเป็นปุ๋ยหมัก: บันทึกวันที่และปริมาณปุ๋ยหมักที่ได้จริง (กก.) ของรอบนั้น */
function harvestCpBatch(departmentId, pin, batchId, data) {
  assertAuthorized_(departmentId, pin);
  data = data || {};
  const amount = cpNum_(data.YieldKg, 'ปริมาณปุ๋ยหมักที่ได้');
  if (amount <= 0) throw new Error('กรุณาระบุปริมาณปุ๋ยหมักที่ได้ (กก.)');
  const hd = cpDate_(data.HarvestDate || new Date(), 'วันที่เก็บปุ๋ยหมัก');
  return withShareLock_(() => {
    const found = findRowById_(SHEET_CP_BATCHES, 'BatchID', batchId);
    if (!found) throw new Error('ไม่พบรอบการหมัก');
    const cur = found.values[found.headers.indexOf('Status')] || CP_STATUS_FERMENT;
    if (cur === CP_STATUS_CANCEL) throw new Error('รอบนี้ถูกยกเลิกแล้ว');
    const startN = cpDayNum_(found.values[found.headers.indexOf('StartDate')]);
    if (startN !== null && cpDayNum_(hd) < startN) throw new Error('วันที่เก็บปุ๋ยหมักต้องไม่ก่อนวันที่เริ่มหมัก');
    if (cur === CP_STATUS_DONE && amount < cpUsedForBatch_(batchId)) throw new Error('ปริมาณที่ได้ต้องไม่น้อยกว่าที่นำไปใช้แล้ว (' + round2_(cpUsedForBatch_(batchId)) + ' กก.)');
    cpSetFields_(SHEET_CP_BATCHES, 'BatchID', batchId, { Status: CP_STATUS_DONE, HarvestDate: hd, YieldKg: amount, HarvestNote: String(data.Note || '').trim(), UpdatedAt: new Date() });
    syncCpLinkedProject_();
    return { ok: true };
  });
}

/** เพิ่มผลตรวจ (ไม่ล็อก/ไม่เช็กสิทธิ์) */
function cpInsertCheck_(batchId, data, recordedBy) {
  data = data || {};
  const cd = cpDate_(data.CheckDate || new Date(), 'วันที่ตรวจ');
  const ph = (data.TempC === '' || data.TempC === null || data.TempC === undefined) ? '' : Number(data.TempC);
  if (ph !== '' && (isNaN(ph) || ph < 0 || ph > 100)) throw new Error('อุณหภูมิกองต้องอยู่ระหว่าง 0–100 °C');
  ensureCpSheets_();
  const b = findRowById_(SHEET_CP_BATCHES, 'BatchID', batchId);
  if (!b) throw new Error('ไม่พบรอบการหมัก');
  const startN = cpDayNum_(b.values[b.headers.indexOf('StartDate')]);
  const dayNo = startN === null ? 0 : cpDayNum_(cd) - startN;
  if (dayNo < 0) throw new Error('วันที่ตรวจต้องไม่ก่อนวันที่เริ่มหมัก');
  const week = Math.round(Number(data.Week) || 0);
  if (week > 0 && readRowsIfSheet_(SHEET_CP_CHECKS).some(c => c.BatchID === batchId && Number(c.Week) === week)) throw new Error('บันทึกผลตรวจของสัปดาห์ที่ ' + week + ' ของรอบนี้ไว้แล้ว (ลบรายการเดิมก่อนถ้าต้องการบันทึกใหม่)');
  const id = newId_('CPC');
  appendByHeaders_(SHEET_CP_CHECKS, { CheckID: id, BatchID: batchId, CheckDate: cd, Week: week, DayNo: dayNo, TempC: ph, Smell: String(data.Smell || '').trim(),
    Appearance: String(data.Appearance || '').trim(), Action: String(data.Action || '').trim(), Note: String(data.Note || '').trim(), RecordedBy: recordedBy, CreatedAt: new Date() });
  return { ok: true, checkId: id };
}

function addCpCheck(departmentId, pin, batchId, data) {
  const auth = assertAuthorized_(departmentId, pin);
  data = data || {};
  return withShareLock_(() => {
    const r = cpInsertCheck_(batchId, data, auth.departmentName);
    syncCpLinkedProject_();
    return r;
  });
}

function deleteCpCheck(departmentId, pin, checkId) {
  assertAuthorized_(departmentId, pin);
  const f = findRowById_(SHEET_CP_CHECKS, 'CheckID', checkId);
  if (!f) throw new Error('ไม่พบรายการ');
  getSheet_(SHEET_CP_CHECKS).deleteRow(f.rowIndex);
  return { ok: true };
}

/** บันทึกการนำปุ๋ยหมักไปใช้: วันที่ ปริมาณ สถานที่ ผู้ใช้ วัตถุประสงค์ (ไม่เกินปริมาณคงเหลือ) */
function addCpUsage(departmentId, pin, data) {
  const auth = assertAuthorized_(departmentId, pin);
  data = data || {};
  const amount = cpNum_(data.Kg, 'ปริมาณที่นำไปใช้');
  if (amount <= 0) throw new Error('กรุณาระบุปริมาณที่นำไปใช้ (กก.)');
  const place = String(data.Place || '').trim(), who = String(data.UsedBy || '').trim();
  if (!place) throw new Error('กรุณาระบุสถานที่ที่นำไปใช้');
  if (!who) throw new Error('กรุณาระบุชื่อผู้นำไปใช้');
  const ud = cpDate_(data.UsageDate || new Date(), 'วันที่นำไปใช้');
  return withShareLock_(() => {
    ensureCpSheets_();
    const stock = cpStockKg_();
    if (amount > stock + 1e-9) throw new Error('ปุ๋ยหมักคงเหลือไม่พอ (คงเหลือ ' + stock + ' กก.)');
    const batchId = String(data.BatchID || '').trim();
    if (batchId) {
      const b = findRowById_(SHEET_CP_BATCHES, 'BatchID', batchId);
      if (!b) throw new Error('ไม่พบรอบการหมักที่ระบุ');
      if (b.values[b.headers.indexOf('Status')] !== CP_STATUS_DONE) throw new Error('รอบที่เลือกยังไม่ได้เก็บเป็นปุ๋ยหมัก');
      const remain = Number(b.values[b.headers.indexOf('YieldKg')] || 0) - cpUsedForBatch_(batchId);
      if (amount > remain + 1e-9) throw new Error('ปุ๋ยหมักของรอบนี้คงเหลือไม่พอ (คงเหลือ ' + round2_(remain) + ' กก.)');
    }
    const id = newId_('CPU');
    appendByHeaders_(SHEET_CP_USAGE, { UsageID: id, BatchID: batchId, UsageDate: ud, Kg: amount, Place: place, UsedBy: who, Purpose: String(data.Purpose || '').trim(),
      Note: String(data.Note || '').trim(), RecordedBy: auth.departmentName, CreatedAt: new Date() });
    syncCpLinkedProject_();
    return { ok: true, usageId: id, stockKg: round2_(stock - amount) };
  });
}

function deleteCpUsage(departmentId, pin, usageId) {
  assertAuthorized_(departmentId, pin);
  const f = findRowById_(SHEET_CP_USAGE, 'UsageID', usageId);
  if (!f) throw new Error('ไม่พบรายการ');
  getSheet_(SHEET_CP_USAGE).deleteRow(f.rowIndex);
  syncCpLinkedProject_();
  return { ok: true };
}

/** ซิงก์สถานะแถวโครงการในหน้าภาพรวม */
function syncCpLinkedProject_() {
  try {
    const t = getCpOverview().totals;
    updateLinkedProjectFields_(LINKED_PROJECT_CP_ID, {
      CurrentStatus: t.batches
        ? `หมักแล้ว ${t.batches} รอบ (${t.piles} กอง) • กำลังหมัก ${t.fermentingPiles} กอง${t.readyPiles ? ' (ครบกำหนดเก็บ ' + t.readyPiles + ' กอง)' : ''} • ได้ปุ๋ยหมัก ${t.harvestedKg.toLocaleString()} กก. • ใช้แล้ว ${t.usedKg.toLocaleString()} • คงเหลือ ${t.stockKg.toLocaleString()} กก.`
        : 'เริ่มโครงการ — ยังไม่มีรอบการหมัก (ดูรายละเอียดที่แท็บ "ปุ๋ยหมัก")',
      Progress: yearProgressPct_()
    });
  } catch (e) { /* ยังไม่มีชีต/ข้อมูล ก็ข้ามไปเงียบๆ */ }
}


// ----------------------------- ฟอร์มสาธารณะ: บันทึกการนำปุ๋ยหมักไปใช้ (ลิงก์ / QR Code) -----------------------------
/** ลิงก์ฟอร์ม "บันทึกการนำปุ๋ยหมักไปใช้" สำหรับแชร์/สร้าง QR Code (ไม่ต้องมี PIN) */
function getCpUsageFormLink() {
  return (ScriptApp.getService().getUrl() || '') + '?page=compost&action=use';
}

/**
 * ผู้ใช้ปุ๋ยหมักบันทึกการนำไปใช้เองผ่านลิงก์/QR Code — รายการจะเป็น "รอยืนยัน" ยังไม่หักจากคงเหลือ
 * จนกว่าผู้ดูแลระบบจะยืนยัน (verifySubmissions key = 'cpuse')
 */
function submitCpUsage(data) {
  data = data || {};
  if (String(data.website || '').trim()) return { ok: true, pending: true }; // ช่องล่อบอท (honeypot)
  const amount = cpNum_(data.Kg, 'ปริมาณที่นำไปใช้');
  if (amount <= 0) throw new Error('กรุณาระบุปริมาณที่นำไปใช้ (กก.)');
  if (amount > 100000) throw new Error('ปริมาณมากผิดปกติ กรุณาตรวจสอบอีกครั้ง');
  const place = String(data.Place || '').trim().slice(0, 150), who = String(data.UsedBy || '').trim().slice(0, 120);
  if (!place) throw new Error('กรุณาระบุสถานที่ที่นำไปใช้');
  if (!who) throw new Error('กรุณาระบุชื่อผู้นำไปใช้');
  const ud = cpDate_(data.UsageDate || new Date(), 'วันที่นำไปใช้');
  const todayN = cpToday_(), un = cpDayNum_(ud);
  if (un > todayN) throw new Error('วันที่นำไปใช้ต้องไม่เป็นวันในอนาคต');
  if (un < todayN - 90) throw new Error('วันที่นำไปใช้ย้อนหลังได้ไม่เกิน 90 วัน (หากเก่ากว่านี้ กรุณาแจ้งผู้ดูแลระบบ)');
  return withShareLock_(() => {
    ensureCpSheets_();
    const rows = readRowsIfSheet_(SHEET_CP_USAGE);
    if (rows.filter(u => String(u.VerifyStatus || '') === VERIFY_PENDING).length >= 200) throw new Error('มีรายการรอยืนยันจำนวนมากแล้ว กรุณาแจ้งผู้ดูแลระบบ');
    const stock = cpStockKg_();
    if (amount > stock + 1e-9) throw new Error('ขณะนี้ปุ๋ยหมักคงเหลือ ' + stock + ' กก. ไม่พอตามที่ระบุ กรุณาตรวจสอบปริมาณ');
    const id = newId_('CPU');
    appendByHeaders_(SHEET_CP_USAGE, { UsageID: id, BatchID: '', UsageDate: ud, Kg: amount, Place: place, UsedBy: who, Purpose: String(data.Purpose || '').trim().slice(0, 200),
      Note: String(data.Note || '').trim().slice(0, 300), RecordedBy: 'ผู้ใช้บันทึกเองผ่านลิงก์', CreatedAt: new Date(), VerifyStatus: VERIFY_PENDING });
    return { ok: true, usageId: id, pending: true };
  });
}


// ----------------------------- บันทึกการหมักประจำวัน (ผู้ปฏิบัติกรอกผ่านลิงก์/QR → ผู้ดูแลยืนยัน → ผูกกับรอบการหมัก) -----------------------------
/**
 * ประเภทบันทึก:
 *  - เริ่มหมักกองใหม่   : เมื่อยืนยัน → สร้างรอบการหมักใหม่ (วันที่ วัตถุดิบ จำนวนกอง สถานที่)
 *  - เติมวัตถุดิบเพิ่ม  : เมื่อยืนยัน → บวกวัตถุดิบ (กก.) เข้ารอบที่เลือก (และจำนวนกองเพิ่ม ถ้าระบุ)
 *  - ตรวจ/ดูแลประจำวัน : เมื่อยืนยัน → บันทึกเป็นผลตรวจของรอบที่เลือก (นับในผลตรวจทุก 7 วัน)
 * รายการที่ยังไม่ยืนยันจะไม่นับในข้อมูลใดๆ ทั้งสิ้น
 */
function getCpDailyLogFormLink() {
  return (ScriptApp.getService().getUrl() || '') + '?page=compost&action=log';
}

function cpLogPendingRow_(r) {
  const b = r.BatchID ? findRowById_(SHEET_CP_BATCHES, 'BatchID', r.BatchID) : null;
  const no = b ? Number(b.values[b.headers.indexOf('BatchNo')]) || 0 : 0;
  const kg = (Number(r.DryKg) || 0) + (Number(r.FreshKg) || 0) + (Number(r.BranchKg) || 0);
  let detail = '';
  if (r.LogType === CP_LOG_NEW) detail = 'วัตถุดิบ ' + round2_(kg) + ' กก. (ใบไม้แห้ง ' + (Number(r.DryKg) || 0) + ' • ใบไม้สด ' + (Number(r.FreshKg) || 0) + ' • กิ่งไม้/เศษไม้ ' + (Number(r.BranchKg) || 0) + ') • ' + (Number(r.Piles) || 0) + ' กอง' + (r.Location ? ' • ' + r.Location : '');
  else if (r.LogType === CP_LOG_ADD) detail = 'เติมวัตถุดิบ ' + round2_(kg) + ' กก.' + (Number(r.Piles) > 0 ? ' • เพิ่ม ' + Number(r.Piles) + ' กอง' : '');
  else detail = [r.TempC !== '' && r.TempC !== null && r.TempC !== undefined ? r.TempC + ' °C' : '', r.Smell, r.Appearance, r.Action].filter(Boolean).join(' • ');
  return { id: r.LogID, at: r.CreatedAt ? toIsoDateTime_(r.CreatedAt) : '', date: toIsoDate_(r.LogDate), name: r.Operator || '', type: r.LogType, batchNo: no, detail: detail, note: r.Note || '' };
}

function submitCpDailyLog(data) {
  data = data || {};
  if (String(data.website || '').trim()) return { ok: true, pending: true }; // ช่องล่อบอท
  const type = String(data.LogType || '').trim();
  if (CP_LOG_TYPES.indexOf(type) === -1) throw new Error('กรุณาเลือกประเภทบันทึก');
  const who = String(data.Operator || '').trim().slice(0, 120);
  if (!who) throw new Error('กรุณาระบุชื่อผู้ปฏิบัติงาน');
  const ld = cpDate_(data.LogDate || new Date(), 'วันที่ปฏิบัติงาน');
  const todayN = cpToday_(), ln = cpDayNum_(ld);
  if (ln > todayN) throw new Error('วันที่ปฏิบัติงานต้องไม่เป็นวันในอนาคต');
  if (ln < todayN - 90) throw new Error('วันที่ปฏิบัติงานย้อนหลังได้ไม่เกิน 90 วัน (หากเก่ากว่านี้ กรุณาแจ้งผู้ดูแลระบบ)');
  const food = cpNum_(data.DryKg, 'ใบไม้แห้ง'), fv = cpNum_(data.FreshKg, 'ใบไม้สด/เศษหญ้า'), meat = cpNum_(data.BranchKg, 'กิ่งไม้/เศษไม้');
  if (food > 100000 || fv > 100000 || meat > 100000) throw new Error('ปริมาณมากผิดปกติ กรุณาตรวจสอบอีกครั้ง');
  const piles = Math.round(cpNum_(data.Piles, 'จำนวนกอง'));
  const ph = (data.TempC === '' || data.TempC === null || data.TempC === undefined) ? '' : Number(data.TempC);
  if (ph !== '' && (isNaN(ph) || ph < 0 || ph > 100)) throw new Error('อุณหภูมิกองต้องอยู่ระหว่าง 0–100 °C');
  const batchId = String(data.BatchID || '').trim();
  if (type === CP_LOG_NEW) {
    if (food + fv + meat <= 0) throw new Error('กรุณาระบุปริมาณใบไม้แห้ง ใบไม้สด/เศษหญ้า หรือกิ่งไม้/เศษไม้ อย่างน้อย 1 ช่อง (กก.)');
    if (piles <= 0) throw new Error('กรุณาระบุจำนวนกองที่หมัก');
  } else {
    if (!batchId) throw new Error('กรุณาเลือกรอบการหมัก');
    if (type === CP_LOG_ADD && food + fv + meat <= 0) throw new Error('กรุณาระบุปริมาณวัตถุดิบที่เติม (กก.)');
    if (type === CP_LOG_CHECK && ph === '' && !String(data.Smell || '').trim() && !String(data.Appearance || '').trim() && !String(data.Action || '').trim() && !String(data.Note || '').trim()) throw new Error('กรุณาบันทึกผลที่ตรวจ/สิ่งที่ทำอย่างน้อย 1 ช่อง');
  }
  return withShareLock_(() => {
    ensureCpSheets_();
    if (batchId) {
      const b = findRowById_(SHEET_CP_BATCHES, 'BatchID', batchId);
      if (!b) throw new Error('ไม่พบรอบการหมักที่เลือก');
      const status = b.values[b.headers.indexOf('Status')] || CP_STATUS_FERMENT;
      if (status !== CP_STATUS_FERMENT) throw new Error('รอบที่เลือกไม่ได้อยู่ระหว่างการหมักแล้ว');
      const sn = cpDayNum_(b.values[b.headers.indexOf('StartDate')]);
      if (sn !== null && ln < sn) throw new Error('วันที่ปฏิบัติงานต้องไม่ก่อนวันที่เริ่มหมักของรอบนี้');
    }
    if (readRowsIfSheet_(SHEET_CP_LOGS).filter(l => String(l.VerifyStatus || '') === VERIFY_PENDING).length >= 200) throw new Error('มีรายการรอยืนยันจำนวนมากแล้ว กรุณาแจ้งผู้ดูแลระบบ');
    const id = newId_('CPL');
    appendByHeaders_(SHEET_CP_LOGS, { LogID: id, LogDate: ld, BatchID: batchId, LogType: type, DryKg: food, FreshKg: fv, BranchKg: meat, Piles: piles,
      Location: String(data.Location || '').trim().slice(0, 120), TempC: ph, Smell: String(data.Smell || '').trim().slice(0, 100), Appearance: String(data.Appearance || '').trim().slice(0, 200),
      Action: String(data.Action || '').trim().slice(0, 200), Note: String(data.Note || '').trim().slice(0, 300), Operator: who, CreatedAt: new Date(), VerifyStatus: VERIFY_PENDING });
    return { ok: true, logId: id, pending: true };
  });
}

/** ผู้ดูแลระบบยืนยันบันทึกประจำวัน: ทำผลตามประเภทและผูกกับรอบการหมัก (ยืนยันได้ทีละรายการ; รายการที่ทำผลไม่ได้จะถูกข้ามและแจ้งเหตุผล) */
function verifyCpLogs_(auth, ids) {
  ensureVerifyColumns_(SHEET_CP_LOGS);
  const want = {}; (ids || []).forEach(id => { want[id] = true; });
  return withShareLock_(() => {
    const sh = getSheet_(SHEET_CP_LOGS);
    const values = sh.getDataRange().getValues(), h = values[0];
    const col = n => h.indexOf(n), now = new Date();
    let n = 0; const errors = [];
    for (let r = 1; r < values.length; r++) {
      if (!want[values[r][col('LogID')]] || String(values[r][col('VerifyStatus')] || '') !== VERIFY_PENDING) continue;
      const o = {}; h.forEach((k, i) => { o[k] = values[r][i]; });
      try {
        const batchId = cpApplyLog_(o, auth.departmentName);
        if (batchId && !o.BatchID) sh.getRange(r + 1, col('BatchID') + 1).setValue(batchId);
        sh.getRange(r + 1, col('VerifyStatus') + 1).setValue(VERIFY_OK);
        if (col('VerifiedBy') !== -1) sh.getRange(r + 1, col('VerifiedBy') + 1).setValue(auth.departmentName);
        if (col('VerifiedAt') !== -1) sh.getRange(r + 1, col('VerifiedAt') + 1).setValue(now);
        n++;
      } catch (e) { errors.push((o.Operator || '') + ' (' + toIsoDate_(o.LogDate) + '): ' + (e && e.message || e)); }
    }
    invalidateReadCache_();
    if (n) { try { syncCpLinkedProject_(); } catch (e) {} }
    if (!n && errors.length) throw new Error('ยืนยันไม่ได้: ' + errors.join(' | '));
    return { ok: true, changed: n, errors: errors };
  });
}

/** ทำผลของบันทึกประจำวันหนึ่งรายการ คืนค่า BatchID ที่เกี่ยวข้อง */
function cpApplyLog_(o, by) {
  if (o.LogType === CP_LOG_NEW) {
    const r = cpInsertBatch_({ StartDate: o.LogDate, DryKg: o.DryKg, FreshKg: o.FreshKg, BranchKg: o.BranchKg, PileCount: o.Piles, Location: o.Location,
      Note: 'จากบันทึกประจำวันโดย ' + (o.Operator || '') + (o.Note ? ' • ' + o.Note : '') }, by);
    return r.batchId;
  }
  const f = findRowById_(SHEET_CP_BATCHES, 'BatchID', o.BatchID);
  if (!f) throw new Error('ไม่พบรอบการหมักที่เกี่ยวข้อง (อาจถูกลบแล้ว)');
  const val = k => f.values[f.headers.indexOf(k)];
  if ((val('Status') || CP_STATUS_FERMENT) !== CP_STATUS_FERMENT) throw new Error('รอบนี้ไม่ได้อยู่ระหว่างการหมักแล้ว');
  if (o.LogType === CP_LOG_ADD) {
    const upd = { DryKg: (Number(val('DryKg')) || 0) + (Number(o.DryKg) || 0), FreshKg: (Number(val('FreshKg')) || 0) + (Number(o.FreshKg) || 0),
      BranchKg: (Number(val('BranchKg')) || 0) + (Number(o.BranchKg) || 0), UpdatedAt: new Date() };
    if (Number(o.Piles) > 0) upd.PileCount = (Number(val('PileCount')) || 0) + Number(o.Piles);
    cpSetFields_(SHEET_CP_BATCHES, 'BatchID', o.BatchID, upd);
    return o.BatchID;
  }
  // ตรวจ/ดูแล → ผลตรวจ: เลือกครั้งที่ (สัปดาห์) ที่ถึงกำหนดและยังไม่ได้บันทึกให้อัตโนมัติ
  const st = getCpSettings_(), interval = st.CheckIntervalDays;
  const sn = cpDayNum_(val('StartDate')), ln = cpDayNum_(o.LogDate), age = sn === null ? 0 : ln - sn;
  const ready = Number(val('ReadyDays')) > 0 ? Number(val('ReadyDays')) : st.ReadyDays, nW = Math.max(1, Math.floor(ready / interval));
  const done = readRowsIfSheet_(SHEET_CP_CHECKS).filter(c => c.BatchID === o.BatchID).map(c => Number(c.Week));
  let week = 0;
  for (let k = 1; k <= nW; k++) { if (interval * k <= age && done.indexOf(k) === -1) { week = k; break; } }
  cpInsertCheck_(o.BatchID, { CheckDate: o.LogDate, Week: week, TempC: o.TempC, Smell: o.Smell, Appearance: o.Appearance, Action: o.Action,
    Note: (o.Note ? o.Note + ' • ' : '') + 'จากบันทึกประจำวันโดย ' + (o.Operator || '') }, by);
  return o.BatchID;
}



// ==== COMPOST-END ====

// ============================================================================================
// คำขอรับน้ำยา EM / ปุ๋ยหมักใบไม้ (ลงทะเบียนผ่านลิงก์ → ผู้ขอยืนยันว่าได้รับ → ผู้ดูแลระบบยืนยัน + ส่งอีเมล)
//  สถานะ: รอรับของ → ผู้ขอยืนยันรับแล้ว → ผู้ดูแลยืนยันแล้ว (หักสต็อก + ส่งอีเมล) | ยกเลิก
//  ผู้ขอ: เจ้าหน้าที่ (รหัส 7 หลัก + ชื่อ-นามสกุล + E-mail) | บุคคลภายนอก (ชื่อ + โทร + E-mail) | บริษัทภายนอก (บริษัท + ผู้ติดต่อ + โทร + E-mail)
// ============================================================================================
const SHEET_PRODUCT_REQ = 'ProductRequests';
const PREQ_HEADERS = ['RequestID', 'RequestCode', 'Project', 'CreatedAt', 'RequesterType', 'StaffCode', 'FullName', 'Organization', 'Phone', 'Email',
  'Quantity', 'Unit', 'Purpose', 'Place', 'WantedDate', 'Note', 'Token', 'Status', 'ReceivedAt', 'ConfirmedQty', 'AdminBy', 'AdminAt', 'UsageID',
  'EmailSentAt', 'CancelReason'];
const PREQ_WAIT = 'รอรับของ', PREQ_RECEIVED = 'ผู้ขอยืนยันรับแล้ว', PREQ_DONE = 'ผู้ดูแลยืนยันแล้ว', PREQ_CANCEL = 'ยกเลิก';
const PREQ_TYPES = ['เจ้าหน้าที่', 'บุคคลภายนอก', 'บริษัทภายนอก'];
const PREQ_CFG = {
  em:      { label: 'น้ำยา EM', full: 'น้ำยา EM จากเศษอาหาร', unit: 'ลิตร', page: 'em', prefix: 'EM', usage: SHEET_EM_USAGE, qty: 'Liters', usageId: 'EMU' },
  compost: { label: 'ปุ๋ยหมัก', full: 'ปุ๋ยหมักจากขยะใบไม้', unit: 'กก.', page: 'compost', prefix: 'CP', usage: SHEET_CP_USAGE, qty: 'Kg', usageId: 'CPU' }
};
function preqCfg_(project) {
  const c = PREQ_CFG[project];
  if (!c) throw new Error('โครงการนี้ไม่มีการลงทะเบียนขอรับ');
  return c;
}
function preqStock_(project) { return project === 'em' ? emStockLiters_() : cpStockKg_(); }
function preqSync_(project) { try { project === 'em' ? syncEmLinkedProject_() : syncCpLinkedProject_(); } catch (e) {} }
function preqEnsure_(project) {
  if (project === 'em') ensureEmSheets_(); else ensureCpSheets_();
  ensureSheet_(SHEET_PRODUCT_REQ, PREQ_HEADERS);
}
function preqEmailOk_(v) { return /^[^\s@]+@[^\s@]+\.[^\s@]{2,}$/.test(String(v || '').trim()); }
function preqLink_(r) {
  return (ScriptApp.getService().getUrl() || '') + '?page=' + preqCfg_(r.Project).page + '&action=receive&rid=' + encodeURIComponent(r.RequestID) + '&t=' + encodeURIComponent(r.Token);
}
function preqWho_(r) { return r.RequesterType === 'บริษัทภายนอก' ? (r.Organization + (r.FullName ? ' (' + r.FullName + ')' : '')) : r.FullName; }
function preqHtmlEsc_(v) { return String(v == null ? '' : v).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c])); }
function preqMail_(to, subject, lines, link, linkText) {
  if (!preqEmailOk_(to)) return false;
  try {
    const body = lines.map(l => '<p style="margin:0 0 10px">' + l + '</p>').join('') +
      (link ? '<p style="margin:18px 0"><a href="' + link + '" style="background:#16a34a;color:#fff;padding:11px 20px;border-radius:10px;text-decoration:none;font-weight:bold">' + preqHtmlEsc_(linkText) + '</a></p>' : '');
    MailApp.sendEmail({ to: String(to).trim(), subject: subject, htmlBody:
      '<div style="font-family:Tahoma,sans-serif;font-size:15px;color:#1f2937;max-width:560px">' + body +
      '<p style="color:#6b7280;font-size:13px;margin-top:24px">อีเมลนี้ส่งอัตโนมัติจากระบบติดตามโครงการสิ่งแวดล้อมโรงพยาบาล</p></div>' });
    return true;
  } catch (e) { Logger.log('ส่งอีเมลไม่สำเร็จ: ' + e); return false; }
}
function preqPublic_(r) {
  const c = preqCfg_(r.Project);
  return { RequestID: r.RequestID, RequestCode: r.RequestCode, Project: r.Project, ProjectLabel: c.full, Unit: r.Unit || c.unit, CreatedAt: r.CreatedAt ? toIsoDateTime_(r.CreatedAt) : '',
    RequesterType: r.RequesterType, Who: preqWho_(r), Quantity: Number(r.Quantity) || 0, ConfirmedQty: Number(r.ConfirmedQty) || 0, Purpose: r.Purpose || '', Place: r.Place || '',
    WantedDate: toIsoDate_(r.WantedDate), Status: r.Status || PREQ_WAIT, ReceivedAt: r.ReceivedAt ? toIsoDateTime_(r.ReceivedAt) : '', AdminAt: r.AdminAt ? toIsoDateTime_(r.AdminAt) : '' };
}
function preqFindRow_(rid) {
  const f = findRowById_(SHEET_PRODUCT_REQ, 'RequestID', rid);
  if (!f) return null;
  const o = {}; f.headers.forEach((h, i) => { o[h] = f.values[i]; });
  return { f: f, r: o };
}
function preqSet_(f, obj) {
  const sh = getSheet_(SHEET_PRODUCT_REQ);
  Object.keys(obj).forEach(k => { const c = f.headers.indexOf(k); if (c !== -1) sh.getRange(f.rowIndex, c + 1).setValue(obj[k]); });
  invalidateReadCache_();
}

/** ลิงก์ฟอร์มลงทะเบียนขอรับ (สำหรับแชร์/QR Code) */
function getProductRequestFormLink(project) {
  return (ScriptApp.getService().getUrl() || '') + '?page=' + preqCfg_(project).page + '&action=request';
}

/** สาธารณะ: ลงทะเบียนขอรับน้ำยา EM / ปุ๋ยหมัก */
function submitProductRequest(data) {
  data = data || {};
  if (String(data.website || '').trim()) return { ok: true }; // honeypot
  const project = String(data.Project || ''), c = preqCfg_(project);
  const type = String(data.RequesterType || '');
  if (PREQ_TYPES.indexOf(type) === -1) throw new Error('กรุณาเลือกประเภทผู้ขอ');
  const clip = (v, n) => String(v == null ? '' : v).replace(/\s+/g, ' ').trim().slice(0, n);
  const staff = clip(data.StaffCode, 7), name = clip(data.FullName, 120), org = clip(data.Organization, 150), phone = clip(data.Phone, 30), email = clip(data.Email, 120);
  if (type === 'เจ้าหน้าที่' && !/^\d{7}$/.test(staff)) throw new Error('รหัสเจ้าหน้าที่ต้องเป็นตัวเลข 7 หลัก');
  if (type === 'บริษัทภายนอก' && !org) throw new Error('กรุณากรอกชื่อบริษัท');
  if (name.length < 3) throw new Error(type === 'บริษัทภายนอก' ? 'กรุณากรอกชื่อผู้ติดต่อ' : 'กรุณากรอกชื่อ-นามสกุล');
  if (type !== 'เจ้าหน้าที่' && !/^[0-9+\-\s()]{9,}$/.test(phone)) throw new Error('กรุณากรอกเบอร์โทรศัพท์ให้ถูกต้อง');
  if (!preqEmailOk_(email)) throw new Error('กรุณากรอก E-mail ให้ถูกต้อง');
  const qty = emNum_(data.Quantity, 'ปริมาณที่ขอรับ');
  if (!(qty > 0)) throw new Error('กรุณาระบุปริมาณที่ขอรับ (' + c.unit + ')');
  if (qty > 100000) throw new Error('ปริมาณมากผิดปกติ กรุณาตรวจสอบอีกครั้ง');
  const purpose = clip(data.Purpose, 200);
  if (!purpose) throw new Error('กรุณาระบุวัตถุประสงค์ในการนำไปใช้');
  const wanted = data.WantedDate ? emDate_(data.WantedDate, 'วันที่ต้องการรับ') : '';

  const out = withShareLock_(() => {
    preqEnsure_(project);
    const rows = readRowsIfSheet_(SHEET_PRODUCT_REQ);
    if (rows.filter(r => r.Status === PREQ_WAIT || r.Status === PREQ_RECEIVED).length >= 300) throw new Error('มีคำขอค้างจำนวนมาก กรุณาติดต่อผู้ดูแลโครงการ');
    const stock = preqStock_(project);
    if (qty > stock + 1e-9) throw new Error('ขณะนี้' + c.label + ' คงเหลือ ' + round2_(stock) + ' ' + c.unit + ' ไม่พอตามที่ขอ กรุณาลดปริมาณ');
    const alpha = 'ABCDEFGHJKMNPQRSTUVWXYZ23456789';
    let code; do { code = c.prefix + '-'; for (let i = 0; i < 5; i++) code += alpha.charAt(Math.floor(Math.random() * alpha.length)); } while (rows.some(r => r.RequestCode === code));
    const row = { RequestID: newId_('RQ'), RequestCode: code, Project: project, CreatedAt: new Date(), RequesterType: type, StaffCode: type === 'เจ้าหน้าที่' ? staff : '',
      FullName: name, Organization: type === 'บริษัทภายนอก' ? org : '', Phone: phone, Email: email, Quantity: qty, Unit: c.unit, Purpose: purpose,
      Place: clip(data.Place, 150), WantedDate: wanted, Note: clip(data.Note, 300), Token: Utilities.getUuid().replace(/-/g, ''), Status: PREQ_WAIT };
    appendByHeaders_(SHEET_PRODUCT_REQ, row);
    return row;
  });
  const link = preqLink_(out);
  const mailed = preqMail_(out.Email, '[' + c.full + '] ลงทะเบียนขอรับเรียบร้อย — รหัสคำขอ ' + out.RequestCode, [
    'เรียน ' + preqHtmlEsc_(preqWho_(out)),
    'ระบบได้รับคำขอรับ<b>' + c.label + ' ' + out.Quantity + ' ' + c.unit + '</b> ของท่านแล้ว รหัสคำขอ <b>' + out.RequestCode + '</b>',
    'เมื่อท่านได้รับของแล้ว กรุณากดปุ่มด้านล่างเพื่อยืนยันการรับ จากนั้นผู้ดูแลโครงการจะยืนยันอีกครั้งและแจ้งผลทางอีเมลนี้'
  ], link, '✅ ยืนยันว่าได้รับแล้ว');
  return { ok: true, requestId: out.RequestID, code: out.RequestCode, link: link, mailed: mailed };
}

/** สาธารณะ: ดูคำขอจากลิงก์ส่วนตัว (rid + token) หรือจากรหัสคำขอ + E-mail/เบอร์โทร/รหัสเจ้าหน้าที่ */
function getProductRequestForReceive(q) {
  q = q || {};
  ensureSheet_(SHEET_PRODUCT_REQ, PREQ_HEADERS);
  const rows = readRowsIfSheet_(SHEET_PRODUCT_REQ);
  let r = null;
  if (q.rid && q.token) r = rows.find(x => x.RequestID === q.rid && String(x.Token) === String(q.token));
  else if (q.code && q.contact) {
    const code = String(q.code).trim().toUpperCase(), ct = String(q.contact).replace(/\s+/g, '').toLowerCase();
    r = rows.find(x => String(x.RequestCode).toUpperCase() === code &&
      [x.Email, x.Phone, x.StaffCode].some(v => v && String(v).replace(/\s+/g, '').toLowerCase() === ct));
  }
  if (!r) throw new Error('ไม่พบคำขอ — ตรวจสอบรหัสคำขอและ E-mail / เบอร์โทร / รหัสเจ้าหน้าที่ อีกครั้ง');
  return Object.assign(preqPublic_(r), { token: r.Token });
}

/** สาธารณะ: ผู้ขอกดยืนยันว่าได้รับของแล้ว */
function confirmProductReceived(rid, token) {
  return withShareLock_(() => {
    const x = preqFindRow_(rid);
    if (!x || String(x.r.Token) !== String(token || '')) throw new Error('ลิงก์ไม่ถูกต้องหรือหมดอายุ');
    const st = x.r.Status || PREQ_WAIT;
    if (st === PREQ_CANCEL) throw new Error('คำขอนี้ถูกยกเลิกแล้ว');
    if (st === PREQ_WAIT) preqSet_(x.f, { Status: PREQ_RECEIVED, ReceivedAt: new Date() });
    return Object.assign(preqPublic_(preqFindRow_(rid).r), { ok: true });
  });
}

/** ผู้ดูแลระบบ: รายการคำขอทั้งหมดของโครงการ (ใหม่สุดก่อน) */
function getProductRequests(departmentId, pin, project) {
  assertAuthorized_(departmentId, pin);
  preqCfg_(project);
  ensureSheet_(SHEET_PRODUCT_REQ, PREQ_HEADERS);
  return readRowsIfSheet_(SHEET_PRODUCT_REQ).filter(r => r.RequestID && r.Project === project)
    .map(r => Object.assign(preqPublic_(r), { FullName: r.FullName, Organization: r.Organization, StaffCode: String(r.StaffCode || ''), Phone: r.Phone || '', Email: r.Email || '',
      Note: r.Note || '', AdminBy: r.AdminBy || '', CancelReason: r.CancelReason || '', EmailSentAt: r.EmailSentAt ? toIsoDateTime_(r.EmailSentAt) : '' }))
    .sort((a, b) => String(b.CreatedAt).localeCompare(String(a.CreatedAt))).slice(0, 300);
}

/** ผู้ดูแลระบบ: ยืนยันว่าส่งมอบแล้ว → บันทึกการนำไปใช้ (หักสต็อก) + ส่งอีเมลแจ้งผู้ขอ */
function adminConfirmProductRequest(departmentId, pin, rid, qty) {
  const auth = assertAuthorized_(departmentId, pin);
  const res = withShareLock_(() => {
    const x = preqFindRow_(rid);
    if (!x) throw new Error('ไม่พบคำขอ');
    const r = x.r, c = preqCfg_(r.Project);
    if (r.Status === PREQ_DONE) throw new Error('คำขอนี้ยืนยันไปแล้ว');
    if (r.Status === PREQ_CANCEL) throw new Error('คำขอนี้ถูกยกเลิกแล้ว');
    const q = qty === undefined || qty === null || qty === '' ? Number(r.Quantity) : emNum_(qty, 'ปริมาณที่ส่งมอบ');
    if (!(q > 0)) throw new Error('ปริมาณที่ส่งมอบต้องมากกว่า 0');
    preqEnsure_(r.Project);
    const stock = preqStock_(r.Project);
    if (q > stock + 1e-9) throw new Error(c.label + ' คงเหลือไม่พอ (คงเหลือ ' + round2_(stock) + ' ' + c.unit + ')');
    const now = new Date(), uid = newId_(c.usageId), usage = { UsageID: uid, BatchID: '', UsageDate: r.ReceivedAt || now, Place: r.Place || (r.RequesterType === 'เจ้าหน้าที่' ? 'ภายในโรงพยาบาล' : 'ภายนอกโรงพยาบาล'),
      UsedBy: preqWho_(r) + (r.StaffCode ? ' (' + r.StaffCode + ')' : ''), Purpose: r.Purpose || '', Note: 'จากคำขอ ' + r.RequestCode + ' • ' + r.RequesterType,
      RecordedBy: auth.departmentName, CreatedAt: now, VerifyStatus: VERIFY_OK, VerifiedBy: auth.departmentName, VerifiedAt: now };
    usage[c.qty] = q;
    appendByHeaders_(c.usage, usage);
    preqSet_(x.f, { Status: PREQ_DONE, ConfirmedQty: q, AdminBy: auth.departmentName, AdminAt: now, UsageID: uid, ReceivedAt: r.ReceivedAt || now });
    return { r: r, c: c, q: q, f: x.f };
  });
  const r = res.r, c = res.c;
  const mailed = preqMail_(r.Email, '[' + c.full + '] ยืนยันการรับ' + c.label + ' — รหัสคำขอ ' + r.RequestCode, [
    'เรียน ' + preqHtmlEsc_(preqWho_(r)),
    'ผู้ดูแลโครงการ <b>' + c.full + '</b> ยืนยันแล้วว่าท่านได้รับ<b>' + c.label + ' ' + res.q + ' ' + c.unit + '</b> เรียบร้อย',
    'รหัสคำขอ: <b>' + r.RequestCode + '</b>' + (r.Purpose ? '<br>วัตถุประสงค์: ' + preqHtmlEsc_(r.Purpose) : ''),
    'ขอบคุณที่ร่วมนำผลผลิตจากโครงการสิ่งแวดล้อมของโรงพยาบาลไปใช้ประโยชน์ 🌱'
  ]);
  if (mailed) withShareLock_(() => preqSet_(res.f, { EmailSentAt: new Date() }));
  preqSync_(r.Project);
  return { ok: true, mailed: mailed };
}

/** ผู้ดูแลระบบ: ยกเลิกคำขอ */
function adminCancelProductRequest(departmentId, pin, rid, reason) {
  const auth = assertAuthorized_(departmentId, pin);
  return withShareLock_(() => {
    const x = preqFindRow_(rid);
    if (!x) throw new Error('ไม่พบคำขอ');
    if (x.r.Status === PREQ_DONE) throw new Error('คำขอนี้ยืนยันส่งมอบแล้ว ยกเลิกไม่ได้ (ลบรายการการใช้ในประวัติแทน)');
    preqSet_(x.f, { Status: PREQ_CANCEL, CancelReason: String(reason || '').slice(0, 200), AdminBy: auth.departmentName, AdminAt: new Date() });
    return { ok: true };
  });
}

// ----------------------------- ตั้งค่ารอบการหมัก (ตามเกณฑ์) จากหน้าเว็บ -----------------------------
const FERMENT_SETTING_KEYS = ['Stage1Days', 'Stage2Days', 'ReadyDays', 'CheckIntervalDays'];
function saveFermentSettings(departmentId, pin, project, values) {
  assertAuthorized_(departmentId, pin);
  preqCfg_(project);
  values = values || {};
  const v = {};
  FERMENT_SETTING_KEYS.forEach(k => {
    const n = Math.round(Number(values[k]));
    if (!(n >= 1 && n <= 365)) throw new Error('จำนวนวันต้องอยู่ระหว่าง 1–365 วัน');
    v[k] = n;
  });
  if (!(v.Stage1Days < v.Stage2Days && v.Stage2Days < v.ReadyDays)) throw new Error('ต้องเรียงจากน้อยไปมาก: ระยะที่ 1 < ระยะที่ 2 < ครบกำหนด');
  if (v.CheckIntervalDays > v.ReadyDays) throw new Error('ระยะตรวจต้องไม่เกินจำนวนวันครบกำหนด');
  const sheet = project === 'em' ? SHEET_EM_SETTINGS : SHEET_CP_SETTINGS;
  if (project === 'em') ensureEmSettingsSeeded_(); else ensureCpSettingsSeeded_();
  return withShareLock_(() => {
    const sh = getSheet_(sheet), vals = sh.getDataRange().getValues();
    FERMENT_SETTING_KEYS.forEach(k => {
      let row = -1;
      for (let r = 1; r < vals.length; r++) if (vals[r][0] === k) { row = r + 1; break; }
      if (row === -1) sh.appendRow([k, v[k]]); else sh.getRange(row, 2).setValue(v[k]);
    });
    invalidateReadCache_();
    preqSync_(project);
    return { ok: true, settings: v };
  });
}

// ============================================================================================
// วัตถุดิบเข้า-ออกรายเดือน (น้ำยา EM: ขยะเปียก / ปุ๋ยหมัก: ใบไม้ เศษหญ้า กิ่งไม้)
//  เข้า  = ปริมาณที่ผู้ดูแลระบบบันทึกว่ามีในระบบแต่ละเดือน (ชีต RawMaterialMonthly)
//  ออก  = วัตถุดิบที่นำไปหมักจริง: ปริมาณตั้งต้นของรอบ → เดือนที่เริ่มหมัก, การเติมวัตถุดิบ (ยืนยันแล้ว) → เดือนที่เติม
//  คงเหลือยกไป = คงเหลือยกมา + เข้า − ออก (รอบที่ยกเลิกไม่นับ)
// ============================================================================================
const SHEET_RAW_MONTHLY = 'RawMaterialMonthly';
const RAW_HEADERS = ['Key', 'Project', 'Year', 'Month', 'InKg', 'Note', 'UpdatedBy', 'UpdatedAt'];
const RAW_CFG = {
  em:      { label: 'ขยะเปียก (เศษอาหาร ผัก ผลไม้ เนื้อสัตว์)', batches: () => SHEET_EM_BATCHES, logs: () => SHEET_EM_LOGS, cancel: () => EM_STATUS_CANCEL, add: () => EM_LOG_ADD,
             kg: o => (Number(o.FoodKg) || 0) + (Number(o.FruitVegKg) || 0) + (Number(o.MeatKg) || 0) },
  compost: { label: 'ใบไม้ เศษหญ้า กิ่งไม้', batches: () => SHEET_CP_BATCHES, logs: () => SHEET_CP_LOGS, cancel: () => CP_STATUS_CANCEL, add: () => CP_LOG_ADD,
             kg: o => (Number(o.DryKg) || 0) + (Number(o.FreshKg) || 0) + (Number(o.BranchKg) || 0) }
};
function rawCfg_(project) { const c = RAW_CFG[project]; if (!c) throw new Error('โครงการนี้ไม่มีบัญชีวัตถุดิบ'); return c; }
function rawYm_(v) { const s = toIsoDate_(v); return s ? s.slice(0, 7) : ''; }

/** ปริมาณวัตถุดิบที่นำไปหมักจริง แยกตามเดือน { 'YYYY-MM': kg } */
function rawUsedByMonth_(project) {
  const c = rawCfg_(project), out = {};
  const put = (ym, kg) => { if (ym && kg > 0) out[ym] = round2_((out[ym] || 0) + kg); };
  const adds = readRowsIfSheet_(c.logs()).filter(l => l.LogType === c.add() && l.BatchID && String(l.VerifyStatus || '') === VERIFY_OK);
  readRowsIfSheet_(c.batches()).filter(b => b.BatchID && b.Status !== c.cancel()).forEach(b => {
    const mine = adds.filter(l => l.BatchID === b.BatchID);
    const addKg = mine.reduce((s, l) => s + c.kg(l), 0);
    put(rawYm_(b.StartDate), Math.max(0, c.kg(b) - addKg));
    mine.forEach(l => put(rawYm_(l.LogDate), c.kg(l)));
  });
  return out;
}

/** บัญชีวัตถุดิบเข้า-ออกรายเดือนของปี (ค.ศ.) — เปิดดูได้ทุกคน */
function getRawBalance(project, year) {
  const c = rawCfg_(project);
  year = Number(year) || new Date().getFullYear();
  ensureSheet_(SHEET_RAW_MONTHLY, RAW_HEADERS);
  const ins = {}, notes = {};
  readRowsIfSheet_(SHEET_RAW_MONTHLY).filter(r => r.Project === project).forEach(r => {
    const ym = Number(r.Year) + '-' + String(Number(r.Month)).padStart(2, '0');
    ins[ym] = round2_(Number(r.InKg) || 0); notes[ym] = r.Note || '';
  });
  const used = rawUsedByMonth_(project);
  const allYm = Object.keys(ins).concat(Object.keys(used));
  const years = allYm.map(k => Number(k.slice(0, 4))).concat([new Date().getFullYear()]).filter((y, i, a) => a.indexOf(y) === i).sort((a, b) => b - a);
  // ยอดยกมาต้นปี = สะสมก่อนปีนี้
  const before = k => k < year + '-01';
  let carry = round2_(allYm.filter((k, i, a) => a.indexOf(k) === i && before(k)).reduce((s, k) => s + (ins[k] || 0) - (used[k] || 0), 0));
  const opening = carry, nowYm = toIsoDate_(new Date()).slice(0, 7);
  const months = [];
  for (let m = 1; m <= 12; m++) {
    const ym = year + '-' + String(m).padStart(2, '0'), i = ins[ym] || 0, u = used[ym] || 0;
    const open = carry; carry = round2_(carry + i - u);
    months.push({ month: m, ym: ym, opening: open, inKg: i, usedKg: u, closing: carry, hasIn: Object.prototype.hasOwnProperty.call(ins, ym), note: notes[ym] || '', future: ym > nowYm,
      usePct: (open + i) > 0 ? Math.round(u / (open + i) * 100) : null });
  }
  const tIn = round2_(months.reduce((s, x) => s + x.inKg, 0)), tUsed = round2_(months.reduce((s, x) => s + x.usedKg, 0));
  return { ok: true, project: project, label: c.label, year: year, years: years, opening: opening, months: months,
    totals: { inKg: tIn, usedKg: tUsed, closing: carry, usePct: (opening + tIn) > 0 ? Math.round(tUsed / (opening + tIn) * 100) : null } };
}

/** ผู้ดูแลระบบ: บันทึก/แก้ไขปริมาณวัตถุดิบที่มีในระบบของเดือน (kg ว่าง = ลบรายการเดือนนั้น) */
function saveRawMonthly(departmentId, pin, project, year, month, kg, note) {
  const auth = assertAuthorized_(departmentId, pin);
  rawCfg_(project);
  year = Math.round(Number(year)); month = Math.round(Number(month));
  if (!(year >= 2000 && year <= 2100) || !(month >= 1 && month <= 12)) throw new Error('เดือน/ปีไม่ถูกต้อง');
  const empty = kg === '' || kg === null || kg === undefined;
  const v = empty ? 0 : Number(kg);
  if (!empty && !(v >= 0 && v < 1e7)) throw new Error('ปริมาณ (กก.) ไม่ถูกต้อง');
  const key = project + '-' + year + '-' + String(month).padStart(2, '0');
  return withShareLock_(() => {
    ensureSheet_(SHEET_RAW_MONTHLY, RAW_HEADERS);
    const f = findRowById_(SHEET_RAW_MONTHLY, 'Key', key);
    if (empty) { if (f) getSheet_(SHEET_RAW_MONTHLY).deleteRow(f.rowIndex); invalidateReadCache_(); return { ok: true, deleted: true }; }
    const row = { Key: key, Project: project, Year: year, Month: month, InKg: round2_(v), Note: String(note || '').slice(0, 200), UpdatedBy: auth.departmentName, UpdatedAt: new Date() };
    if (f) { const sh = getSheet_(SHEET_RAW_MONTHLY); f.headers.forEach((h, i) => { if (h in row) sh.getRange(f.rowIndex, i + 1).setValue(row[h]); }); }
    else appendByHeaders_(SHEET_RAW_MONTHLY, row);
    invalidateReadCache_();
    return { ok: true };
  });
}
