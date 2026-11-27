/**********************************************************************
 * Hub.gs — ศูนย์ร่วมโครงการ (Eco Hub)
 * ----------------------------------------------------------------
 * แก้ปัญหา: โครงการมีทั้งแบบ "รายบุคคล" "รายหน่วยงาน" และ "คนนอกร่วมได้" แต่หน้าเว็บเดิม
 * ต้องสมัครทีละโครงการ รอผู้ดูแลอนุมัติทุกครั้ง และคนนอกไม่รู้ว่าร่วมอะไรได้บ้าง
 *
 * หลักการ
 *  1) ผู้ดูแลกำหนดในชีต/หน้าเว็บว่าแต่ละโครงการ "เปิดให้ใครร่วม" ได้บ้าง (HubProjects):
 *       เจ้าหน้าที่ (staff) · หน่วยงาน (dept) · ประชาชนทั่วไป (public) · บริษัท/องค์กร (company)
 *     พร้อม "อนุมัติอัตโนมัติ" ได้รายโครงการ → ลดขั้นตอนรออนุมัติสำหรับโครงการทั่วไป
 *  2) สมัครได้หลายโครงการในครั้งเดียว ทั้งบุคคลและหน่วยงาน
 *       - เจ้าหน้าที่: ใช้บัญชีเดิมของ Portal.gs (ชีต PortalJoinRequests → ใช้กับกฎสมาชิกโครงการเดิมได้ทันที)
 *       - หน่วยงาน: เจ้าหน้าที่ที่เข้าสู่ระบบยื่นคำขอแทนหน่วยงาน → ผู้ดูแลอนุมัติ → เพิ่มชื่อหน่วยงานในโครงการให้อัตโนมัติ
 *       - ประชาชน/บริษัท: ลงทะเบียนครั้งเดียว (ไม่ต้องตั้งรหัสผ่าน) ได้ "รหัสสมาชิก" ใช้คู่กับเบอร์โทรเพื่อดู/เพิ่มโครงการ
 *  3) ตัวเลขของทุกโครงการดูได้ทั้ง "ปีปฏิทิน พ.ศ." และ "ปีงบประมาณ" (ต.ค.–ก.ย.) — ใช้ getPeriodSummary(key, year, basis)
 *  4) PDPA: API สาธารณะคืนเฉพาะตัวเลขรวมและชื่อหน่วยงาน · ข้อมูลบุคคลภายนอกเก็บในไฟล์ส่วนตัว (portalSs_)
 *     หน้าเว็บแสดงชื่อ/เบอร์แบบปกปิดบางส่วน · ผู้ดูแลเห็นข้อมูลเต็มหลังตรวจ PIN และทุกครั้งถูกบันทึก audit
 **********************************************************************/

const SHEET_HUB_PROJECTS = 'HubProjects';
const HUB_PROJECT_HEADERS = ['ProjectKey', 'Label', 'Emoji', 'Color', 'Tagline', 'HowTo', 'ForStaff', 'ForDept', 'ForPublic',
  'ForCompany', 'IsOpen', 'AutoApprove', 'SortOrder', 'UpdatedAt', 'UpdatedBy'];
const SHEET_HUB_MEMBERS = 'HubMembers';
const HUB_MEMBER_HEADERS = ['MemberID', 'MemberType', 'Name', 'OrgName', 'Phone', 'Email', 'CreatedAt', 'ConsentAt', 'NoticeVersion', 'LastSeenAt'];
const SHEET_HUB_JOINS = 'HubJoins';
const HUB_JOIN_HEADERS = ['JoinID', 'MemberType', 'MemberRef', 'ProjectKey', 'Status', 'CreatedAt', 'RequestedBy', 'ReviewedBy', 'ReviewedAt', 'Note'];

const HUB_AUDIENCES = ['staff', 'dept', 'public', 'company'];
const HUB_AUD_COL = { staff: 'ForStaff', dept: 'ForDept', public: 'ForPublic', company: 'ForCompany' };
const HUB_AUD_LABEL = { staff: 'เจ้าหน้าที่', dept: 'หน่วยงาน', public: 'ประชาชนทั่วไป', company: 'บริษัท/องค์กร' };
/** key ของโครงการ (ตรงกับ PORTAL_JOIN_PROJECTS) → key ของ getPeriodSummary */
const HUB_PERIOD_KEY = { frytofly: 'ftf', paperless: 'paperless', cartridge: 'cartridge', kitchen: 'kitchen', calendar: 'calendar', share: 'share', em: 'em', compost: 'compost' };
/** ตัวเลขหลักที่โชว์บนการ์ดของแต่ละโครงการ (คอลัมน์ของ getPeriodSummary) */
const HUB_MAIN_METRIC = { frytofly: 'liters', paperless: 'sheets', cartridge: 'contributed', kitchen: 'liters', calendar: 'kg', share: 'value', em: 'emLiters', compost: 'cpKg' };
/** ฟอร์มร่วมกิจกรรมในระบบหลัก (?page=<key>&action=<action>) */
const HUB_ACTIONS = { frytofly: 'form', cartridge: 'form', calendar: 'form', share: 'form', em: 'request', compost: 'request' };
/** โครงการที่มีชีตรายชื่อหน่วยงานอยู่แล้ว — อนุมัติคำขอหน่วยงานแล้วเพิ่มชื่อในชีตนั้นให้ (addProjectDepartment) */
const HUB_DEPT_SHEET_KEYS = { paperless: true, cartridge: true, share: true };

/** ค่าเริ่มต้นตอนสร้างชีต HubProjects ครั้งแรก — ผู้ดูแลแก้ได้ภายหลังจากหน้าเว็บ (ปุ่มผู้ดูแล → ตั้งค่าโครงการ) */
const HUB_DEFAULTS = [
  { key: 'frytofly',  e: '🛢️', c: '#d97706', aud: [1, 0, 0, 0], open: 1, auto: 1,
    t: 'นำน้ำมันพืชใช้แล้วจากบ้านมาบริจาค สะสมแต้ม ลุ้นระดับสมาชิก', how: 'เทน้ำมันใช้แล้วใส่แกลลอน นำมาส่งจุดรับ แล้วบันทึกด้วยรหัสสมาชิก' },
  { key: 'paperless', e: '📄', c: '#2563eb', aud: [0, 1, 0, 0], open: 1, auto: 0,
    t: 'เปลี่ยนงานเอกสารของหน่วยงานเป็นอิเล็กทรอนิกส์ ลดการพิมพ์กระดาษ', how: 'หน่วยงานรายงานจำนวนเรื่อง/แผ่นที่ลดได้ทุกเดือน' },
  { key: 'cartridge', e: '🖨️', c: '#7c3aed', aud: [0, 1, 0, 0], open: 1, auto: 0,
    t: 'รวบรวมตลับหมึกใช้แล้วของหน่วยงาน ส่งจำหน่าย ลดขยะอันตราย', how: 'หน่วยงานรวบรวมตลับหมึก แล้วบันทึกส่งเข้าโครงการ' },
  { key: 'kitchen',   e: '🍳', c: '#ea580c', aud: [0, 1, 0, 0], open: 0, auto: 0,
    t: 'น้ำมันทอดจากครัวโรงพยาบาล ส่งต่อเป็นเชื้อเพลิงอากาศยานยั่งยืน', how: 'ดำเนินการโดยฝ่ายโภชนาการ — ทุกคนติดตามผลได้' },
  { key: 'calendar',  e: '📅', c: '#0891b2', aud: [1, 1, 1, 1], open: 1, auto: 1,
    t: 'บริจาคปฏิทินตั้งโต๊ะเก่า เพื่อทำสื่ออักษรเบรลล์ให้ผู้พิการทางสายตา', how: 'นำปฏิทินเก่ามาส่งที่จุดรับ หรือแจ้งส่งผ่านฟอร์ม' },
  { key: 'share',     e: '♻️', c: '#059669', aud: [1, 1, 0, 0], open: 1, auto: 1,
    t: 'แบ่งปันวัสดุสำนักงานที่เกินจำเป็น ให้หน่วยงานอื่นเบิกใช้แทนการซื้อใหม่', how: 'แจ้งของที่จะแบ่งปันผ่านฟอร์ม แล้วนำมาส่งคลังแบ่งปัน' },
  { key: 'em',        e: '🧪', c: '#0d9488', aud: [1, 1, 1, 1], open: 1, auto: 1,
    t: 'เศษอาหารกลายเป็นน้ำยา EM — ลงทะเบียนขอรับไปใช้ดับกลิ่น รดต้นไม้', how: 'กรอกคำขอรับน้ำยา EM แล้วรับของตามนัด' },
  { key: 'compost',   e: '🪴', c: '#65a30d', aud: [1, 1, 1, 1], open: 1, auto: 1,
    t: 'ใบไม้และเศษหญ้าในโรงพยาบาลกลายเป็นปุ๋ยหมัก — ลงทะเบียนขอรับได้', how: 'กรอกคำขอรับปุ๋ยหมัก แล้วรับของตามนัด' }
];

const HUB_JOIN_PENDING = 'pending', HUB_JOIN_APPROVED = 'approved', HUB_JOIN_REJECTED = 'rejected';
const HUB_ID_CHARS = 'ABCDEFGHJKMNPQRSTUVWXYZ23456789';   // ไม่มี 0/O/1/I/L ที่อ่านสับสน
const HUB_MAX_FAILS = 6;
const HUB_STATS_CACHE_SEC = 300;
const HUB_REGISTER_BURST = 30;   // ลงทะเบียนคนนอกได้ไม่เกินเท่านี้ต่อ 10 นาที (กันสแปม)

/** ============================ ROUTERS ============================ **/

function hubHandleGet_(params) {
  try {
    const api = String(params.api || '');
    if (api === 'hubHome') return hubHome_(params.basis, params.year);
    if (api === 'hubProject') return hubProject_(params.key, params.basis, params.year);
    if (api === 'hubDepartments') return hubDeptMatrix_();
    if (api === 'hubAdminDepts') return { ok: true, departments: hubLoginDepartments_() };
    return { ok: false, error: 'unknown api' };
  } catch (err) {
    return { ok: false, error: (err && err.message) || String(err), code: err && err.code };
  }
}

function hubHandlePost_(req) {
  try {
    switch (String(req.action)) {
      case 'hubJoinStaff': return hubJoinStaff_(req);
      case 'hubStaffStatus': return hubStaffStatus_(req);
      case 'hubDeptJoin': return hubDeptJoin_(req);
      case 'hubRegister': return hubRegister_(req);
      case 'hubMemberStatus': return hubMemberStatus_(req);
      case 'hubMemberJoin': return hubMemberJoin_(req);
      case 'hubMemberLeave': return hubMemberLeave_(req);
      case 'hubMemberDelete': return hubMemberDelete_(req);
      case 'hubRecoverMember': return hubRecoverMember_(req);
      case 'hubAdmin': return hubAdmin_(req);
      default: return { ok: false, error: 'ไม่รู้จักคำสั่ง' };
    }
  } catch (err) {
    return { ok: false, error: (err && err.message) || String(err), code: err && err.code };
  }
}

/** ============================ STORAGE HELPERS ============================ **/

function hubBool_(v) {
  if (v === true || v === 1) return true;
  const s = String(v == null ? '' : v).trim().toUpperCase();
  return s === 'TRUE' || s === '1' || s === 'Y' || s === 'YES' || s === 'ใช่';
}

/** อ่านชีตแบบไม่ผ่านแคช (เห็นสิ่งที่เพิ่งเขียนทันที) → [{_row, Header: value}] */
function hubReadRows_(sh) {
  const values = sh.getDataRange().getValues();
  if (values.length < 2) return [];
  const headers = values[0];
  const out = [];
  for (let i = 1; i < values.length; i++) {
    if (!values[i][0]) continue;
    const o = { _row: i + 1 };
    headers.forEach((h, j) => { o[h] = values[i][j]; });
    out.push(o);
  }
  return out;
}
function hubSetCells_(sh, row, obj) {
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  Object.keys(obj).forEach(k => {
    const c = headers.indexOf(k);
    if (c !== -1) sh.getRange(row, c + 1).setValue(obj[k]);
  });
}
function hubAppend_(sh, headers, obj) {
  const row = sh.getLastRow() + 1;
  sh.getRange(row, 1, 1, headers.length).setValues([headers.map(h => (h in obj ? obj[h] : ''))]);
  return row;
}

function hubProjectsSheet_() { return ensureSheet_(SHEET_HUB_PROJECTS, HUB_PROJECT_HEADERS); }
/** สมาชิกภายนอก/คำขอหน่วยงานเป็นข้อมูลส่วนบุคคล → เก็บในไฟล์ส่วนตัวเดียวกับ PortalUsers (ถ้าตั้งไว้) */
function hubMembersSheet_() {
  const sh = portalSheet_(SHEET_HUB_MEMBERS, HUB_MEMBER_HEADERS);
  return sh;
}
function hubJoinsSheet_() { return portalSheet_(SHEET_HUB_JOINS, HUB_JOIN_HEADERS); }

function hubLock_(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try { return fn(); } finally { lock.releaseLock(); }
}

/** ============================ CATALOG (ผู้ดูแลกำหนด) ============================ **/

/** อ่านรายการโครงการ — สร้าง/เติมแถวที่ขาดจากค่าเริ่มต้นให้อัตโนมัติ */
function hubCatalog_() {
  const sh = hubProjectsSheet_();
  let rows = hubReadRows_(sh);
  const have = {};
  rows.forEach(r => { have[String(r.ProjectKey)] = true; });
  const missing = HUB_DEFAULTS.filter(d => !have[d.key]);
  if (missing.length) {
    const now = new Date();
    missing.forEach((d, i) => {
      hubAppend_(sh, HUB_PROJECT_HEADERS, {
        ProjectKey: d.key, Label: PORTAL_JOIN_PROJECTS[d.key] || d.key, Emoji: d.e, Color: d.c, Tagline: d.t, HowTo: d.how,
        ForStaff: !!d.aud[0], ForDept: !!d.aud[1], ForPublic: !!d.aud[2], ForCompany: !!d.aud[3],
        IsOpen: !!d.open, AutoApprove: !!d.auto, SortOrder: HUB_DEFAULTS.indexOf(d) + 1, UpdatedAt: now, UpdatedBy: 'system'
      });
    });
    rows = hubReadRows_(sh);
  }
  return rows.filter(r => Object.prototype.hasOwnProperty.call(PORTAL_JOIN_PROJECTS, String(r.ProjectKey)))
    .map(r => {
      const key = String(r.ProjectKey);
      const aud = {};
      HUB_AUDIENCES.forEach(a => { aud[a] = hubBool_(r[HUB_AUD_COL[a]]); });
      return {
        key: key, row: r._row, label: String(r.Label || PORTAL_JOIN_PROJECTS[key]), emoji: String(r.Emoji || '🌿'),
        color: /^#[0-9a-f]{3,8}$/i.test(String(r.Color || '')) ? String(r.Color) : '#16a34a',
        tagline: String(r.Tagline || ''), howTo: String(r.HowTo || ''), audiences: aud,
        isOpen: hubBool_(r.IsOpen), autoApprove: hubBool_(r.AutoApprove), sortOrder: Number(r.SortOrder) || 99,
        action: HUB_ACTIONS[key] || ''
      };
    })
    .sort((a, b) => a.sortOrder - b.sortOrder);
}

function hubCatalogMap_() {
  const m = {};
  hubCatalog_().forEach(p => { m[p.key] = p; });
  return m;
}

/** ตรวจว่ากลุ่มนี้สมัครโครงการเหล่านี้ได้ → คืนรายการโครงการที่ผ่าน (โยน error ถ้าไม่มีเลย) */
function hubValidKeys_(keys, audience) {
  const cat = hubCatalogMap_();
  const list = (Array.isArray(keys) ? keys : [keys]).map(String).filter((k, i, a) => k && a.indexOf(k) === i).slice(0, 20);
  if (!list.length) throw portalError_('กรุณาเลือกโครงการอย่างน้อย 1 โครงการ', 'NO_PROJECT');
  const bad = list.filter(k => !cat[k] || !cat[k].isOpen || !cat[k].audiences[audience]);
  if (bad.length) {
    const names = bad.map(k => (cat[k] ? cat[k].label : k)).join(', ');
    throw portalError_('โครงการต่อไปนี้ไม่เปิดรับ' + HUB_AUD_LABEL[audience] + ': ' + names, 'NOT_OPEN');
  }
  return list.map(k => cat[k]);
}

/** ============================ PUBLIC: หน้าแรก / สถิติ ============================ **/

function hubCurrentYear_(basis) { return periodYearOf_(new Date(), basis === 'fiscal' ? 'fiscal' : 'calendar'); }

/** จำนวนสมาชิกที่อนุมัติแล้วของแต่ละโครงการ แยกกลุ่ม (ตัวเลขรวม) */
function hubMemberCounts_() {
  const out = {};
  Object.keys(PORTAL_JOIN_PROJECTS).forEach(k => { out[k] = { staff: 0, dept: 0, public: 0, company: 0 }; });
  portalReadJoins_().filter(j => j.status === PORTAL_JOIN_APPROVED).forEach(j => { if (out[j.projectKey]) out[j.projectKey].staff += 1; });
  const deptSeen = {};
  hubDeptParticipation_().forEach(d => d.projects.forEach(k => {
    if (out[k] && !deptSeen[k + '|' + d.name]) { deptSeen[k + '|' + d.name] = 1; out[k].dept += 1; }
  }));
  hubReadRows_(hubJoinsSheet_()).filter(j => String(j.Status) === HUB_JOIN_APPROVED && (j.MemberType === 'public' || j.MemberType === 'company'))
    .forEach(j => { const k = String(j.ProjectKey); if (out[k]) out[k][j.MemberType] += 1; });
  return out;
}

/** ตัวเลขหลักของทุกโครงการตามปีที่เลือก (แคช 5 นาที) */
function hubStats_(basis, year) {
  const b = basis === 'fiscal' ? 'fiscal' : 'calendar';
  const y = Number(year) || hubCurrentYear_(b);
  const ckey = 'hub_stats_v1_' + b + '_' + y;
  const cache = CacheService.getScriptCache();
  const hit = cache.get(ckey);
  if (hit) { try { return JSON.parse(hit); } catch (e) { /* คำนวณใหม่ */ } }
  const projects = {};
  const yearSet = {};
  let co2Year = 0, co2Total = 0;
  Object.keys(HUB_PERIOD_KEY).forEach(k => {
    let ps = null;
    try { ps = getPeriodSummary(HUB_PERIOD_KEY[k], y, b); } catch (e) { ps = null; }
    if (!ps) { projects[k] = { hasData: false }; return; }
    (ps.years || []).forEach(v => { yearSet[v] = true; });
    const mk = HUB_MAIN_METRIC[k];
    const col = ps.columns.filter(c => c.key === mk)[0] || ps.columns[0];
    co2Year += Number(ps.yearTotal.co2) || 0;
    co2Total += Number(ps.total.co2) || 0;
    projects[k] = {
      hasData: !!ps.hasData,
      main: col ? { key: col.key, label: col.label, unit: col.unit, dec: col.dec, year: Number(ps.yearTotal[col.key]) || 0, total: Number(ps.total[col.key]) || 0 } : null,
      co2Year: round2_(Number(ps.yearTotal.co2) || 0)
    };
  });
  yearSet[hubCurrentYear_(b)] = true;
  const out = {
    basis: b, year: y, years: Object.keys(yearSet).map(Number).sort((p, q) => q - p),
    projects: projects, co2Year: round2_(co2Year), co2Total: round2_(co2Total)
  };
  try { cache.put(ckey, JSON.stringify(out), HUB_STATS_CACHE_SEC); } catch (e) { /* ใหญ่เกินแคช */ }
  return out;
}

/** หน้าแรกของศูนย์ร่วมโครงการ: รายการโครงการ + ใครร่วมได้ + จำนวนสมาชิก + ตัวเลขปีที่เลือก */
function hubHome_(basis, year) {
  const cat = hubCatalog_();
  const counts = hubMemberCounts_();
  const stats = hubStats_(basis, year);
  return {
    ok: true, generatedAt: new Date().toISOString(),
    basis: stats.basis, year: stats.year, years: stats.years, co2Year: stats.co2Year, co2Total: stats.co2Total,
    projects: cat.map(p => Object.assign({}, p, { row: undefined, members: counts[p.key] || {}, stats: stats.projects[p.key] || { hasData: false } }))
  };
}

/** รายละเอียดตัวเลขรายเดือน/รายปีของโครงการเดียว (ตัวเลขรวม ไม่มีข้อมูลบุคคล) */
function hubProject_(key, basis, year) {
  const k = String(key || '');
  if (!HUB_PERIOD_KEY[k]) throw portalError_('ไม่พบโครงการนี้', 'BAD_PROJECT');
  const b = basis === 'fiscal' ? 'fiscal' : 'calendar';
  const ps = getPeriodSummary(HUB_PERIOD_KEY[k], Number(year) || hubCurrentYear_(b), b);
  return Object.assign({ ok: true, projectKey: k, mainKey: HUB_MAIN_METRIC[k] }, ps);
}

/** ============================ หน่วยงาน ============================ **/

/** หน่วยงานที่ร่วมโครงการ (อนุมัติแล้ว): จากชีตรายชื่อหน่วยงานของ Paperless/ตลับหมึก/คัดแยก + คำขอหน่วยงานที่อนุมัติใน HubJoins */
function hubDeptParticipation_() {
  const m = {};
  const norm = s => String(s || '').replace(/\s+/g, ' ').trim();
  const add = (name, key) => {
    const n = norm(name);
    if (!n) return;
    const id = n.toLowerCase();
    (m[id] = m[id] || { name: n, projects: [] });
    if (m[id].projects.indexOf(key) === -1) m[id].projects.push(key);
  };
  Object.keys(HUB_DEPT_SHEET_KEYS).forEach(key => {
    readRowsIfSheet_(projectDeptSheet_(key)).forEach(d => {
      if (!d.DepartmentName || d.DepartmentID === PL_ALL_ID || !shareActive_(d.IsActive)) return;
      add(d.DepartmentName, key);
    });
  });
  hubReadRows_(hubJoinsSheet_()).filter(j => j.MemberType === 'dept' && String(j.Status) === HUB_JOIN_APPROVED)
    .forEach(j => add(j.MemberRef, String(j.ProjectKey)));
  return Object.keys(m).map(k => m[k]).sort((a, b) => a.name.localeCompare(b.name, 'th'));
}

/** ชื่อหน่วยงานทั้งหมดที่รู้จัก (ให้เลือกตอนยื่นคำขอ) */
function hubKnownDeptNames_() {
  const set = {};
  const add = n => { const s = String(n || '').replace(/\s+/g, ' ').trim(); if (s) set[s.toLowerCase()] = s; };
  readRowsIfSheet_(SHEET_DEPARTMENTS).forEach(d => { if (d.IsActive !== false) add(d.DepartmentName); });
  Object.keys(HUB_DEPT_SHEET_KEYS).forEach(key => readRowsIfSheet_(projectDeptSheet_(key)).forEach(d => { if (d.DepartmentID !== PL_ALL_ID) add(d.DepartmentName); }));
  return Object.keys(set).map(k => set[k]).sort((a, b) => a.localeCompare(b, 'th'));
}

/** สาธารณะ: หน่วยงานไหนร่วมโครงการอะไร (ชื่อหน่วยงานเท่านั้น) */
function hubDeptMatrix_() {
  const cat = hubCatalog_().filter(p => p.audiences.dept);
  return {
    ok: true, generatedAt: new Date().toISOString(),
    projects: cat.map(p => ({ key: p.key, label: p.label, emoji: p.emoji, color: p.color })),
    departments: hubDeptParticipation_(), names: hubKnownDeptNames_()
  };
}

/** หน่วยงานสำหรับหน้าเข้าสู่ระบบผู้ดูแล (ชื่อเดียวกับที่แอปหลักแสดงอยู่แล้ว) */
function hubLoginDepartments_() {
  return readRowsIfSheet_(SHEET_DEPARTMENTS).filter(d => d.DepartmentID && d.IsActive !== false)
    .sort((a, b) => (Number(a.SortOrder) || 0) - (Number(b.SortOrder) || 0))
    .map(d => ({ id: d.DepartmentID, name: d.DepartmentName }));
}

/** ============================ JOINS ============================ **/

/** เพิ่ม/เปิดคำขอใหม่ใน HubJoins → 'created' | 'reopened' | 'exists' */
function hubUpsertJoin_(type, ref, project, requestedBy) {
  const sh = hubJoinsSheet_();
  const cur = hubReadRows_(sh).filter(j => j.MemberType === type && String(j.MemberRef).toLowerCase() === String(ref).toLowerCase() && String(j.ProjectKey) === project.key)[0];
  // หน่วยงานต้องให้ผู้ดูแลตั้งค่าในโครงการ (เช่น PIN ของ Paperless) จึงรออนุมัติเสมอ
  const status = project.autoApprove && type !== 'dept' ? HUB_JOIN_APPROVED : HUB_JOIN_PENDING;
  const now = new Date();
  if (!cur) {
    hubAppend_(sh, HUB_JOIN_HEADERS, { JoinID: newId_('HJ'), MemberType: type, MemberRef: ref, ProjectKey: project.key, Status: status,
      CreatedAt: now, RequestedBy: requestedBy || '', ReviewedBy: status === HUB_JOIN_APPROVED ? 'auto' : '', ReviewedAt: status === HUB_JOIN_APPROVED ? now : '' });
    return status === HUB_JOIN_APPROVED ? 'approved' : 'created';
  }
  if (String(cur.Status) === HUB_JOIN_REJECTED) {
    hubSetCells_(sh, cur._row, { Status: status, CreatedAt: now, RequestedBy: requestedBy || '', ReviewedBy: status === HUB_JOIN_APPROVED ? 'auto' : '', ReviewedAt: status === HUB_JOIN_APPROVED ? now : '' });
    return status === HUB_JOIN_APPROVED ? 'approved' : 'created';
  }
  return 'exists';
}

function hubJoinMapOf_(type, ref) {
  const m = {};
  hubReadRows_(hubJoinsSheet_()).filter(j => j.MemberType === type && String(j.MemberRef).toLowerCase() === String(ref).toLowerCase())
    .forEach(j => { m[String(j.ProjectKey)] = String(j.Status); });
  return m;
}

/**
 * เจ้าหน้าที่สมัครหลายโครงการในครั้งเดียว — เขียนลง PortalJoinRequests (ใช้กับกฎ "สมาชิกโครงการ" ของฟอร์มเดิมได้ทันที)
 * โครงการที่ตั้ง "อนุมัติอัตโนมัติ" → เป็นสมาชิกทันที (มีผลเมื่อบัญชีผ่านการอนุมัติแล้ว เพราะ portalAssertMember_ ตรวจทั้งสองอย่าง)
 * บัญชีที่ยังรออนุมัติก็สมัครล่วงหน้าได้ — ผู้ดูแลอนุมัติบัญชีแล้วใช้ได้เลยไม่ต้องสมัครซ้ำ
 */
function hubJoinStaff_(req) {
  const user = portalRequireUser_(req.token);
  const projects = hubValidKeys_(req.projectKeys, 'staff');
  const result = {};
  let needReview = false;
  hubLock_(() => {
    const sh = portalJoinsSheet_();
    const joins = portalReadJoins_().filter(j => j.staffCode === user.staffCode);
    const now = new Date();
    projects.forEach(p => {
      const ex = joins.filter(j => j.projectKey === p.key)[0];
      const st = p.autoApprove ? PORTAL_JOIN_APPROVED : PORTAL_JOIN_PENDING;
      const rev = st === PORTAL_JOIN_APPROVED ? ['auto', now] : ['', ''];
      if (!ex) {
        const row = sh.getLastRow() + 1;
        sh.getRange(row, 2).setNumberFormat('@');
        sh.getRange(row, 1, 1, PORTAL_JOIN_HEADERS.length).setValues([[newId_('JR'), user.staffCode, p.key, st, now, rev[0], rev[1]]]);
        result[p.key] = st;
      } else if (ex.status === PORTAL_JOIN_REJECTED || (ex.status === PORTAL_JOIN_PENDING && p.autoApprove)) {
        sh.getRange(ex.row, 4, 1, 4).setValues([[st, now, rev[0], rev[1]]]);
        result[p.key] = st;
      } else result[p.key] = ex.status;
      if (result[p.key] === PORTAL_JOIN_PENDING) needReview = true;
    });
    portalAudit_(maskStaffCode_(user.staffCode), 'hubJoinStaff', maskStaffCode_(user.staffCode), projects.map(p => p.key).join(','));
  });
  if (needReview || user.status === PORTAL_STATUS_PENDING) hubNotifyAdmins_();
  return { ok: true, result: result, joins: portalJoinMap_(user.staffCode), accountPending: user.status === PORTAL_STATUS_PENDING };
}

/** สถานะของเจ้าหน้าที่ที่เข้าสู่ระบบ: โครงการที่ร่วม + หน่วยงานที่เคยยื่นคำขอแทน */
function hubStaffStatus_(req) {
  const user = portalRequireUser_(req.token);
  const masked = maskStaffCode_(user.staffCode);
  const deptReqs = hubReadRows_(hubJoinsSheet_()).filter(j => j.MemberType === 'dept' && String(j.RequestedBy) === masked)
    .map(j => ({ dept: String(j.MemberRef), projectKey: String(j.ProjectKey), status: String(j.Status) }));
  return {
    ok: true, user: portalPublicUser_(user, String(req.noticeVersion || '')),
    joins: portalJoinMap_(user.staffCode), deptRequests: deptReqs,
    projects: user.status === PORTAL_STATUS_APPROVED ? portalParticipation_(user.staffCode, user.name) : []
  };
}

/** เจ้าหน้าที่ยื่นคำขอให้ "หน่วยงาน" ร่วมหลายโครงการ → ผู้ดูแลอนุมัติ (แล้วระบบเพิ่มชื่อหน่วยงานในโครงการให้) */
function hubDeptJoin_(req) {
  const user = portalRequireUser_(req.token);
  const dept = String(req.deptName || '').replace(/\s+/g, ' ').trim();
  if (dept.length < 2 || dept.length > 120) throw portalError_('กรุณาระบุชื่อหน่วยงาน (2–120 ตัวอักษร)', 'BAD_DEPT');
  const projects = hubValidKeys_(req.projectKeys, 'dept');
  const result = {};
  hubLock_(() => {
    projects.forEach(p => { result[p.key] = hubUpsertJoin_('dept', dept, p, maskStaffCode_(user.staffCode)); });
    const note = String(req.note || '').slice(0, 300);
    if (note) {
      const sh = hubJoinsSheet_();
      hubReadRows_(sh).filter(j => j.MemberType === 'dept' && String(j.MemberRef).toLowerCase() === dept.toLowerCase() && String(j.Status) === HUB_JOIN_PENDING)
        .forEach(j => hubSetCells_(sh, j._row, { Note: note }));
    }
    portalAudit_(maskStaffCode_(user.staffCode), 'hubDeptJoin', dept, projects.map(p => p.key).join(','));
  });
  hubNotifyAdmins_();
  return { ok: true, dept: dept, result: result, joins: hubJoinMapOf_('dept', dept) };
}

/** ============================ ประชาชน / บริษัท ============================ **/

function hubNormPhone_(v) {
  let s = String(v == null ? '' : v).replace(/[^\d+]/g, '');
  if (s.indexOf('+66') === 0) s = '0' + s.slice(3);
  s = s.replace(/\D/g, '');
  return /^0\d{8,9}$/.test(s) ? s : '';
}
function hubMaskPhone_(p) { const s = String(p || ''); return s.length < 4 ? '•••' : s.slice(0, 2) + '•'.repeat(Math.max(0, s.length - 4)) + s.slice(-2); }
function hubMaskEmail_(e) {
  const s = String(e || ''); const at = s.indexOf('@');
  return at < 1 ? '' : s.charAt(0) + '•••' + s.slice(at);
}
function hubNewMemberId_(type) {
  const hex = (Utilities.getUuid() + Utilities.getUuid()).replace(/-/g, '');
  let id = '';
  for (let i = 0; i < 6; i++) id += HUB_ID_CHARS.charAt(parseInt(hex.substr(i * 4, 4), 16) % HUB_ID_CHARS.length);
  return (type === 'company' ? 'C' : 'P') + '-' + id;
}
function hubNormMemberId_(v) { return String(v || '').toUpperCase().replace(/[^A-Z0-9]/g, '').replace(/^([PC])/, '$1-'); }

function hubFindMember_(id) {
  const mid = hubNormMemberId_(id);
  return hubReadRows_(hubMembersSheet_()).filter(m => String(m.MemberID) === mid)[0] || null;
}

function hubPublicMember_(m) {
  return {
    memberId: String(m.MemberID), type: String(m.MemberType),
    nameMasked: maskPersonName_(m.Name), orgName: String(m.OrgName || ''),
    phoneMasked: hubMaskPhone_(m.Phone), emailMasked: hubMaskEmail_(m.Email),
    firstName: portalFirstName_(m.Name)
  };
}

/** ยืนยันสมาชิกภายนอกด้วย รหัสสมาชิก + เบอร์โทร (จำกัดจำนวนครั้งที่ผิด) */
function hubVerifyMember_(memberId, phone) {
  const mid = hubNormMemberId_(memberId);
  if (!/^[PC]-[A-Z0-9]{6}$/.test(mid)) throw portalError_('รหัสสมาชิกไม่ถูกต้อง (เช่น P-7KQ2MX)', 'BAD_MEMBER');
  const cache = CacheService.getScriptCache();
  const fk = 'hub_mfail_' + mid;
  const fails = Number(cache.get(fk) || 0);
  if (fails >= HUB_MAX_FAILS) throw portalError_('ตรวจสอบไม่สำเร็จหลายครั้ง กรุณารอ 10 นาทีแล้วลองใหม่', 'LOCKED');
  const m = hubFindMember_(mid);
  const ok = !!m && portalSafeEq_(hubNormPhone_(m.Phone), hubNormPhone_(phone));
  if (!ok) {
    cache.put(fk, String(fails + 1), 600);
    throw portalError_('รหัสสมาชิกหรือเบอร์โทรไม่ถูกต้อง', 'BAD_MEMBER');
  }
  cache.remove(fk);
  return m;
}

/**
 * ลงทะเบียนประชาชน/บริษัท (ครั้งเดียว ไม่ต้องตั้งรหัสผ่าน) + เลือกโครงการได้หลายโครงการ
 * req: { type:'public'|'company', name, orgName, phone, email, ack:true, noticeVersion, projectKeys:[] }
 */
function hubRegister_(req) {
  const type = req.type === 'company' ? 'company' : 'public';
  const name = String(req.name || '').replace(/\s+/g, ' ').trim().slice(0, 100);
  const org = String(req.orgName || '').replace(/\s+/g, ' ').trim().slice(0, 150);
  const phone = hubNormPhone_(req.phone);
  const email = String(req.email || '').trim().slice(0, 120);
  if (name.length < 2) throw portalError_(type === 'company' ? 'กรุณากรอกชื่อผู้ติดต่อ' : 'กรุณากรอกชื่อ-นามสกุล', 'BAD_NAME');
  if (type === 'company' && org.length < 2) throw portalError_('กรุณากรอกชื่อบริษัท/องค์กร', 'BAD_ORG');
  if (!phone) throw portalError_('เบอร์โทรไม่ถูกต้อง (ตัวเลข 9–10 หลัก ขึ้นต้นด้วย 0)', 'BAD_PHONE');
  if (email && !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) throw portalError_('รูปแบบอีเมลไม่ถูกต้อง', 'BAD_EMAIL');
  if (req.ack !== true) throw portalError_('กรุณารับทราบประกาศความเป็นส่วนตัวก่อนลงทะเบียน', 'NO_ACK');
  const projects = hubValidKeys_(req.projectKeys, type);

  const cache = CacheService.getScriptCache();
  const burst = Number(cache.get('hub_reg_burst') || 0);
  if (burst >= HUB_REGISTER_BURST) throw portalError_('มีการลงทะเบียนจำนวนมากในช่วงสั้นๆ กรุณารอสักครู่แล้วลองใหม่', 'LOCKED');
  cache.put('hub_reg_burst', String(burst + 1), 600);

  let member;
  const result = {};
  hubLock_(() => {
    const sh = hubMembersSheet_();
    const dup = hubReadRows_(sh).filter(m => m.MemberType === type && hubNormPhone_(m.Phone) === phone)[0];
    // ไม่บอกรหัสเดิมกลับไป (กันคนรู้เบอร์ผู้อื่นแล้วสวมรอย) — ให้ใช้ "ลืมรหัสสมาชิก" ส่งไปทางอีเมลแทน
    if (dup) throw portalError_('เบอร์นี้ลงทะเบียนไว้แล้ว — ใช้รหัสสมาชิกเดิมกด "เข้าร่วมเพิ่ม" หรือกด "ลืมรหัสสมาชิก" ให้ระบบส่งรหัสทางอีเมล', 'HAS_MEMBER');
    const ids = {};
    hubReadRows_(sh).forEach(m => { ids[String(m.MemberID)] = 1; });
    let id = hubNewMemberId_(type);
    for (let i = 0; ids[id] && i < 10; i++) id = hubNewMemberId_(type);
    const now = new Date();
    const row = hubAppend_(sh, HUB_MEMBER_HEADERS, {
      MemberID: id, MemberType: type, Name: name, OrgName: org, Phone: "'" + phone, Email: email,
      CreatedAt: now, ConsentAt: now, NoticeVersion: String(req.noticeVersion || '').slice(0, 40), LastSeenAt: now
    });
    member = { _row: row, MemberID: id, MemberType: type, Name: name, OrgName: org, Phone: phone, Email: email };
    projects.forEach(p => { result[p.key] = hubUpsertJoin_(type, id, p, ''); });
    portalAudit_(id, 'hubRegister', id, type + ':' + projects.map(p => p.key).join(','));
  });
  if (email) hubMailMember_(member, 'ยืนยันการลงทะเบียนร่วมโครงการสิ่งแวดล้อม', projects.map(p => p.emoji + ' ' + p.label));
  if (Object.keys(result).some(k => result[k] === 'created')) hubNotifyAdmins_();
  return { ok: true, member: hubPublicMember_(member), result: result, joins: hubJoinMapOf_(type, member.MemberID) };
}

function hubMemberStatus_(req) {
  const m = hubVerifyMember_(req.memberId, req.phone);
  hubSetCells_(hubMembersSheet_(), m._row, { LastSeenAt: new Date() });
  return {
    ok: true, member: hubPublicMember_(m), joins: hubJoinMapOf_(String(m.MemberType), String(m.MemberID)),
    projects: portalParticipation_('', String(m.Name || ''))
  };
}

function hubMemberJoin_(req) {
  const m = hubVerifyMember_(req.memberId, req.phone);
  const type = String(m.MemberType) === 'company' ? 'company' : 'public';
  const projects = hubValidKeys_(req.projectKeys, type);
  const result = {};
  hubLock_(() => { projects.forEach(p => { result[p.key] = hubUpsertJoin_(type, String(m.MemberID), p, ''); }); });
  portalAudit_(String(m.MemberID), 'hubMemberJoin', String(m.MemberID), projects.map(p => p.key).join(','));
  if (Object.keys(result).some(k => result[k] === 'created')) hubNotifyAdmins_();
  return { ok: true, member: hubPublicMember_(m), result: result, joins: hubJoinMapOf_(type, String(m.MemberID)) };
}

/** เลิกเข้าร่วม/ยกเลิกคำขอ (ลบแถวคำขอของโครงการนั้น) */
function hubMemberLeave_(req) {
  const m = hubVerifyMember_(req.memberId, req.phone);
  const key = String(req.projectKey || '');
  hubLock_(() => {
    const sh = hubJoinsSheet_();
    hubReadRows_(sh).filter(j => j.MemberType === m.MemberType && String(j.MemberRef) === String(m.MemberID) && String(j.ProjectKey) === key)
      .sort((a, b) => b._row - a._row).forEach(j => sh.deleteRow(j._row));
  });
  portalAudit_(String(m.MemberID), 'hubMemberLeave', String(m.MemberID), key);
  return { ok: true, joins: hubJoinMapOf_(String(m.MemberType), String(m.MemberID)) };
}

/** PDPA: ลบข้อมูลสมาชิกภายนอกและคำขอทั้งหมดของตนเอง */
function hubMemberDelete_(req) {
  const m = hubVerifyMember_(req.memberId, req.phone);
  hubLock_(() => {
    const js = hubJoinsSheet_();
    hubReadRows_(js).filter(j => String(j.MemberRef) === String(m.MemberID)).sort((a, b) => b._row - a._row).forEach(j => js.deleteRow(j._row));
    const fresh = hubFindMember_(m.MemberID);
    if (fresh) hubMembersSheet_().deleteRow(fresh._row);
  });
  portalAudit_('(สมาชิกภายนอกที่ลบข้อมูล)', 'hubMemberDelete', '', '');
  return { ok: true, deleted: true };
}

/** ลืมรหัสสมาชิก: ส่งรหัสไปที่อีเมลที่ลงทะเบียนไว้ (ตอบเหมือนกันทุกกรณี ไม่บอกว่ามีสมาชิกหรือไม่) */
function hubRecoverMember_(req) {
  const phone = hubNormPhone_(req.phone);
  const email = String(req.email || '').trim().toLowerCase();
  const msg = { ok: true, message: 'ถ้าข้อมูลตรงกับที่ลงทะเบียนไว้ ระบบจะส่งรหัสสมาชิกไปที่อีเมลนั้นภายในไม่กี่นาที' };
  if (!phone || !email) return msg;
  const cache = CacheService.getScriptCache();
  const k = 'hub_recover_' + phone;
  if (cache.get(k)) return msg;
  cache.put(k, '1', 600);
  const m = hubReadRows_(hubMembersSheet_()).filter(x => hubNormPhone_(x.Phone) === phone && String(x.Email || '').trim().toLowerCase() === email)[0];
  if (m) hubMailMember_(m, 'รหัสสมาชิกโครงการสิ่งแวดล้อมของท่าน', null);
  return msg;
}

function hubMailMember_(m, subject, projectLines) {
  try {
    if (!m.Email) return;
    MailApp.sendEmail({
      to: String(m.Email), subject: '[โครงการสิ่งแวดล้อม] ' + subject,
      htmlBody: '<p>เรียน ' + hubEscape_(portalFirstName_(m.Name) || m.Name) + '</p>' +
        '<p>รหัสสมาชิกของท่านคือ <b style="font-size:18px;letter-spacing:2px">' + hubEscape_(m.MemberID) + '</b></p>' +
        (projectLines && projectLines.length ? '<p>โครงการที่เลือก:<br>' + projectLines.map(hubEscape_).join('<br>') + '</p>' : '') +
        '<p>ใช้รหัสนี้คู่กับเบอร์โทรที่ลงทะเบียน เพื่อดูสถานะหรือเลือกโครงการเพิ่มที่หน้า "ร่วมโครงการ"</p>' +
        '<p style="color:#666;font-size:12px">หากท่านไม่ได้ลงทะเบียน โปรดเพิกเฉยอีเมลนี้</p>'
    });
  } catch (err) { /* ส่งอีเมลไม่ได้ไม่กระทบการลงทะเบียน */ }
}
function hubEscape_(s) { return String(s == null ? '' : s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' })[c]); }

/** ============================ แจ้งเตือนผู้ดูแล ============================ **/

function hubPendingCount_() {
  return hubReadRows_(hubJoinsSheet_()).filter(j => String(j.Status) === HUB_JOIN_PENDING).length;
}

function hubNotifyAdmins_() {
  try {
    const cache = CacheService.getScriptCache();
    if (cache.get('hub_notify_gap')) return false;
    const to = getInquiryRecipients_();
    if (!to.length) return false;
    cache.put('hub_notify_gap', '1', PORTAL_NOTIFY_GAP_SEC);
    const a = portalPendingCount_(), s = portalPendingJoinCount_(), h = hubPendingCount_();
    MailApp.sendEmail({
      to: to.join(','),
      subject: '[ระบบสิ่งแวดล้อม] รอตรวจสอบ ' + (a + s + h) + ' รายการ (ศูนย์ร่วมโครงการ)',
      htmlBody: '<p>รายการที่รอผู้ดูแลระบบตรวจสอบ:</p><ul>' +
        '<li>บัญชีเจ้าหน้าที่ใหม่: <b>' + a + '</b></li><li>คำขอเข้าร่วมโครงการของเจ้าหน้าที่: <b>' + s + '</b></li>' +
        '<li>คำขอของหน่วยงาน / ประชาชน / บริษัท: <b>' + h + '</b></li></ul>' +
        '<p>วิธีตรวจ: หน้าเว็บ "ร่วมโครงการ" → ปุ่ม <b>ผู้ดูแล</b> (มุมขวาบน) → เข้าสู่ระบบด้วยหน่วยงานผู้ดูแล + PIN → แท็บ "รออนุมัติ"</p>' +
        '<p style="color:#666;font-size:12px">อีเมลนี้ไม่มีข้อมูลส่วนบุคคลของผู้สมัคร เพื่อความปลอดภัยตาม PDPA</p>'
    });
    return true;
  } catch (err) { return false; }
}

/** ============================ ผู้ดูแลระบบ ============================ **/

/** ตรวจ PIN ผู้ดูแล (จำกัดจำนวนครั้งที่ผิดต่อหน่วยงาน) */
function hubAssertAdmin_(departmentId, pin) {
  const cache = CacheService.getScriptCache();
  const fk = 'hub_afail_' + String(departmentId || '');
  const fails = Number(cache.get(fk) || 0);
  if (fails >= 5) throw portalError_('ใส่ PIN ผิดหลายครั้ง กรุณารอ 10 นาที', 'LOCKED');
  try {
    const who = assertAuthorized_(String(departmentId || ''), String(pin || ''));
    cache.remove(fk);
    return who;
  } catch (err) {
    cache.put(fk, String(fails + 1), 600);
    throw portalError_((err && err.message) || 'ไม่ได้รับอนุญาต', 'ADMIN_AUTH');
  }
}

/** คำขอใน HubJoins ที่รออนุมัติ พร้อมข้อมูลผู้ขอ (ผู้ดูแลเห็นข้อมูลเต็ม) */
function hubPendingList_() {
  const members = {};
  hubReadRows_(hubMembersSheet_()).forEach(m => { members[String(m.MemberID)] = m; });
  const cat = hubCatalogMap_();
  return hubReadRows_(hubJoinsSheet_()).filter(j => String(j.Status) === HUB_JOIN_PENDING).map(j => {
    const m = members[String(j.MemberRef)];
    return {
      joinId: String(j.JoinID), type: String(j.MemberType), ref: String(j.MemberRef),
      name: j.MemberType === 'dept' ? String(j.MemberRef) : (m ? String(m.Name) : '(ไม่พบสมาชิก)'),
      orgName: m ? String(m.OrgName || '') : '', phone: m ? hubNormPhone_(m.Phone) : '', email: m ? String(m.Email || '') : '',
      projectKey: String(j.ProjectKey), projectLabel: cat[String(j.ProjectKey)] ? cat[String(j.ProjectKey)].label : String(j.ProjectKey),
      requestedBy: String(j.RequestedBy || ''), note: String(j.Note || ''), createdAt: portalDateTime_(j.CreatedAt),
      needsPin: j.MemberType === 'dept' && String(j.ProjectKey) === 'paperless'
    };
  }).sort((a, b) => (a.createdAt < b.createdAt ? -1 : 1));
}

function hubInbox_() {
  return { accounts: portalPendingList_(), staffJoins: portalJoinPendingList_(), hubJoins: hubPendingList_() };
}

/** ผู้ดูแลอนุมัติ/ไม่อนุมัติคำขอใน HubJoins — หน่วยงาน: เพิ่มชื่อในชีตรายชื่อหน่วยงานของโครงการให้ด้วย */
function hubReviewJoin_(who, departmentId, pin, joinId, approve, deptPin) {
  const find = () => {
    const j = hubReadRows_(hubJoinsSheet_()).filter(x => String(x.JoinID) === String(joinId))[0];
    if (!j) throw portalError_('ไม่พบคำขอนี้ (อาจถูกจัดการไปแล้ว)', 'NOT_FOUND');
    if (String(j.Status) !== HUB_JOIN_PENDING) throw portalError_('คำขอนี้ได้รับการพิจารณาแล้ว', 'DONE');
    return j;
  };
  const first = find();
  const key = String(first.ProjectKey);
  if (approve === true && first.MemberType === 'dept' && HUB_DEPT_SHEET_KEYS[key]) {
    // ใช้ฟังก์ชันเดิมของระบบ (เพิ่มใหม่/เปิดใช้ชื่อเดิม · Paperless ต้องตั้ง PIN 4–8 หลัก) — มี Lock ของตัวเอง จึงเรียกนอก hubLock_
    addProjectDepartment(departmentId, pin, key, String(first.MemberRef), deptPin, []);
  }
  return hubLock_(() => {
    const j = find();
    hubSetCells_(hubJoinsSheet_(), j._row, { Status: approve === true ? HUB_JOIN_APPROVED : HUB_JOIN_REJECTED, ReviewedBy: who.departmentName, ReviewedAt: new Date() });
    portalAudit_(who.departmentName, approve === true ? 'hubApprove' : 'hubReject', String(j.MemberRef), key);
    return true;
  });
}

/** บันทึกการตั้งค่าโครงการจากหน้าเว็บ: items = [{key, label, tagline, howTo, forStaff, forDept, forPublic, forCompany, isOpen, autoApprove, sortOrder}] */
function hubSaveCatalog_(who, items) {
  if (!Array.isArray(items)) throw portalError_('ข้อมูลไม่ถูกต้อง', 'BAD_INPUT');
  return hubLock_(() => {
    const sh = hubProjectsSheet_();
    const cat = hubCatalogMap_();
    const now = new Date();
    let n = 0;
    items.forEach(it => {
      const p = cat[String(it && it.key)];
      if (!p) return;
      const txt = (v, max, dflt) => { const s = String(v == null ? dflt : v).replace(/\s+/g, ' ').trim(); return s.slice(0, max); };
      hubSetCells_(sh, p.row, {
        Label: txt(it.label, 60, p.label) || p.label, Tagline: txt(it.tagline, 200, p.tagline), HowTo: txt(it.howTo, 300, p.howTo),
        ForStaff: it.forStaff === true, ForDept: it.forDept === true, ForPublic: it.forPublic === true, ForCompany: it.forCompany === true,
        IsOpen: it.isOpen === true, AutoApprove: it.autoApprove === true,
        SortOrder: Math.max(1, Math.min(99, Number(it.sortOrder) || p.sortOrder)), UpdatedAt: now, UpdatedBy: who.departmentName
      });
      n += 1;
    });
    portalAudit_(who.departmentName, 'hubSaveCatalog', '', 'items=' + n);
    return n;
  });
}

function hubMemberList_() {
  const joins = {};
  hubReadRows_(hubJoinsSheet_()).forEach(j => { (joins[String(j.MemberRef)] = joins[String(j.MemberRef)] || []).push(String(j.ProjectKey) + ':' + String(j.Status)); });
  return hubReadRows_(hubMembersSheet_()).map(m => ({
    memberId: String(m.MemberID), type: String(m.MemberType), name: String(m.Name), orgName: String(m.OrgName || ''),
    phone: hubNormPhone_(m.Phone), email: String(m.Email || ''), createdAt: portalDateTime_(m.CreatedAt),
    lastSeenAt: portalDateTime_(m.LastSeenAt), joins: joins[String(m.MemberID)] || []
  })).sort((a, b) => (a.createdAt < b.createdAt ? 1 : -1));
}

/**
 * คำสั่งผู้ดูแลทั้งหมด (ผ่าน POST — PIN ไม่อยู่ใน URL)
 * op: 'inbox' | 'catalog' | 'saveCatalog' | 'reviewAccount' | 'reviewStaffJoin' | 'reviewHubJoin' | 'members'
 */
function hubAdmin_(req) {
  const departmentId = String(req.departmentId || ''), pin = String(req.pin || '');
  const who = hubAssertAdmin_(departmentId, pin);
  const op = String(req.op || '');
  if (op === 'login') return { ok: true, admin: who.departmentName, counts: { accounts: portalPendingCount_(), staffJoins: portalPendingJoinCount_(), hubJoins: hubPendingCount_() } };
  if (op === 'inbox') { portalAudit_(who.departmentName, 'hubViewInbox', '', ''); return Object.assign({ ok: true }, hubInbox_()); }
  if (op === 'catalog') return { ok: true, projects: hubCatalog_().map(p => Object.assign({}, p, { row: undefined })) };
  if (op === 'saveCatalog') {
    const n = hubSaveCatalog_(who, req.items);
    return { ok: true, saved: n, projects: hubCatalog_().map(p => Object.assign({}, p, { row: undefined })) };
  }
  if (op === 'reviewAccount') { reviewPortalUser(departmentId, pin, req.staffCode, req.approve === true); return Object.assign({ ok: true }, hubInbox_()); }
  if (op === 'reviewStaffJoin') { reviewPortalJoin(departmentId, pin, req.requestId, req.approve === true); return Object.assign({ ok: true }, hubInbox_()); }
  if (op === 'reviewHubJoin') { hubReviewJoin_(who, departmentId, pin, req.joinId, req.approve === true, req.deptPin); return Object.assign({ ok: true }, hubInbox_()); }
  if (op === 'members') { portalAudit_(who.departmentName, 'hubViewMembers', '', ''); return { ok: true, members: hubMemberList_() }; }
  throw portalError_('ไม่รู้จักคำสั่งผู้ดูแล', 'BAD_OP');
}
