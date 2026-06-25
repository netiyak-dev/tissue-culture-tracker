/**
 * backend/Code.gs
 * Google Apps Script — ทำหน้าที่เป็น Backend API + ฐานข้อมูล (Google Sheet) + ตัวส่งแจ้งเตือน Line
 *
 * วิธีติดตั้ง: อ่าน README.md ในโฟลเดอร์โปรเจกต์ทั้งหมดก่อนเริ่ม
 *
 * ก่อนใช้งาน ต้องตั้งค่า Script Properties (Project Settings > Script Properties):
 *   SPREADSHEET_ID              -> ID ของ Google Sheet ที่จะใช้เป็นฐานข้อมูล (จำเป็น)
 *   LINE_CHANNEL_ACCESS_TOKEN   -> Channel access token ของ LINE Messaging API (จำเป็น)
 *   ADMIN_LINE_USER_ID          -> LINE user ID ของคนแรกที่จะเป็นแอดมินได้ (ใช้ตอนเริ่มระบบครั้งแรกเท่านั้น
 *     เพื่อปลดล็อกหน้า "สิทธิ์ผู้ใช้" ในตั้งค่า — หลังจากนั้น Admin คนนี้ไปตั้งสิทธิ์ admin ให้ตัวเอง/คนอื่นในชีต
 *     LineUsers ต่อได้เลย ไม่ต้องพึ่ง Property นี้อีก ถ้าไม่ใส่ จะไม่มีใครเป็นแอดมินได้จนกว่าจะไปแก้คอลัมน์ role ในชีตตรงๆ)
 *     หา LINE user ID ของตัวเองได้จากชีต "LineUsers" หลังจากเปิด LIFF ผ่าน Line ไปแล้วอย่างน้อย 1 ครั้ง
 *
 * แล้วรันฟังก์ชัน setupSheets() หนึ่งครั้งจากใน Apps Script editor เพื่อสร้างชีตและหัวคอลัมน์ให้ครบ
 * ถ้าเคยรัน setupSheets() ไปแล้วก่อนหน้า (มีชีตอยู่แล้ว) ให้รัน migrateSheets() แทน เพื่อเพิ่มคอลัมน์ใหม่โดยไม่ลบข้อมูลเดิม
 */

const STAGE_NAMES = ["ขยาย", "กระตุ้นราก", "พร้อมออกปลูก"];
const MEDIA_TYPES = ["TIB", "SS"];
const LABS = ["KA", "NJ", "CC"];

// ===================================================================
// เข้าถึง Spreadsheet
// ===================================================================
function getSS() {
  const id = PropertiesService.getScriptProperties().getProperty("SPREADSHEET_ID");
  if (!id) throw new Error("ยังไม่ได้ตั้งค่า SPREADSHEET_ID ใน Script Properties");
  return SpreadsheetApp.openById(id);
}

const SHEET_SCHEMAS = {
  PlantTypes: [
    "plant_id", "plant_name",
    "stage1_rounds", "stage1_days", "stage1_recipe_tib", "stage1_recipe_ss",
    "stage1_multiplier_tib", "stage1_multiplier_ss",
    "stage2_days", "stage2_recipe_tib", "stage2_recipe_ss",
    "stage3_recipe_tib", "stage3_recipe_ss",
    "active",
  ],
  Storage: ["type", "value", "active"],
  Transfers: [
    "record_id", "plant_id", "plant_name", "lot_no", "lab", "round_no", "transfer_date", "quantity",
    "bottle_count", "plants_per_bottle",
    "media_type", "bottle_no", "shelf", "sub_shelf",
    "stage", "recipe_used", "next_transfer_date", "recorded_by", "created_at",
  ],
  LineUsers: ["line_user_id", "display_name", "role", "lab", "active", "approved"],
  // ค่าพารามิเตอร์ระบบ SS/TIB ระดับระบบ (ไม่ใช่ต่อชนิดพืช) — มีแถวเดียวเสมอ (config_id = "global")
  // ใช้ทั้งคำนวณพยากรณ์อัตโนมัติที่ Dashboard และเป็นค่าเริ่มต้นของหน้า "จำลองรอบการผลิต" (แก้ได้ที่ตั้งค่า > ระบบ SS/TIB)
  SystemConfig: [
    "config_id", "cycle_days",
    "ss_bottle_limit", "ss_pieces_per_bottle", "ss_multiplication_factor",
    "tib_bottle_limit", "tib_pieces_per_bottle", "tib_multiplication_factor",
  ],
};

const SYSTEM_CONFIG_DEFAULTS = {
  cycle_days: 15,
  ss_bottle_limit: 2000, ss_pieces_per_bottle: 8, ss_multiplication_factor: 5,
  tib_bottle_limit: 200, tib_pieces_per_bottle: 20, tib_multiplication_factor: 20,
};

function seedSystemConfigIfEmpty() {
  if (getSheetData("SystemConfig").length === 0) {
    appendRow("SystemConfig", { config_id: "global", ...SYSTEM_CONFIG_DEFAULTS });
  }
}

/** รันครั้งเดียวตอนติดตั้งระบบครั้งแรก: สร้างชีต + หัวคอลัมน์ (ไม่ลบของเดิมถ้ามีอยู่แล้ว) */
function setupSheets() {
  const ss = getSS();
  Object.keys(SHEET_SCHEMAS).forEach(name => {
    let sheet = ss.getSheetByName(name);
    if (!sheet) {
      sheet = ss.insertSheet(name);
    }
    if (sheet.getRange(1, 1).getValue() === "") {
      sheet.getRange(1, 1, 1, SHEET_SCHEMAS[name].length).setValues([SHEET_SCHEMAS[name]]);
      sheet.setFrozenRows(1);
    }
  });
  const storageRows = getSheetData("Storage");
  if (storageRows.length === 0) {
    ["1", "2", "3", "4"].forEach(v => appendRow("Storage", { type: "shelf", value: v, active: true }));
    ["A", "B", "C"].forEach(v => appendRow("Storage", { type: "sub_shelf", value: v, active: true }));
  }
  seedSystemConfigIfEmpty();
  Logger.log("ติดตั้งชีตเรียบร้อย");
}

/**
 * รันเมื่อมีการอัปเดตโค้ดที่เพิ่มคอลัมน์ใหม่ (เช่นตอนนี้) แล้วชีตเดิมมีข้อมูลอยู่แล้ว
 * จะเติมคอลัมน์ที่ขาดไปต่อท้ายคอลัมน์เดิม โดยไม่แก้ไข/ลบข้อมูลที่มีอยู่
 * ปลอดภัย รันซ้ำได้หลายครั้งไม่มีผลเสีย — ถ้าชีตที่เพิ่งเพิ่มเข้าระบบ (เช่น SystemConfig) ยังไม่มีอยู่เลย จะสร้างให้ด้วย
 */
function migrateSheets() {
  const ss = getSS();
  Object.keys(SHEET_SCHEMAS).forEach(name => {
    let sheet = ss.getSheetByName(name);
    if (!sheet) {
      sheet = ss.insertSheet(name);
      sheet.getRange(1, 1, 1, SHEET_SCHEMAS[name].length).setValues([SHEET_SCHEMAS[name]]);
      sheet.setFrozenRows(1);
      Logger.log(`สร้างชีตใหม่: ${name}`);
      return;
    }
    const lastCol = Math.max(sheet.getLastColumn(), 1);
    const existingHeaders = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    const missing = SHEET_SCHEMAS[name].filter(h => existingHeaders.indexOf(h) === -1);
    if (missing.length > 0) {
      sheet.getRange(1, lastCol + 1, 1, missing.length).setValues([missing]);
      Logger.log(`เพิ่มคอลัมน์ใหม่ในชีต ${name}: ${missing.join(", ")}`);
    }
  });
  seedSystemConfigIfEmpty();
  Logger.log("ย้ายโครงสร้างชีตเรียบร้อย");
}

// ===================================================================
// อ่าน/เขียนชีตแบบเป็น object array (อิงหัวคอลัมน์จริงในชีต ไม่ใช่ลำดับคงที่
// เพื่อให้ migrateSheets() เพิ่มคอลัมน์ใหม่ต่อท้ายได้โดยไม่พังโครงสร้างเดิม)
// ===================================================================
function getSheet(name) {
  const sheet = getSS().getSheetByName(name);
  if (!sheet) throw new Error(`ไม่พบชีต ${name} — รัน setupSheets() ก่อน`);
  return sheet;
}

function getHeaders(sheet) {
  const lastCol = sheet.getLastColumn();
  if (lastCol === 0) return [];
  return sheet.getRange(1, 1, 1, lastCol).getValues()[0];
}

function getSheetData(name) {
  const sheet = getSheet(name);
  const range = sheet.getDataRange().getValues();
  if (range.length < 2) return [];
  const headers = range[0];
  return range.slice(1)
    .filter(row => row.some(cell => cell !== ""))
    .map((row, idx) => {
      const obj = { _row: idx + 2 };
      headers.forEach((h, i) => (obj[h] = row[i]));
      return obj;
    });
}

function appendRow(name, obj) {
  const sheet = getSheet(name);
  const headers = getHeaders(sheet);
  const row = headers.map(h => (obj[h] !== undefined ? obj[h] : ""));
  sheet.appendRow(row);
  return obj;
}

function updateRow(name, rowIndex, obj) {
  const sheet = getSheet(name);
  const headers = getHeaders(sheet);
  const row = headers.map(h => (obj[h] !== undefined ? obj[h] : ""));
  sheet.getRange(rowIndex, 1, 1, headers.length).setValues([row]);
}

function genId(prefix) {
  return prefix + "_" + new Date().getTime() + "_" + Math.floor(Math.random() * 1000);
}

function formatThaiDate(date) {
  const months = ["ม.ค.", "ก.พ.", "มี.ค.", "เม.ย.", "พ.ค.", "มิ.ย.", "ก.ค.", "ส.ค.", "ก.ย.", "ต.ค.", "พ.ย.", "ธ.ค."];
  const d = new Date(date);
  const buddhistYear = d.getFullYear() + 543;
  return `${d.getDate()} ${months[d.getMonth()]} ${String(buddhistYear).slice(-2)}`;
}

// ===================================================================
// สิทธิ์ผู้ใช้ — อ่านบทบาท (admin/user) + ห้อง Lab ของผู้เรียก API จาก line_user_id
// ที่ liff-app/js/api.js แนบมาให้อัตโนมัติทุก request
// ===================================================================

/** หาแถวผู้ใช้จาก line_user_id เพียวๆ ไม่ throw ถ้าไม่พบ (คืน null) */
function findLineUser(line_user_id) {
  if (!line_user_id) return null;
  const users = getSheetData("LineUsers");
  return users.find(u => u.line_user_id === line_user_id) || null;
}

/**
 * แอดมินอนุมัติแล้วหรือยัง — ถือว่าอนุมัติอัตโนมัติถ้าเป็น role admin หรือมี lab ตั้งไว้แล้ว
 * (กันคนที่ใช้งานอยู่ก่อนฟีเจอร์อนุมัติจะมาถูกบล็อกย้อนหลัง เพราะแค่มี lab ก็แปลว่าเคยผ่านขั้นนี้มาแล้วจริง)
 */
function isUserApproved(row) {
  return row.role === "admin" || row.approved === true || !!row.lab;
}

/**
 * ดึงข้อมูลผู้เรียก API จาก payload.line_user_id
 * isAdmin = true ถ้า role ในชีตเป็น "admin" หรือถ้าเป็นคนที่ตรงกับ ADMIN_LINE_USER_ID
 * (ใช้ปลดล็อกแอดมินคนแรกตอนยังไม่มีใคร role admin ในชีตเลย — bootstrap admin ถือว่า approved ด้วยเสมอ)
 */
function getCurrentUser(payload) {
  const line_user_id = payload.line_user_id || "";
  const row = findLineUser(line_user_id);
  const bootstrapAdminId = PropertiesService.getScriptProperties().getProperty("ADMIN_LINE_USER_ID");
  const isBootstrapAdmin = !!bootstrapAdminId && line_user_id === bootstrapAdminId;
  const isAdmin = (row && row.role === "admin") || isBootstrapAdmin;
  const approved = isBootstrapAdmin || (row ? isUserApproved(row) : false);
  return {
    line_user_id,
    display_name: row ? row.display_name : "",
    lab: row ? row.lab : "",
    isAdmin,
    approved,
  };
}

/** ใช้ใน handler ที่ต้องเป็นแอดมินเท่านั้น (เช่น จัดการสิทธิ์ผู้ใช้) */
function requireAdmin(payload) {
  const me = getCurrentUser(payload);
  if (!me.isAdmin) throw new Error("ต้องเป็นแอดมินเท่านั้นถึงทำรายการนี้ได้");
  return me;
}

/**
 * หา Lab ที่จะใช้กรองข้อมูลจริงสำหรับคำขอนี้
 * - User ทั่วไป: ผูกกับ Lab ตัวเองเสมอ ไม่สนใจ payload.view_lab ที่ส่งมา (ป้องกันแก้ payload เพื่อแอบดู Lab อื่น)
 * - Admin: เลือกได้ — ถ้าไม่ส่ง view_lab มา (หรือส่งมาไม่ตรงกับ LABS) = null หมายถึง "ดูรวมทุก Lab"
 *   ถ้าส่ง view_lab มาเป็น KA/NJ/CC = ดูแค่ Lab นั้น
 */
function resolveViewLab(me, payload) {
  if (!me.isAdmin) return me.lab;
  return (payload.view_lab && LABS.indexOf(payload.view_lab) !== -1) ? payload.view_lab : null;
}

// ===================================================================
// API ROUTER — เรียกจาก liff-app/js/api.js ด้วย { action, payload }
// ===================================================================
function doPost(e) {
  let response;
  try {
    const body = JSON.parse(e.postData.contents);
    const action = body.action;
    const payload = body.payload || {};
    const handlers = {
      registerLineUser,
      setMyLab,
      listUsers,
      setUserPermissions,
      listPlantsWithCount,
      listPlantsFull,
      savePlantType,
      deactivatePlantType,
      listStorage,
      addStorageValue,
      deactivateStorageValue,
      recordTransfer,
      getDashboard,
      listActiveLots,
      getLotHistory,
      simulateProduction,
      getSystemSettings,
      saveSystemSettings,
    };
    if (!handlers[action]) throw new Error("ไม่รู้จัก action: " + action);
    const result = handlers[action](payload);
    response = { ok: true, result };
  } catch (err) {
    response = { ok: false, message: err.message };
  }
  return ContentService.createTextOutput(JSON.stringify(response))
    .setMimeType(ContentService.MimeType.JSON);
}

// ===================================================================
// Handlers: ผู้ใช้ Line
// ===================================================================
function registerLineUser(payload) {
  const users = getSheetData("LineUsers");
  const existing = users.find(u => u.line_user_id === payload.line_user_id);
  const me = getCurrentUser(payload);
  // role ที่คืนให้ฝั่งหน้าจอต้องอิง me.isAdmin (เช็ค ADMIN_LINE_USER_ID บูตสแตรปด้วย) ไม่ใช่แค่คอลัมน์ role ในชีตตรงๆ
  // ไม่งั้นแอดมินคนแรกที่ปลดล็อกผ่าน bootstrap จะเห็นแอปเป็น user ธรรมดา (แท็บสิทธิ์ผู้ใช้/ป้าย Lab ผิด) ทั้งที่ getDashboard ฝั่ง backend เห็นว่าเป็นแอดมินแล้ว
  const role = me.isAdmin ? "admin" : (existing ? (existing.role || "user") : "user");
  if (existing) {
    if (!existing.active || existing.display_name !== payload.display_name) {
      updateRow("LineUsers", existing._row, { ...existing, display_name: payload.display_name, active: true });
    }
    return { updated: true, role, lab: existing.lab || "", approved: me.approved };
  }
  appendRow("LineUsers", { line_user_id: payload.line_user_id, display_name: payload.display_name, role: "user", lab: "", active: true, approved: false });
  return { created: true, role, lab: "", approved: me.approved };
}

/** ผู้ใช้เลือก Lab ของตัวเองครั้งแรกที่เปิดแอป (ตั้งได้ครั้งเดียว — ถ้าจะเปลี่ยนทีหลังต้องให้แอดมินตั้งให้ผ่านหน้า "สิทธิ์ผู้ใช้") */
function setMyLab(payload) {
  if (LABS.indexOf(payload.lab) === -1) throw new Error("กรุณาเลือก Lab ให้ถูกต้อง");
  const me = getCurrentUser(payload);
  if (!me.approved) throw new Error("ยังไม่ได้รับการอนุมัติจากแอดมิน กรุณารอแอดมินอนุมัติก่อนเลือก Lab");
  const existing = findLineUser(payload.line_user_id);
  if (!existing) throw new Error("ไม่พบผู้ใช้นี้ กรุณาเปิดแอปผ่าน Line ใหม่อีกครั้ง");
  if (existing.lab) return { lab: existing.lab }; // ตั้งไปแล้ว ไม่ให้ตั้งซ้ำ (ให้แอดมินเปลี่ยนแทน)
  updateRow("LineUsers", existing._row, { ...existing, lab: payload.lab });
  return { lab: payload.lab };
}

/** รายชื่อผู้ใช้ทั้งหมด + บทบาท/Lab — สำหรับหน้า "สิทธิ์ผู้ใช้" (แอดมินเท่านั้น) */
function listUsers(payload) {
  requireAdmin(payload);
  return getSheetData("LineUsers").map(u => ({
    line_user_id: u.line_user_id,
    display_name: u.display_name,
    role: u.role || "user",
    lab: u.lab || "",
    active: u.active,
    approved: isUserApproved(u),
  }));
}

/**
 * แอดมินเปลี่ยนบทบาท/Lab ของผู้ใช้คนใดก็ได้ และ/หรืออนุมัติให้เข้าใช้งาน
 * การตั้ง Lab ให้ใคร ถือเป็นการอนุมัติไปในตัวเสมอ (จะอนุมัติแยกโดยไม่ตั้ง Lab ก็ได้ผ่าน payload.approved)
 */
function setUserPermissions(payload) {
  requireAdmin(payload);
  if (!payload.target_line_user_id) throw new Error("ไม่พบผู้ใช้ที่จะแก้ไข");
  if (payload.role && ["admin", "user"].indexOf(payload.role) === -1) throw new Error("บทบาทต้องเป็น admin หรือ user");
  if (payload.lab && LABS.indexOf(payload.lab) === -1) throw new Error("Lab ไม่ถูกต้อง");
  const existing = findLineUser(payload.target_line_user_id);
  if (!existing) throw new Error("ไม่พบผู้ใช้นี้");
  const approveNow = payload.approved === true || payload.lab !== undefined;
  updateRow("LineUsers", existing._row, {
    ...existing,
    role: payload.role !== undefined ? payload.role : existing.role,
    lab: payload.lab !== undefined ? payload.lab : existing.lab,
    approved: approveNow ? true : existing.approved,
  });
  return { updated: true };
}

// ===================================================================
// Handlers: ชนิดพืช
// ===================================================================
function listPlantsFull() {
  return getSheetData("PlantTypes");
}

function listPlantsWithCount(payload) {
  const me = getCurrentUser(payload);
  const viewLab = resolveViewLab(me, payload);
  const plants = getSheetData("PlantTypes").filter(p => p.active !== false);
  const transfers = getSheetData("Transfers").filter(t => t.lab && (viewLab ? t.lab === viewLab : me.isAdmin));
  // เลขล็อตล่าสุดที่เคยใช้ ต้องอิง Lab ของผู้เรียกเองเสมอ (me.lab) ไม่ใช่ viewLab — เพราะใช้ช่วยจำว่า
  // "ล็อตใหม่ที่จะบันทึกเข้า Lab ตัวเอง" ควรเป็นเลขอะไรต่อ ถ้าใช้ viewLab แอดมินที่ดูรวมทุก Lab จะได้เลขที่ปนกันข้าม Lab ไม่มีประโยชน์
  const myLabTransfers = me.lab ? getSheetData("Transfers").filter(t => t.lab === me.lab) : [];
  return plants.map(p => {
    const latestByLot = {};
    transfers.filter(t => t.plant_id === p.plant_id).forEach(t => {
      const key = t.lab + "|" + t.lot_no;
      if (!latestByLot[key] || parseThaiDateApprox(t.transfer_date) > parseThaiDateApprox(latestByLot[key].transfer_date)) {
        latestByLot[key] = t;
      }
    });
    const activeCount = Object.values(latestByLot).filter(t => !!t.next_transfer_date).length;
    // เทียบเลขล็อตเป็นตัวเลขเท่านั้น (ข้ามค่าที่ไม่ใช่ตัวเลข เช่นกรอกตัวอักษรปนมา) เอาเลขสูงสุดที่เคยใช้ในชนิดพืชนี้ + Lab ตัวเอง
    const numericLotNos = myLabTransfers
      .filter(t => t.plant_id === p.plant_id)
      .map(t => Number(t.lot_no))
      .filter(n => !isNaN(n));
    const latest_lot_no = numericLotNos.length > 0 ? Math.max(...numericLotNos) : null;
    return { plant_id: p.plant_id, plant_name: p.plant_name, active_count: activeCount, latest_lot_no };
  });
}

function savePlantType(payload) {
  const required = [
    "plant_name",
    "stage1_rounds", "stage1_days", "stage1_recipe_tib", "stage1_recipe_ss",
    "stage2_days", "stage2_recipe_tib", "stage2_recipe_ss",
    "stage3_recipe_tib", "stage3_recipe_ss",
  ];
  required.forEach(f => {
    if (payload[f] === undefined || payload[f] === "") throw new Error("กรุณากรอกข้อมูลให้ครบ: " + f);
  });
  // ตัวคูณการขยายไม่บังคับ แต่ถ้ากรอกมาต้องเป็นเลข >= 0 (ปล่อยว่างได้ = ไม่มีตัวคูณ)
  ["stage1_multiplier_tib", "stage1_multiplier_ss"].forEach(f => {
    if (payload[f] !== undefined && payload[f] !== "" && (isNaN(Number(payload[f])) || Number(payload[f]) < 0)) {
      throw new Error("ตัวคูณการขยายต้องเป็นตัวเลขมากกว่าหรือเท่ากับ 0: " + f);
    }
  });
  const plants = getSheetData("PlantTypes");
  if (payload.plant_id) {
    const existing = plants.find(p => p.plant_id === payload.plant_id);
    if (!existing) throw new Error("ไม่พบชนิดพืชนี้");
    const active = payload.active !== undefined ? payload.active : true;
    updateRow("PlantTypes", existing._row, { ...payload, active });
    return { updated: true, plant_id: payload.plant_id };
  }
  const plant_id = genId("plant");
  appendRow("PlantTypes", { ...payload, plant_id, active: true });
  return { created: true, plant_id };
}

function deactivatePlantType(payload) {
  if (!payload.plant_id) throw new Error("ไม่พบชนิดพืชนี้");
  const plants = getSheetData("PlantTypes");
  const existing = plants.find(p => p.plant_id === payload.plant_id);
  if (!existing) throw new Error("ไม่พบชนิดพืชนี้");
  updateRow("PlantTypes", existing._row, { ...existing, active: false });
  return { deactivated: true };
}

// ===================================================================
// Handlers: ที่เก็บ
// ===================================================================
function listStorage() {
  const rows = getSheetData("Storage").filter(s => s.active !== false);
  return {
    shelves: rows.filter(s => s.type === "shelf").map(s => s.value),
    subs: rows.filter(s => s.type === "sub_shelf").map(s => s.value),
  };
}

function addStorageValue(payload) {
  if (!payload.type || !payload.value) throw new Error("ข้อมูลที่เก็บไม่ครบ");
  const rows = getSheetData("Storage");
  const dup = rows.find(s => s.type === payload.type && String(s.value) === String(payload.value));
  if (dup) {
    if (dup.active === false) {
      updateRow("Storage", dup._row, { type: dup.type, value: dup.value, active: true });
      return { reactivated: true };
    }
    return { duplicated: true };
  }
  appendRow("Storage", { type: payload.type, value: payload.value, active: true });
  return { created: true };
}

function deactivateStorageValue(payload) {
  if (!payload.type || !payload.value) throw new Error("ข้อมูลที่เก็บไม่ครบ");
  const rows = getSheetData("Storage");
  const row = rows.find(s => s.type === payload.type && String(s.value) === String(payload.value));
  if (!row) throw new Error("ไม่พบรายการนี้");
  updateRow("Storage", row._row, { type: row.type, value: row.value, active: false });
  return { deactivated: true };
}

// ===================================================================
// Handler หลัก: บันทึกการถ่ายโอนเนื้อเยื่อ + คำนวณระยะ/สูตร/วันถัดไปอัตโนมัติ
// ===================================================================
function recordTransfer(payload) {
  const required = ["plant_id", "lot_no", "bottle_count", "plants_per_bottle", "shelf", "sub_shelf", "media_type"];
  required.forEach(f => {
    if (payload[f] === undefined || payload[f] === "") throw new Error("กรุณากรอกข้อมูลให้ครบ: " + f);
  });
  if (Number(payload.bottle_count) <= 0 || Number(payload.plants_per_bottle) <= 0) {
    throw new Error("จำนวนขวดและจำนวนต้นต่อขวดต้องมากกว่า 0");
  }
  if (MEDIA_TYPES.indexOf(payload.media_type) === -1) {
    throw new Error("ประเภทรอบต้องเป็น TIB หรือ SS");
  }
  if (payload.media_type === "TIB" && !payload.bottle_no) {
    throw new Error("กรุณาใส่หมายเลขขวด (จำเป็นสำหรับ TIB)");
  }

  const me = getCurrentUser(payload);
  if (!me.approved) throw new Error("ยังไม่ได้รับการอนุมัติจากแอดมิน กรุณารอแอดมินอนุมัติก่อนใช้งาน");
  if (!me.lab) throw new Error("ยังไม่ได้ระบุ Lab ของคุณ กรุณาเปิดแอปใหม่อีกครั้ง");

  const plants = getSheetData("PlantTypes");
  const plant = plants.find(p => p.plant_id === payload.plant_id);
  if (!plant) throw new Error("ไม่พบชนิดพืชนี้");

  // ระบุรอบที่ด้วย (plant_id, lot_no, lab) ไม่ใช้ตำแหน่งเก็บ (shelf/sub_shelf) อีกต่อไป
  // เพราะล็อตเดียวกันอาจย้ายตำแหน่งเก็บได้ระหว่างรอบ — ตำแหน่งเป็นแค่ metadata ของรอบนั้นๆ
  // ผูกกับ lab ด้วยเพื่อไม่ให้ Lab อื่นที่ใช้หมายเลขล็อตเดียวกันโดยบังเอิญมาปนกัน
  // ถ้า lot_no นี้เคยจบไปแล้ว (แถวล่าสุดไม่มี next_transfer_date) จะถือเป็นล็อตใหม่ เริ่มรอบที่ 1 ใหม่
  const transfers = getSheetData("Transfers");
  const sameLot = transfers.filter(t =>
    t.plant_id === payload.plant_id && String(t.lot_no) === String(payload.lot_no) && t.lab === me.lab
  );
  let previous = null;
  sameLot.forEach(t => {
    if (t.next_transfer_date && (!previous || parseThaiDateApprox(t.transfer_date) > parseThaiDateApprox(previous.transfer_date))) {
      previous = t;
    }
  });

  // ผู้ใช้เลือกโหมด "ล็อตใหม่" ตั้งใจสร้างล็อตที่ไม่เคยมีมาก่อน — ถ้าเลขล็อตนี้ยังมีรอบที่แอคทีฟอยู่
  // (ยังไม่ถึงระยะพร้อมออกปลูก) ถือว่าซ้ำกับล็อตเดิม ต้องกันไว้ ไม่ให้บันทึกทับเป็นรอบต่อไปของล็อตเดิมโดยไม่ตั้งใจ
  // (เลขล็อตที่ "จบไปแล้ว" ยังนำมาใช้ใหม่ได้ตามปกติ ไม่ถือว่าซ้ำ)
  if (payload.lot_mode === "new" && previous) {
    throw new Error("เลขล็อต " + payload.lot_no + " กำลังถูกติดตามอยู่แล้ว (ยังไม่ถึงระยะพร้อมออกปลูก) กรุณาใช้เลขล็อตอื่น หรือเลือก \"ล็อตที่มีอยู่\" แทน");
  }

  const round_no = previous ? Number(previous.round_no) + 1 : 1;
  const stage1Rounds = Number(plant.stage1_rounds) || 1;

  const today = new Date();
  let stage, recipe_used, next_transfer_date = "", round_in_stage;
  const isTib = payload.media_type === "TIB";

  if (round_no <= stage1Rounds) {
    stage = STAGE_NAMES[0];
    recipe_used = isTib ? plant.stage1_recipe_tib : plant.stage1_recipe_ss;
    const next = new Date(today); next.setDate(next.getDate() + Number(plant.stage1_days));
    next_transfer_date = formatThaiDate(next);
    round_in_stage = `${round_no}/${stage1Rounds}`;
  } else if (round_no === stage1Rounds + 1) {
    stage = STAGE_NAMES[1];
    recipe_used = isTib ? plant.stage2_recipe_tib : plant.stage2_recipe_ss;
    const next = new Date(today); next.setDate(next.getDate() + Number(plant.stage2_days));
    next_transfer_date = formatThaiDate(next);
    round_in_stage = "1/1";
  } else {
    stage = STAGE_NAMES[2];
    recipe_used = isTib ? plant.stage3_recipe_tib : plant.stage3_recipe_ss;
    next_transfer_date = "";
    round_in_stage = "-";
  }

  const bottle_count = Number(payload.bottle_count);
  const plants_per_bottle = Number(payload.plants_per_bottle);
  const quantity = bottle_count * plants_per_bottle;

  const record = {
    record_id: genId("rec"),
    plant_id: plant.plant_id,
    plant_name: plant.plant_name,
    lot_no: payload.lot_no,
    lab: me.lab,
    round_no,
    transfer_date: formatThaiDate(today),
    quantity,
    bottle_count,
    plants_per_bottle,
    media_type: payload.media_type,
    bottle_no: isTib ? payload.bottle_no : "",
    shelf: payload.shelf,
    sub_shelf: payload.sub_shelf,
    stage,
    recipe_used,
    next_transfer_date,
    recorded_by: payload.recorded_by || "ไม่ระบุชื่อ",
    created_at: today.toISOString(),
  };
  appendRow("Transfers", record);

  notifyAfterRecord(record, round_in_stage, me.line_user_id);

  return {
    lot_no: record.lot_no,
    stage,
    recipe_used,
    next_transfer_date,
    round_in_stage,
    media_type: record.media_type,
    bottle_no: record.bottle_no,
  };
}

/**
 * แจ้งเตือนทันทีหลังบันทึกสำเร็จ: ส่งให้ (1) ผู้บันทึกเอง (ยืนยันว่าบันทึกของตัวเองสำเร็จ)
 * และ (2) แอดมินทุกคน (เห็นทุก Lab) — ไม่ส่งให้คนอื่นใน Lab เดียวกันที่ไม่ได้เป็นคนบันทึก
 */
function notifyAfterRecord(record, round_in_stage, recorderLineUserId) {
  try {
    const users = getSheetData("LineUsers").filter(u => u.active !== false);
    const adminIds = users.filter(u => u.role === "admin").map(u => u.line_user_id);
    const bootstrapAdminId = PropertiesService.getScriptProperties().getProperty("ADMIN_LINE_USER_ID");
    const recipients = new Set(adminIds);
    if (bootstrapAdminId) recipients.add(bootstrapAdminId);
    if (recorderLineUserId) recipients.add(recorderLineUserId);
    if (recipients.size === 0) return;

    const lines = [
      "📝 มีการบันทึกถ่ายโอนเนื้อเยื่อใหม่",
      `Lab ${record.lab} — ${record.plant_name} — ล็อต ${record.lot_no} — ${record.stage} (รอบที่ ${round_in_stage})`,
      `ประเภท: ${record.media_type}${record.bottle_no ? " · ขวดเลขที่ " + record.bottle_no : ""}`,
      `ที่เก็บ: ชั้น ${record.shelf}/${record.sub_shelf} · จำนวน ${record.quantity} ต้น`,
      `บันทึกโดย: ${record.recorded_by}`,
    ];
    pushLineMessage([...recipients], lines.join("\n"));
  } catch (err) {
    Logger.log("แจ้งเตือนหลังบันทึกไม่สำเร็จ: " + err.message);
  }
}

// ===================================================================
// Handler: Dashboard
// ===================================================================

/**
 * คำนวณข้อมูล Dashboard จากรายการ transfers ที่กรองมาแล้ว (กรองตาม Lab ไว้ก่อนเรียกฟังก์ชันนี้)
 * แยกออกมาเป็นฟังก์ชันกลาง เพื่อให้ getDashboard() (ต่อ Lab ของผู้ใช้) และ sendDueReminders()
 * (ต้องคำนวณทีละ Lab + รวมทุก Lab สำหรับแอดมิน) ใช้ตรรกะเดียวกันไม่ซ้ำโค้ด
 */
function computeDashboardData(transfers, plants, storage, targetDate) {
  // นับ "ล็อต" ที่กำลังติดตามอยู่ ไม่ใช่ตำแหน่งเก็บ — แถวเก่าที่ไม่มี lot_no/lab (ก่อนอัปเดตฟีเจอร์นี้) ถูกข้าม
  const latestByLot = {};
  transfers.forEach(t => {
    if (!t.lot_no) return;
    const key = t.plant_id + "|" + t.lab + "|" + t.lot_no;
    if (!latestByLot[key] || parseThaiDateApprox(t.transfer_date) > parseThaiDateApprox(latestByLot[key].transfer_date)) {
      latestByLot[key] = t;
    }
  });
  const activeLots = Object.values(latestByLot).filter(t => !!t.next_transfer_date);

  const total_lots = activeLots.length;
  const total_plants = activeLots.reduce((sum, t) => sum + Number(t.quantity || 0), 0);

  const now = new Date();
  const withDays = activeLots
    .map(t => ({ ...t, _daysLeft: (parseThaiDateApprox(t.next_transfer_date) - now) / 86400000 }))
    .filter(t => t._daysLeft <= 3)
    .sort((a, b) => a._daysLeft - b._daysLeft)
    .slice(0, 20);

  const due_soon = withDays.map(t => {
    const plant = plants.find(p => p.plant_id === t.plant_id);
    const stage1Rounds = plant ? (Number(plant.stage1_rounds) || 1) : 1;
    let round_in_stage = "-";
    if (t.stage === STAGE_NAMES[0]) round_in_stage = `${t.round_no}/${stage1Rounds}`;
    else if (t.stage === STAGE_NAMES[1]) round_in_stage = "1/1";
    return {
      plant_name: t.plant_name,
      lot_no: t.lot_no,
      lab: t.lab,
      shelf: t.shelf,
      sub_shelf: t.sub_shelf,
      media_type: t.media_type,
      bottle_no: t.bottle_no,
      stage: t.stage,
      round_in_stage,
      next_transfer_date: t.next_transfer_date,
      next_recipe: nextRecipeFor(plants, t),
      recorded_by: t.recorded_by,
    };
  });

  // คำนวณพยากรณ์อัตโนมัติต่อชนิดพืชใน Dashboard — เริ่มจาก TODAY() เสมอ (ไม่ใช่ next_transfer_date ของล็อต)
  // นับเฉพาะ "ต้น" ที่ยังอยู่ระยะขยายเป็นชิ้นตั้งต้น ไม่ใช้ตัวคูณการขยายต่อชนิดพืชอีกต่อไป
  // ใช้แบบจำลองเดียวกับเครื่องมือ "จำลองรอบการผลิต" ทุกประการ (เรียก runProductionCycles ตรงๆ ไม่คำนวณซ้ำ) — ติดเพดานความจุ
  // ของระบบ SS/TIB เหมือนกัน ส่วนเกินความจุถูกนำออกไปกระตุ้นรากแบบเดียวกัน ตั้ง acclimatizationRate = 100% เสมอ
  // (Dashboard ไม่มีช่องกรอกอัตรารอด เป็นแค่ตัวเลขประเมินคร่าวๆ) — รวมจำนวนที่ยังอยู่ในวงจรขยาย + ที่นำออกไปกระตุ้นรากแล้วทุกรอบ
  // เป็นยอดเดียวกัน เพื่อให้ "พยากรณ์" หมายถึงจำนวนต้นทั้งหมดที่จะมีอยู่ ณ วันนั้น ไม่ใช่แค่ส่วนที่ยังอยู่ในระบบ
  const sysConfig = targetDate ? getSystemConfig() : null;
  const by_plant = plants.filter(p => p.active !== false).map(p => {
    const lots = activeLots.filter(t => t.plant_id === p.plant_id);
    // stage_counts นับเป็นจำนวน "ต้น" รวม (ผลรวม quantity ของล็อตในระยะนั้น) ไม่ใช่จำนวนล็อต
    const stage_counts = { "ขยาย": 0, "กระตุ้นราก": 0, "พร้อมออกปลูก": 0 };
    lots.forEach(t => { stage_counts[t.stage] = (stage_counts[t.stage] || 0) + Number(t.quantity || 0); });
    const dominant_stage = Object.keys(stage_counts).reduce((a, b) => stage_counts[a] >= stage_counts[b] ? a : b, "ขยาย");
    const result = { plant_name: p.plant_name, total: lots.length, stage_counts, dominant_stage };
    if (targetDate && sysConfig) {
      const initialPieces = stage_counts["ขยาย"];
      if (initialPieces > 0) {
        const sim = runProductionCycles({
          initialPieces,
          startDate: now,
          targetDate,
          targetPlants: null,
          cycleDays: sysConfig.cycle_days,
          ssBottleLimit: sysConfig.ss_bottle_limit,
          ssPiecesPerBottle: sysConfig.ss_pieces_per_bottle,
          ssMultiplier: sysConfig.ss_multiplication_factor,
          tibBottleLimit: sysConfig.tib_bottle_limit,
          tibPiecesPerBottle: sysConfig.tib_pieces_per_bottle,
          tibMultiplier: sysConfig.tib_multiplication_factor,
          acclimatizationRate: 1,
        });
        const lastRow = sim.rows[sim.rows.length - 1];
        // total_pieces_after_cycle ของรอบสุดท้าย = ส่วนที่ยังอยู่ในวงจรขยาย + ส่วนเกินที่กำลังจะถูกนำออกไปกระตุ้นราก (ยังไม่นับใน cumulative)
        // รวมกับยอดสะสม "พร้อมออกปลูก" ของทุกรอบก่อนหน้า ได้ยอดรวมทั้งหมดที่เคย/กำลังจะผลิตได้ ณ วันนั้น
        result.predicted_total = lastRow ? (lastRow.total_pieces_after_cycle + sim.summary.total_ready_for_planting) : initialPieces;
      } else {
        result.predicted_total = 0;
      }
    }
    return result;
  });

  const shelves = storage.filter(s => s.type === "shelf").map(s => s.value);
  const subs = storage.filter(s => s.type === "sub_shelf").map(s => s.value);
  const storage_overview = [];
  shelves.forEach(shelf => {
    subs.forEach(sub => {
      const count = activeLots.filter(t => String(t.shelf) === String(shelf) && String(t.sub_shelf) === String(sub)).length;
      storage_overview.push({ shelf, sub_shelf: sub, count });
    });
  });

  return { total_lots, total_plants, due_soon, by_plant, storage_overview };
}

/**
 * Dashboard ของผู้เรียก — User เห็นแค่ Lab ตัวเอง
 * Admin เลือกได้ผ่าน payload.view_lab: ไม่ส่งมา = ดูรวมทุก Lab, ส่ง KA/NJ/CC = ดูแค่ Lab นั้น
 */
function getDashboard(payload) {
  const me = getCurrentUser(payload);
  const viewLab = resolveViewLab(me, payload);
  const transfers = getSheetData("Transfers").filter(t => t.lab && (viewLab ? t.lab === viewLab : me.isAdmin));
  const plants = getSheetData("PlantTypes");
  const storage = getSheetData("Storage").filter(s => s.active !== false);
  const targetDate = payload.target_date ? new Date(payload.target_date) : null;
  const data = computeDashboardData(transfers, plants, storage, targetDate);
  const members = getSheetData("LineUsers")
    .filter(u => u.active !== false && (viewLab ? u.lab === viewLab : true))
    .map(u => ({ display_name: u.display_name, role: u.role || "user", lab: u.lab || "" }));
  return { ...data, lab: viewLab, is_admin: me.isAdmin, members };
}

// ===================================================================
// Handlers: ประวัติย้อนหลังของล็อต (track ว่าทำ subculture มากี่รอบแล้ว)
// ===================================================================

/**
 * รายการล็อตที่ "กำลังติดตามอยู่" ของชนิดพืชหนึ่ง (ยังไม่ถึงระยะพร้อมออกปลูก)
 * User เห็นเฉพาะ Lab ตัวเอง — Admin เลือกได้ผ่าน payload.view_lab (ไม่ส่งมา = เห็นรวมทุก Lab)
 */
function listActiveLots(payload) {
  if (!payload.plant_id) throw new Error("กรุณาระบุชนิดพืช");
  const me = getCurrentUser(payload);
  const viewLab = resolveViewLab(me, payload);
  const transfers = getSheetData("Transfers").filter(t =>
    t.plant_id === payload.plant_id && t.lab && (viewLab ? t.lab === viewLab : me.isAdmin)
  );
  const latestByLot = {};
  transfers.forEach(t => {
    const key = t.lab + "|" + t.lot_no;
    if (!latestByLot[key] || parseThaiDateApprox(t.transfer_date) > parseThaiDateApprox(latestByLot[key].transfer_date)) {
      latestByLot[key] = t;
    }
  });
  return Object.values(latestByLot)
    .filter(t => !!t.next_transfer_date)
    .map(t => ({
      lot_no: t.lot_no,
      lab: t.lab,
      shelf: t.shelf,
      sub_shelf: t.sub_shelf,
      round_no: t.round_no,
      stage: t.stage,
      media_type: t.media_type,
      bottle_no: t.bottle_no,
      next_transfer_date: t.next_transfer_date,
    }))
    .sort((a, b) => String(a.lot_no) > String(b.lot_no) ? 1 : -1);
}

/**
 * ประวัติทุกรอบของล็อตเดียว (ไล่ย้อนจากรอบปัจจุบันกลับไปรอบที่ 1)
 * อิงรอบที่ตาม (plant_id, lot_no, lab) ไม่ใช่ตำแหน่งเก็บ — เพราะล็อตอาจย้ายชั้น/ช่องระหว่างรอบได้
 * ต้องระบุ lab มาให้ชัดเจนเสมอ (ไม่ใช่แค่ plant_id+lot_no) เพราะ Admin มองเห็นได้หลาย Lab พร้อมกัน —
 * ถ้าสอง Lab บังเอิญใช้หมายเลขล็อตเดียวกันสำหรับพืชชนิดเดียวกัน จะไม่ปนกัน
 */
function getLotHistory(payload) {
  if (!payload.plant_id || !payload.lot_no || !payload.lab) throw new Error("ข้อมูลไม่ครบ");
  const me = getCurrentUser(payload);
  if (!me.isAdmin && payload.lab !== me.lab) throw new Error("ไม่มีสิทธิ์ดูข้อมูลของ Lab นี้");
  const transfers = getSheetData("Transfers").filter(t =>
    t.plant_id === payload.plant_id && String(t.lot_no) === String(payload.lot_no) && t.lab === payload.lab
  );
  let current = null;
  transfers.forEach(t => {
    if (!current || parseThaiDateApprox(t.transfer_date) > parseThaiDateApprox(current.transfer_date)) current = t;
  });
  if (!current) return [];

  const chain = [current];
  let cursor = current;
  while (Number(cursor.round_no) > 1) {
    const targetRound = Number(cursor.round_no) - 1;
    const prevCandidates = transfers.filter(t =>
      Number(t.round_no) === targetRound && parseThaiDateApprox(t.transfer_date) <= parseThaiDateApprox(cursor.transfer_date)
    );
    if (prevCandidates.length === 0) break;
    prevCandidates.sort((a, b) => parseThaiDateApprox(b.transfer_date) - parseThaiDateApprox(a.transfer_date));
    cursor = prevCandidates[0];
    chain.push(cursor);
  }
  chain.reverse();
  return chain.map(t => ({
    round_no: t.round_no,
    transfer_date: t.transfer_date,
    stage: t.stage,
    media_type: t.media_type,
    bottle_no: t.bottle_no,
    shelf: t.shelf,
    sub_shelf: t.sub_shelf,
    quantity: t.quantity,
    recipe_used: t.recipe_used,
    next_transfer_date: t.next_transfer_date,
    recorded_by: t.recorded_by,
  }));
}

// ===================================================================
// Handler: ตั้งค่าระบบ SS/TIB (พารามิเตอร์ระดับระบบ ไม่ใช่ต่อชนิดพืช)
// ===================================================================

/** อ่านค่าพารามิเตอร์ระบบ SS/TIB ปัจจุบัน (ผสมกับค่า default ถ้าบางช่องยังไม่ตั้ง) — ใช้ภายใน ไม่ผ่าน payload */
function getSystemConfig() {
  const rows = getSheetData("SystemConfig");
  const row = rows.find(r => r.config_id === "global");
  const merged = {};
  Object.keys(SYSTEM_CONFIG_DEFAULTS).forEach(key => {
    const raw = row ? row[key] : undefined;
    merged[key] = (raw !== undefined && raw !== "") ? Number(raw) : SYSTEM_CONFIG_DEFAULTS[key];
  });
  return merged;
}

/** สำหรับหน้าตั้งค่า > ระบบ SS/TIB และเป็นค่าเริ่มต้นของหน้า "จำลองรอบการผลิต" */
function getSystemSettings() {
  return getSystemConfig();
}

/** แอดมินเท่านั้นที่แก้พารามิเตอร์ระบบ SS/TIB ได้ (มีผลกับการพยากรณ์ทุกชนิดพืชใน Dashboard ทันที) */
function saveSystemSettings(payload) {
  requireAdmin(payload);
  const fields = Object.keys(SYSTEM_CONFIG_DEFAULTS);
  fields.forEach(f => {
    if (payload[f] === undefined || payload[f] === "" || isNaN(Number(payload[f])) || Number(payload[f]) <= 0) {
      throw new Error("กรุณากรอกค่าให้ถูกต้อง (ต้องเป็นตัวเลขมากกว่า 0): " + f);
    }
  });
  const rows = getSheetData("SystemConfig");
  const existing = rows.find(r => r.config_id === "global");
  const record = { config_id: "global" };
  fields.forEach(f => { record[f] = Number(payload[f]); });
  if (existing) {
    updateRow("SystemConfig", existing._row, record);
  } else {
    appendRow("SystemConfig", record);
  }
  return { updated: true };
}

// ===================================================================
// Handler: จำลองรอบการผลิต (เครื่องมือวางแผนกำลังผลิตล่วงหน้า — ไม่อ่าน/เขียน Transfers แยกอิสระจากข้อมูลล็อตจริง)
// ===================================================================

/**
 * จำลองการผลิตเนื้อเยื่อพืชแบบรอบการผลิต (ทุก cycleDays วัน) จนถึง targetDate หรือจนกว่ายอดสะสม "พร้อมออกปลูก" จะถึง targetPlants
 * แต่ละรอบ: ดันจำนวนชิ้นตั้งต้นเข้าระบบ TIB ให้เต็มความจุก่อนเสมอ (ตัวคูณสูงกว่า ใช้ประโยชน์จากความจุที่มีให้คุ้มที่สุด) เหลือเท่าไหร่
 * ค่อยลงระบบ SS ส่วนที่เกินความจุรวมทั้งสองระบบ (ทั้ง TIB และ SS เต็มแล้ว ไม่มีที่ลงต่อ) ถือเป็น
 * "ส่วนเกิน" ที่ถูกนำออกไปกระตุ้นรากทันที — ไม่มีอัตรา/เกณฑ์ขั้นต่ำให้ตั้งเองอีกต่อไป ความจุของระบบ SS/TIB เป็นตัวกำหนดเองโดยอัตโนมัติ
 * ส่วนที่โหลดเข้า SS/TIB ได้ คูณด้วยตัวคูณของระบบนั้น แล้ววนกลับเป็นชิ้นตั้งต้นของรอบถัดไป (เฉพาะผลผลิตที่คูณแล้วเท่านั้นที่วนต่อ) จากนั้น:
 *   1. ส่วนเกินของรอบนี้ (ที่ไม่มีที่ลง SS/TIB) การกระตุ้นรากใช้เวลา 1 รอบการผลิต — ของที่ถูกนำออกไปในรอบ N จะยังไม่กลายเป็น "พร้อมออกปลูก"
 *      ในรอบ N ทันที แต่ไปโผล่ในรอบ N+1 โดยคูณด้วย acclimatizationRate ตอนนั้น (ส่วนที่ไม่รอดถือว่าสูญไปเลย ไม่กลับเข้าวงจรขยาย)
 *   2. ผลผลิตจาก SS/TIB ที่คูณแล้ว (ไม่ใช่ส่วนเกิน) เท่านั้นที่วนไปเป็นชิ้นตั้งต้นของรอบถัดไป
 * สะสมยอด "พร้อมออกปลูก" ของทุกรอบไว้เทียบกับ targetPlants
 * ถ้ามี targetDate: จำนวนรอบทั้งหมด = floor((targetDate - startDate) / cycleDays) เสมอ (ปัดลง ไม่ปัดขึ้น) — รอบที่ "เริ่ม" แล้วแต่ยังไม่ครบ
 * cycleDays วันก่อนถึง targetDate จะไม่ถูกนับเป็นรอบที่จำลอง (ต้องครบรอบเต็มๆ เท่านั้นถึงนับ)
 * แยกเป็นฟังก์ชันกลาง (รับค่าที่ parse แล้วเท่านั้น ไม่แตะ payload ตรงๆ) เพื่อให้ simulateProduction() (หน้าจำลอง standalone)
 * และ computeDashboardData() (พยากรณ์อัตโนมัติต่อชนิดพืชใน Dashboard, ตั้ง acclimatizationRate = 100% เสมอเพราะ Dashboard ไม่มีช่องกรอกอัตรารอด) ใช้ตรรกะเดียวกันไม่ซ้ำโค้ด
 */
function runProductionCycles(params) {
  const {
    initialPieces, startDate, targetDate, targetPlants, cycleDays,
    ssBottleLimit, ssPiecesPerBottle, ssMultiplier,
    tibBottleLimit, tibPiecesPerBottle, tibMultiplier,
    acclimatizationRate,
  } = params;

  const ssCapacity = ssBottleLimit * ssPiecesPerBottle;
  const tibCapacity = tibBottleLimit * tibPiecesPerBottle;
  // จำนวนรอบสูงสุดที่อนุญาตตาม targetDate — ปัดลงเสมอ (รอบที่ยังไม่ครบ cycleDays วันก่อนถึง targetDate ไม่นับ)
  const maxCyclesByDate = targetDate
    ? Math.floor(((targetDate - new Date(startDate)) / (24 * 60 * 60 * 1000) + 1e-9) / cycleDays)
    : null;

  const rows = [];
  let inputPieces = initialPieces;
  let cycleDate = new Date(startDate);
  let cycle = 0;
  let reachedAt = null;
  let cumulativeReady = 0; // สะสมจำนวน "พร้อมออกปลูก" ที่หักออกไปจริงทุกรอบ (ไม่ใช่ตัวเลขสมมติ)
  let pendingRooting = 0; // ของที่หยิบออกไปกระตุ้นรากรอบที่แล้ว ยังไม่แปลงเป็น "พร้อมออกปลูก" จนกว่าจะถึงรอบนี้
  const MAX_CYCLES = 500; // กันลูปไม่จบ (500 รอบ x 15 วัน ~ 20 ปี เกินพอสำหรับการวางแผนจริง)

  while (cycle < MAX_CYCLES) {
    if (maxCyclesByDate !== null && cycle >= maxCyclesByDate) break;
    cycle++;

    // ดัน TIB ให้เต็มความจุก่อนเสมอ (ตัวคูณสูงกว่า SS) เหลือเท่าไหร่ค่อยลง SS
    const tibLoaded = Math.min(inputPieces, tibCapacity);
    const remainingAfterTib = inputPieces - tibLoaded;
    const ssLoaded = Math.min(remainingAfterTib, ssCapacity);
    const ssBottlesUsed = ssPiecesPerBottle > 0 ? Math.ceil(ssLoaded / ssPiecesPerBottle) : 0;
    const tibBottlesUsed = tibPiecesPerBottle > 0 ? Math.ceil(tibLoaded / tibPiecesPerBottle) : 0;
    const ssOutput = ssLoaded * ssMultiplier;
    const tibOutput = tibLoaded * tibMultiplier;
    // leftover = ส่วนเกินที่ไม่มีที่ลง SS/TIB เพราะระบบเต็มแล้ว (เป็น 0 เสมอถ้าชิ้นตั้งต้นยังไม่เกินความจุรวม) — ถูกนำออกไปกระตุ้นรากทันที ไม่วนกลับเข้าลูป
    const leftover = inputPieces - ssLoaded - tibLoaded;
    const grossOutput = ssOutput + tibOutput;
    const totalAfterCycle = leftover + grossOutput;
    const sentToRooting = leftover;
    const carriedForward = grossOutput; // เฉพาะผลผลิตที่คูณจาก SS/TIB เท่านั้นที่วนเป็นชิ้นตั้งต้นรอบหน้า ส่วนเกินไม่กลับมาขยายอีก
    // acclimatizationRate = อัตรา "รอดหลังออกปลูก" — ใช้แปลง pendingRooting ของรอบที่แล้ว (ไม่ใช่ sentToRooting ของรอบนี้) เพราะกระตุ้นรากใช้เวลา 1 รอบ
    const readyForPlanting = pendingRooting * acclimatizationRate;
    cumulativeReady += readyForPlanting;

    rows.push({
      cycle,
      date: formatThaiDate(cycleDate),
      input_pieces: Math.round(inputPieces),
      ss_capacity: ssCapacity,
      ss_bottles: ssBottlesUsed,
      ss_loaded_pieces: Math.round(ssLoaded),
      ss_multiplication_factor: ssMultiplier,
      ss_output_pieces: Math.round(ssOutput),
      tib_capacity: tibCapacity,
      tib_bottles: tibBottlesUsed,
      tib_loaded_pieces: Math.round(tibLoaded),
      tib_multiplication_factor: tibMultiplier,
      tib_output_pieces: Math.round(tibOutput),
      leftover_pieces: Math.round(leftover),
      gross_output_pieces: Math.round(grossOutput),
      total_pieces_after_cycle: Math.round(totalAfterCycle),
      ready_for_planting_pieces: Math.round(readyForPlanting),
      ready_for_planting_cumulative: Math.round(cumulativeReady),
    });

    if (!reachedAt && targetPlants !== null && cumulativeReady >= targetPlants) {
      reachedAt = { cycle, date: formatThaiDate(cycleDate) };
    }

    pendingRooting = sentToRooting; // ของรอบนี้ จะไปแปลงเป็น "พร้อมออกปลูก" ในรอบหน้า
    inputPieces = carriedForward;
    const next = new Date(cycleDate); next.setDate(next.getDate() + cycleDays);
    cycleDate = next;

    if (reachedAt) break; // ถึงเป้าหมายแล้ว ไม่ต้องจำลองรอบถัดไปอีก
  }

  const lastRow = rows[rows.length - 1] || null;
  const summary = {
    total_cycles: rows.length,
    pending_rooting: Math.round(pendingRooting), // ของรอบสุดท้ายที่ยังค้างอยู่ระยะกระตุ้นราก ยังไม่ถูกนับเป็น "พร้อมออกปลูก" (เพราะ "รอบหน้า" ไม่ถูกจำลองแล้ว)
    total_ready_for_planting: lastRow ? lastRow.ready_for_planting_cumulative : 0,
    reached_target: targetPlants !== null ? !!reachedAt : null,
    reached_cycle: reachedAt ? reachedAt.cycle : null,
    reached_date: reachedAt ? reachedAt.date : null,
    shortfall: (targetPlants !== null && !reachedAt && lastRow)
      ? Math.max(0, Math.round(targetPlants - lastRow.ready_for_planting_cumulative)) : null,
  };

  return { rows, summary };
}

/** Handler ของหน้าจำลอง standalone — parse payload (รับค่าจากผู้ใช้ทับค่า default ของระบบได้ทุกช่อง) แล้วเรียก runProductionCycles() */
function simulateProduction(payload) {
  if (!payload.initial_pieces || Number(payload.initial_pieces) <= 0) {
    throw new Error("กรุณาใส่จำนวนชิ้นตั้งต้น");
  }
  const targetDate = payload.target_date ? new Date(payload.target_date) : null;
  const targetPlants = (payload.target_plants !== undefined && payload.target_plants !== "")
    ? Number(payload.target_plants) : null;
  if (!targetDate && targetPlants === null) {
    throw new Error("กรุณาระบุวันที่ต้องการเนื้อเยื่อหรือจำนวนต้นเป้าหมายอย่างน้อยหนึ่งอย่าง");
  }

  const cfg = getSystemConfig();
  const startDate = payload.start_date ? new Date(payload.start_date) : new Date();

  return runProductionCycles({
    initialPieces: Number(payload.initial_pieces),
    startDate,
    targetDate,
    targetPlants,
    cycleDays: Number(payload.cycle_days) || cfg.cycle_days,
    ssBottleLimit: Number(payload.ss_bottle_limit) || cfg.ss_bottle_limit,
    ssPiecesPerBottle: Number(payload.ss_pieces_per_bottle) || cfg.ss_pieces_per_bottle,
    ssMultiplier: Number(payload.ss_multiplication_factor) || cfg.ss_multiplication_factor,
    tibBottleLimit: Number(payload.tib_bottle_limit) || cfg.tib_bottle_limit,
    tibPiecesPerBottle: Number(payload.tib_pieces_per_bottle) || cfg.tib_pieces_per_bottle,
    tibMultiplier: Number(payload.tib_multiplication_factor) || cfg.tib_multiplication_factor,
    acclimatizationRate: (payload.acclimatization_rate !== undefined && payload.acclimatization_rate !== "") ? Number(payload.acclimatization_rate) : 1,
  });
}

/**
 * พยากรณ์จำนวนต้นของล็อตหนึ่ง ณ วันที่ targetDate ที่ต้องการ โดยใช้ตัวคูณการขยายต่อชนิดพืช (stage1_multiplier_tib/ss)
 * ปัจจุบัน computeDashboardData ไม่ได้เรียกฟังก์ชันนี้แล้ว (เปลี่ยนไปใช้ runProductionCycles แบบจำลอง SS/TIB แทน)
 * เก็บไว้เผื่อใช้งานในอนาคต — ฟิลด์ตัวคูณการขยายในหน้าตั้งค่ายังกรอก/แก้ได้ตามปกติ แค่ไม่ถูกใช้คำนวณ Dashboard แล้ว
 */
function predictQuantityForLot(lot, plant, targetDate) {
  let quantity = Number(lot.quantity || 0);
  let roundNo = Number(lot.round_no || 0);
  let nextDate = lot.next_transfer_date ? parseThaiDateApprox(lot.next_transfer_date) : null;
  if (!nextDate) return quantity; // ล็อตที่จบรอบติดตามไปแล้ว ไม่มีรอบเพิ่มอีก

  const stage1Rounds = Number(plant.stage1_rounds) || 1;
  const stage1Days = Number(plant.stage1_days) || 0;
  const stage2Days = Number(plant.stage2_days) || 0;
  const isTib = lot.media_type === "TIB";
  // ค่าตัวคูณในชีตเป็น "" (ไม่ใช่ undefined) ตอนยังไม่กรอก — ต้อง || 1 ก่อน Number() ไม่งั้น Number("") จะได้ 0 ทำให้คูณทุกอย่างเป็น 0
  const multiplier = Number((isTib ? plant.stage1_multiplier_tib : plant.stage1_multiplier_ss) || 1) || 1;

  let safety = 0; // กันลูปไม่จบถ้า stage1_days/stage2_days ถูกตั้งเป็น 0 โดยไม่ตั้งใจ
  while (nextDate && nextDate <= targetDate && safety < 1000) {
    safety++;
    roundNo++;
    if (roundNo <= stage1Rounds) {
      quantity = quantity * multiplier;
      if (stage1Days <= 0) break;
      const d = new Date(nextDate); d.setDate(d.getDate() + stage1Days);
      nextDate = d;
    } else if (roundNo === stage1Rounds + 1) {
      if (stage2Days <= 0) { nextDate = null; break; }
      const d = new Date(nextDate); d.setDate(d.getDate() + stage2Days);
      nextDate = d;
    } else {
      nextDate = null; // เข้าระยะพร้อมออกปลูก ไม่มีรอบถัดไปอีก หยุดจำลอง
    }
  }
  return quantity;
}

function nextRecipeFor(plants, transferRow) {
  const plant = plants.find(p => p.plant_id === transferRow.plant_id);
  if (!plant) return "";
  const stage1Rounds = Number(plant.stage1_rounds) || 1;
  const nextRound = Number(transferRow.round_no) + 1;
  const isTib = transferRow.media_type === "TIB";
  if (nextRound <= stage1Rounds) return isTib ? plant.stage1_recipe_tib : plant.stage1_recipe_ss;
  if (nextRound === stage1Rounds + 1) return isTib ? plant.stage2_recipe_tib : plant.stage2_recipe_ss;
  return isTib ? plant.stage3_recipe_tib : plant.stage3_recipe_ss;
}

function parseThaiDateApprox(thaiDateStr) {
  const months = { "ม.ค.": 0, "ก.พ.": 1, "มี.ค.": 2, "เม.ย.": 3, "พ.ค.": 4, "มิ.ย.": 5, "ก.ค.": 6, "ส.ค.": 7, "ก.ย.": 8, "ต.ค.": 9, "พ.ย.": 10, "ธ.ค.": 11 };
  const parts = thaiDateStr.split(" ");
  if (parts.length < 3) return new Date(8640000000000000);
  const day = Number(parts[0]);
  const month = months[parts[1]];
  const buddhistYear2digit = Number(parts[2]);
  const fullYear = 2500 + buddhistYear2digit - 543;
  return new Date(fullYear, month, day);
}

// ===================================================================
// แจ้งเตือน Line รายวัน — ตั้ง Time-driven trigger ให้รันฟังก์ชันนี้ทุกเช้า (เช่น 7:00)
// ===================================================================
/**
 * ผู้ใช้แต่ละคนได้รับแจ้งเตือนรายวันแยกตาม Lab ของตัวเอง — เห็นเฉพาะของ Lab ตัวเอง
 * แอดมินได้รับสรุปรวมทุก Lab เป็นข้อความเดียว (ไม่ได้รับซ้ำกับสรุปของ Lab ตัวเองอีกรอบ)
 */
function sendDueReminders() {
  const allTransfers = getSheetData("Transfers");
  const plants = getSheetData("PlantTypes");
  const storage = getSheetData("Storage").filter(s => s.active !== false);
  const users = getSheetData("LineUsers").filter(u => u.active !== false);
  const now = new Date();
  const tomorrow = new Date(now); tomorrow.setDate(tomorrow.getDate() + 1);

  function dueTodayTomorrowOf(transfersForScope) {
    const data = computeDashboardData(transfersForScope, plants, storage);
    return data.due_soon.filter(d => {
      const due = parseThaiDateApprox(d.next_transfer_date);
      return isSameDay(due, now) || isSameDay(due, tomorrow);
    });
  }

  // ส่งให้สมาชิกแต่ละ Lab (ที่ไม่ใช่แอดมิน) เห็นแค่ของ Lab ตัวเอง
  LABS.forEach(lab => {
    const due = dueTodayTomorrowOf(allTransfers.filter(t => t.lab === lab));
    if (due.length === 0) return;
    const lines = due.map(d =>
      `• ${d.plant_name} — ล็อต ${d.lot_no} — ชั้น ${d.shelf}/${d.sub_shelf}\n  สูตรรอบหน้า: ${d.next_recipe || "-"}`
    );
    const message = `🌱 แจ้งเตือนถ่ายโอนเนื้อเยื่อ — Lab ${lab}\nวันนี้/พรุ่งนี้ถึงกำหนด ${due.length} ล็อต:\n\n${lines.join("\n\n")}`;
    const recipients = users.filter(u => u.role !== "admin" && u.lab === lab).map(u => u.line_user_id);
    if (recipients.length > 0) pushLineMessage(recipients, message);
  });

  // แอดมิน: สรุปรวมทุก Lab ในข้อความเดียว
  const dueAll = dueTodayTomorrowOf(allTransfers.filter(t => !!t.lab));
  if (dueAll.length > 0) {
    const adminIds = users.filter(u => u.role === "admin").map(u => u.line_user_id);
    if (adminIds.length > 0) {
      const lines = dueAll.map(d =>
        `• Lab ${d.lab} — ${d.plant_name} — ล็อต ${d.lot_no} — ชั้น ${d.shelf}/${d.sub_shelf}\n  สูตรรอบหน้า: ${d.next_recipe || "-"}`
      );
      const message = `🌱 แจ้งเตือนถ่ายโอนเนื้อเยื่อ — ทุก Lab\nวันนี้/พรุ่งนี้ถึงกำหนด ${dueAll.length} ล็อต:\n\n${lines.join("\n\n")}`;
      pushLineMessage(adminIds, message);
    }
  }
}

function isSameDay(a, b) {
  return a.getFullYear() === b.getFullYear() && a.getMonth() === b.getMonth() && a.getDate() === b.getDate();
}

function pushLineMessage(userIds, text) {
  const token = PropertiesService.getScriptProperties().getProperty("LINE_CHANNEL_ACCESS_TOKEN");
  if (!token) throw new Error("ยังไม่ได้ตั้งค่า LINE_CHANNEL_ACCESS_TOKEN ใน Script Properties");
  UrlFetchApp.fetch("https://api.line.me/v2/bot/message/multicast", {
    method: "post",
    contentType: "application/json",
    headers: { Authorization: "Bearer " + token },
    payload: JSON.stringify({ to: userIds, messages: [{ type: "text", text }] }),
    muteHttpExceptions: true,
  });
}
