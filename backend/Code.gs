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
    "stage2_days", "stage2_recipe_tib", "stage2_recipe_ss",
    "stage3_recipe_tib", "stage3_recipe_ss",
    "active",
  ],
  Storage: ["type", "value", "active"],
  Transfers: [
    "record_id", "plant_id", "plant_name", "lot_no", "lab", "round_no", "transfer_date", "quantity",
    "media_type", "bottle_no", "shelf", "sub_shelf",
    "stage", "recipe_used", "next_transfer_date", "recorded_by", "created_at",
  ],
  LineUsers: ["line_user_id", "display_name", "role", "lab", "active", "approved"],
};

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
  Logger.log("ติดตั้งชีตเรียบร้อย");
}

/**
 * รันเมื่อมีการอัปเดตโค้ดที่เพิ่มคอลัมน์ใหม่ (เช่นตอนนี้) แล้วชีตเดิมมีข้อมูลอยู่แล้ว
 * จะเติมคอลัมน์ที่ขาดไปต่อท้ายคอลัมน์เดิม โดยไม่แก้ไข/ลบข้อมูลที่มีอยู่
 * ปลอดภัย รันซ้ำได้หลายครั้งไม่มีผลเสีย
 */
function migrateSheets() {
  const ss = getSS();
  Object.keys(SHEET_SCHEMAS).forEach(name => {
    const sheet = ss.getSheetByName(name);
    if (!sheet) return;
    const lastCol = Math.max(sheet.getLastColumn(), 1);
    const existingHeaders = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    const missing = SHEET_SCHEMAS[name].filter(h => existingHeaders.indexOf(h) === -1);
    if (missing.length > 0) {
      sheet.getRange(1, lastCol + 1, 1, missing.length).setValues([missing]);
      Logger.log(`เพิ่มคอลัมน์ใหม่ในชีต ${name}: ${missing.join(", ")}`);
    }
  });
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
  return plants.map(p => {
    const latestByLot = {};
    transfers.filter(t => t.plant_id === p.plant_id).forEach(t => {
      const key = t.lab + "|" + t.lot_no;
      if (!latestByLot[key] || new Date(t.transfer_date) > new Date(latestByLot[key].transfer_date)) {
        latestByLot[key] = t;
      }
    });
    const activeCount = Object.values(latestByLot).filter(t => !!t.next_transfer_date).length;
    return { plant_id: p.plant_id, plant_name: p.plant_name, active_count: activeCount };
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
  const required = ["plant_id", "lot_no", "quantity", "shelf", "sub_shelf", "media_type"];
  required.forEach(f => {
    if (payload[f] === undefined || payload[f] === "") throw new Error("กรุณากรอกข้อมูลให้ครบ: " + f);
  });
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
    if (t.next_transfer_date && (!previous || new Date(t.transfer_date) > new Date(previous.transfer_date))) {
      previous = t;
    }
  });
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

  const record = {
    record_id: genId("rec"),
    plant_id: plant.plant_id,
    plant_name: plant.plant_name,
    lot_no: payload.lot_no,
    lab: me.lab,
    round_no,
    transfer_date: formatThaiDate(today),
    quantity: payload.quantity,
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
function computeDashboardData(transfers, plants, storage) {
  // นับ "ล็อต" ที่กำลังติดตามอยู่ ไม่ใช่ตำแหน่งเก็บ — แถวเก่าที่ไม่มี lot_no/lab (ก่อนอัปเดตฟีเจอร์นี้) ถูกข้าม
  const latestByLot = {};
  transfers.forEach(t => {
    if (!t.lot_no) return;
    const key = t.plant_id + "|" + t.lab + "|" + t.lot_no;
    if (!latestByLot[key] || new Date(t.transfer_date) > new Date(latestByLot[key].transfer_date)) {
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

  const by_plant = plants.filter(p => p.active !== false).map(p => {
    const lots = activeLots.filter(t => t.plant_id === p.plant_id);
    const stage_counts = { "ขยาย": 0, "กระตุ้นราก": 0, "พร้อมออกปลูก": 0 };
    lots.forEach(t => { stage_counts[t.stage] = (stage_counts[t.stage] || 0) + 1; });
    const dominant_stage = Object.keys(stage_counts).reduce((a, b) => stage_counts[a] >= stage_counts[b] ? a : b, "ขยาย");
    return { plant_name: p.plant_name, total: lots.length, stage_counts, dominant_stage };
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
  const data = computeDashboardData(transfers, plants, storage);
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
    if (!latestByLot[key] || new Date(t.transfer_date) > new Date(latestByLot[key].transfer_date)) {
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
    if (!current || new Date(t.transfer_date) > new Date(current.transfer_date)) current = t;
  });
  if (!current) return [];

  const chain = [current];
  let cursor = current;
  while (Number(cursor.round_no) > 1) {
    const targetRound = Number(cursor.round_no) - 1;
    const prevCandidates = transfers.filter(t =>
      Number(t.round_no) === targetRound && new Date(t.transfer_date) <= new Date(cursor.transfer_date)
    );
    if (prevCandidates.length === 0) break;
    prevCandidates.sort((a, b) => new Date(b.transfer_date) - new Date(a.transfer_date));
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
