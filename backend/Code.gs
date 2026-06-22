/**
 * backend/Code.gs
 * Google Apps Script — ทำหน้าที่เป็น Backend API + ฐานข้อมูล (Google Sheet) + ตัวส่งแจ้งเตือน Line
 *
 * วิธีติดตั้ง: อ่าน README.md ในโฟลเดอร์โปรเจกต์ทั้งหมดก่อนเริ่ม
 *
 * ก่อนใช้งาน ต้องตั้งค่า Script Properties (Project Settings > Script Properties):
 *   SPREADSHEET_ID              -> ID ของ Google Sheet ที่จะใช้เป็นฐานข้อมูล (จำเป็น)
 *   LINE_CHANNEL_ACCESS_TOKEN   -> Channel access token ของ LINE Messaging API (จำเป็น)
 *   ADMIN_LINE_USER_ID          -> LINE user ID ของแอดมิน/PI ที่จะได้รับแจ้งเตือนทันทีทุกครั้งที่มีการบันทึก (ไม่ใส่ก็ได้ ถ้าไม่ใส่จะไม่มีการแจ้งเตือนทันที)
 *     หา LINE user ID ของตัวเองได้จากชีต "LineUsers" หลังจากเปิด LIFF ผ่าน Line ไปแล้วอย่างน้อย 1 ครั้ง
 *
 * แล้วรันฟังก์ชัน setupSheets() หนึ่งครั้งจากใน Apps Script editor เพื่อสร้างชีตและหัวคอลัมน์ให้ครบ
 * ถ้าเคยรัน setupSheets() ไปแล้วก่อนหน้า (มีชีตอยู่แล้ว) ให้รัน migrateSheets() แทน เพื่อเพิ่มคอลัมน์ใหม่โดยไม่ลบข้อมูลเดิม
 */

const STAGE_NAMES = ["ขยาย", "กระตุ้นราก", "พร้อมออกปลูก"];
const MEDIA_TYPES = ["TIB", "SS"];

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
    "record_id", "plant_id", "plant_name", "round_no", "transfer_date", "quantity",
    "media_type", "bottle_no", "shelf", "sub_shelf",
    "stage", "recipe_used", "next_transfer_date", "recorded_by", "created_at",
  ],
  LineUsers: ["line_user_id", "display_name", "active"],
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
      listPlantsWithCount,
      listPlantsFull,
      savePlantType,
      deactivatePlantType,
      listStorage,
      addStorageValue,
      deactivateStorageValue,
      recordTransfer,
      listLineUsers,
      getDashboard,
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
  if (existing) {
    if (!existing.active || existing.display_name !== payload.display_name) {
      updateRow("LineUsers", existing._row, { line_user_id: payload.line_user_id, display_name: payload.display_name, active: true });
    }
    return { updated: true };
  }
  appendRow("LineUsers", { line_user_id: payload.line_user_id, display_name: payload.display_name, active: true });
  return { created: true };
}

function listLineUsers() {
  return getSheetData("LineUsers").map(u => ({ line_user_id: u.line_user_id, display_name: u.display_name, active: u.active }));
}

// ===================================================================
// Handlers: ชนิดพืช
// ===================================================================
function listPlantsFull() {
  return getSheetData("PlantTypes");
}

function listPlantsWithCount() {
  const plants = getSheetData("PlantTypes").filter(p => p.active !== false);
  const transfers = getSheetData("Transfers");
  return plants.map(p => {
    const latestByPosition = {};
    transfers.filter(t => t.plant_id === p.plant_id).forEach(t => {
      const key = t.shelf + "|" + t.sub_shelf;
      if (!latestByPosition[key] || new Date(t.transfer_date) > new Date(latestByPosition[key].transfer_date)) {
        latestByPosition[key] = t;
      }
    });
    const activeCount = Object.values(latestByPosition).filter(t => !!t.next_transfer_date).length;
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
  const required = ["plant_id", "quantity", "shelf", "sub_shelf", "media_type"];
  required.forEach(f => {
    if (payload[f] === undefined || payload[f] === "") throw new Error("กรุณากรอกข้อมูลให้ครบ: " + f);
  });
  if (MEDIA_TYPES.indexOf(payload.media_type) === -1) {
    throw new Error("ประเภทรอบต้องเป็น TIB หรือ SS");
  }
  if (payload.media_type === "TIB" && !payload.bottle_no) {
    throw new Error("กรุณาใส่หมายเลขขวด (จำเป็นสำหรับ TIB)");
  }

  const plants = getSheetData("PlantTypes");
  const plant = plants.find(p => p.plant_id === payload.plant_id);
  if (!plant) throw new Error("ไม่พบชนิดพืชนี้");

  const transfers = getSheetData("Transfers");
  const sameSpot = transfers.filter(t =>
    t.plant_id === payload.plant_id && String(t.shelf) === String(payload.shelf) && String(t.sub_shelf) === String(payload.sub_shelf)
  );
  let previous = null;
  sameSpot.forEach(t => {
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

  notifyAdminNewRecord(record, round_in_stage);

  return {
    stage,
    recipe_used,
    next_transfer_date,
    round_in_stage,
    media_type: record.media_type,
    bottle_no: record.bottle_no,
  };
}

function notifyAdminNewRecord(record, round_in_stage) {
  try {
    const adminId = PropertiesService.getScriptProperties().getProperty("ADMIN_LINE_USER_ID");
    if (!adminId) return;
    const lines = [
      "📝 มีการบันทึกถ่ายโอนเนื้อเยื่อใหม่",
      `${record.plant_name} — ${record.stage} (รอบที่ ${round_in_stage})`,
      `ประเภท: ${record.media_type}${record.bottle_no ? " · ขวดเลขที่ " + record.bottle_no : ""}`,
      `ที่เก็บ: ชั้น ${record.shelf}/${record.sub_shelf} · จำนวน ${record.quantity} ต้น`,
      `บันทึกโดย: ${record.recorded_by}`,
    ];
    pushLineMessage([adminId], lines.join("\n"));
  } catch (err) {
    Logger.log("แจ้งเตือนแอดมินไม่สำเร็จ: " + err.message);
  }
}

// ===================================================================
// Handler: Dashboard
// ===================================================================
function getDashboard() {
  const transfers = getSheetData("Transfers");
  const plants = getSheetData("PlantTypes");
  const storage = getSheetData("Storage").filter(s => s.active !== false);

  const latestByPosition = {};
  transfers.forEach(t => {
    const key = t.plant_id + "|" + t.shelf + "|" + t.sub_shelf;
    if (!latestByPosition[key] || new Date(t.transfer_date) > new Date(latestByPosition[key].transfer_date)) {
      latestByPosition[key] = t;
    }
  });
  const activeBottles = Object.values(latestByPosition).filter(t => !!t.next_transfer_date);

  const total_bottles = activeBottles.length;
  const total_plants = activeBottles.reduce((sum, t) => sum + Number(t.quantity || 0), 0);

  const now = new Date();
  const withDays = activeBottles
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
      shelf: t.shelf,
      sub_shelf: t.sub_shelf,
      media_type: t.media_type,
      bottle_no: t.bottle_no,
      stage: t.stage,
      round_in_stage,
      next_transfer_date: t.next_transfer_date,
      next_recipe: nextRecipeFor(plants, t),
    };
  });

  const by_plant = plants.filter(p => p.active !== false).map(p => {
    const bottles = activeBottles.filter(t => t.plant_id === p.plant_id);
    const stage_counts = { "ขยาย": 0, "กระตุ้นราก": 0, "พร้อมออกปลูก": 0 };
    bottles.forEach(t => { stage_counts[t.stage] = (stage_counts[t.stage] || 0) + 1; });
    const dominant_stage = Object.keys(stage_counts).reduce((a, b) => stage_counts[a] >= stage_counts[b] ? a : b, "ขยาย");
    return { plant_name: p.plant_name, total: bottles.length, stage_counts, dominant_stage };
  });

  const shelves = storage.filter(s => s.type === "shelf").map(s => s.value);
  const subs = storage.filter(s => s.type === "sub_shelf").map(s => s.value);
  const storage_overview = [];
  shelves.forEach(shelf => {
    subs.forEach(sub => {
      const count = activeBottles.filter(t => String(t.shelf) === String(shelf) && String(t.sub_shelf) === String(sub)).length;
      storage_overview.push({ shelf, sub_shelf: sub, count });
    });
  });

  return { total_bottles, total_plants, due_soon, by_plant, storage_overview };
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
function sendDueReminders() {
  const dashboard = getDashboard();
  const now = new Date();
  const tomorrow = new Date(now); tomorrow.setDate(tomorrow.getDate() + 1);
  const dueTodayTomorrow = dashboard.due_soon.filter(d => {
    const due = parseThaiDateApprox(d.next_transfer_date);
    return isSameDay(due, now) || isSameDay(due, tomorrow);
  });
  if (dueTodayTomorrow.length === 0) return;

  const lines = dueTodayTomorrow.map(d =>
    `• ${d.plant_name} — ชั้น ${d.shelf}/${d.sub_shelf}\n  สูตรรอบหน้า: ${d.next_recipe || "-"}`
  );
  const message = `🌱 แจ้งเตือนถ่ายโอนเนื้อเยื่อ\nวันนี้/พรุ่งนี้ถึงกำหนด ${dueTodayTomorrow.length} ขวด:\n\n${lines.join("\n\n")}`;

  const users = getSheetData("LineUsers").filter(u => u.active !== false).map(u => u.line_user_id);
  if (users.length === 0) return;
  pushLineMessage(users, message);
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
