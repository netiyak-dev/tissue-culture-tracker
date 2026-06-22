/**
 * backend/Code.gs
 * Google Apps Script — ทำหน้าที่เป็น Backend API + ฐานข้อมูล (Google Sheet) + ตัวส่งแจ้งเตือน Line
 *
 * วิธีติดตั้ง: อ่าน README.md ในโฟลเดอร์โปรเจกต์ทั้งหมดก่อนเริ่ม
 *
 * ก่อนใช้งาน ต้องตั้งค่า Script Properties 2 ค่า (Project Settings > Script Properties):
 *   SPREADSHEET_ID              -> ID ของ Google Sheet ที่จะใช้เป็นฐานข้อมูล
 *   LINE_CHANNEL_ACCESS_TOKEN   -> Channel access token ของ LINE Messaging API
 *
 * แล้วรันฟังก์ชัน setupSheets() หนึ่งครั้งจากใน Apps Script editor เพื่อสร้างชีตและหัวคอลัมน์ให้ครบ
 */

const STAGE_NAMES = ["ขยาย", "กระตุ้นราก", "พร้อมออกปลูก"];

// ===================================================================
// เข้าถึง Spreadsheet
// ===================================================================
function getSS() {
  const id = PropertiesService.getScriptProperties().getProperty("SPREADSHEET_ID");
  if (!id) throw new Error("ยังไม่ได้ตั้งค่า SPREADSHEET_ID ใน Script Properties");
  return SpreadsheetApp.openById(id);
}

const SHEET_SCHEMAS = {
  PlantTypes: ["plant_id", "plant_name", "stage1_days", "stage1_recipe", "stage2_days", "stage2_recipe", "stage3_recipe", "active"],
  Storage: ["type", "value", "active"],
  Transfers: ["record_id", "plant_id", "plant_name", "round_no", "transfer_date", "quantity", "shelf", "sub_shelf", "stage", "recipe_used", "next_transfer_date", "recorded_by", "created_at"],
  LineUsers: ["line_user_id", "display_name", "active"],
};

/** รันครั้งเดียวตอนติดตั้งระบบ: สร้างชีต + หัวคอลัมน์ที่ขาด (ไม่ลบของเดิม) */
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
  // ใส่ค่าเริ่มต้นที่เก็บ ชั้น 1-4 / ช่อง A-C ถ้ายังไม่มี
  const storageRows = getSheetData("Storage");
  if (storageRows.length === 0) {
    ["1", "2", "3", "4"].forEach(v => appendRow("Storage", { type: "shelf", value: v, active: true }));
    ["A", "B", "C"].forEach(v => appendRow("Storage", { type: "sub_shelf", value: v, active: true }));
  }
  Logger.log("ติดตั้งชีตเรียบร้อย");
}

// ===================================================================
// อ่าน/เขียนชีตแบบเป็น object array (อิงหัวคอลัมน์)
// ===================================================================
function getSheet(name) {
  const sheet = getSS().getSheetByName(name);
  if (!sheet) throw new Error(`ไม่พบชีต ${name} — รัน setupSheets() ก่อน`);
  return sheet;
}

function getSheetData(name) {
  const sheet = getSheet(name);
  const range = sheet.getDataRange().getValues();
  if (range.length < 2) return [];
  const headers = range[0];
  return range.slice(1)
    .filter(row => row.some(cell => cell !== ""))
    .map((row, idx) => {
      const obj = { _row: idx + 2 }; // เลขแถวจริงในชีต (เผื่อใช้แก้ไข)
      headers.forEach((h, i) => (obj[h] = row[i]));
      return obj;
    });
}

function appendRow(name, obj) {
  const sheet = getSheet(name);
  const headers = SHEET_SCHEMAS[name];
  const row = headers.map(h => (obj[h] !== undefined ? obj[h] : ""));
  sheet.appendRow(row);
  return obj;
}

function updateRow(name, rowIndex, obj) {
  const sheet = getSheet(name);
  const headers = SHEET_SCHEMAS[name];
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
      listStorage,
      addStorageValue,
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
  return getSheetData("LineUsers").map(u => ({ display_name: u.display_name, active: u.active }));
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
    // นับขวดที่ยัง "ตามอยู่" ของชนิดนี้ = แถวล่าสุดต่อตำแหน่ง(shelf+sub_shelf) ที่ยังมี next_transfer_date (ยังไม่ถึงระยะสุดท้าย)
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
  const required = ["plant_name", "stage1_days", "stage1_recipe", "stage2_days", "stage2_recipe", "stage3_recipe"];
  required.forEach(f => {
    if (payload[f] === undefined || payload[f] === "") throw new Error("กรุณากรอกข้อมูลให้ครบ: " + f);
  });
  const plants = getSheetData("PlantTypes");
  if (payload.plant_id) {
    const existing = plants.find(p => p.plant_id === payload.plant_id);
    if (!existing) throw new Error("ไม่พบชนิดพืชนี้");
    updateRow("PlantTypes", existing._row, { ...payload, active: true });
    return { updated: true, plant_id: payload.plant_id };
  }
  const plant_id = genId("plant");
  appendRow("PlantTypes", { ...payload, plant_id, active: true });
  return { created: true, plant_id };
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
  if (dup) return { duplicated: true };
  appendRow("Storage", { type: payload.type, value: payload.value, active: true });
  return { created: true };
}

// ===================================================================
// Handler หลัก: บันทึกการถ่ายโอนเนื้อเยื่อ + คำนวณระยะ/สูตร/วันถัดไปอัตโนมัติ
//
// หมายเหตุการออกแบบ (สำคัญ — ตรวจสอบว่าตรงกับการใช้งานจริงหรือไม่):
// ระบบระบุ "ขวดเดิม" ด้วยคู่ (ชนิดพืช + ชั้น + ช่อง) เดียวกัน
// ถ้าตำแหน่งนี้มีประวัติล่าสุดที่ยังไม่ถึงระยะสุดท้าย -> รอบนี้คือรอบถัดไปของขวดเดิม
// ถ้าตำแหน่งนี้ว่าง หรือขวดก่อนหน้าถึงระยะสุดท้ายไปแล้ว -> เริ่มขวดใหม่ที่รอบ 1
// ===================================================================
function recordTransfer(payload) {
  const required = ["plant_id", "quantity", "shelf", "sub_shelf"];
  required.forEach(f => {
    if (payload[f] === undefined || payload[f] === "") throw new Error("กรุณากรอกข้อมูลให้ครบ: " + f);
  });

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

  const today = new Date();
  let stage, recipe_used, next_transfer_date = "";

  if (round_no === 1) {
    stage = STAGE_NAMES[0];
    recipe_used = plant.stage1_recipe;
    const next = new Date(today); next.setDate(next.getDate() + Number(plant.stage1_days));
    next_transfer_date = formatThaiDate(next);
  } else if (round_no === 2) {
    stage = STAGE_NAMES[1];
    recipe_used = plant.stage2_recipe;
    const next = new Date(today); next.setDate(next.getDate() + Number(plant.stage2_days));
    next_transfer_date = formatThaiDate(next);
  } else {
    stage = STAGE_NAMES[2];
    recipe_used = plant.stage3_recipe;
    next_transfer_date = ""; // จบรอบติดตาม
  }

  const record = {
    record_id: genId("rec"),
    plant_id: plant.plant_id,
    plant_name: plant.plant_name,
    round_no,
    transfer_date: formatThaiDate(today),
    quantity: payload.quantity,
    shelf: payload.shelf,
    sub_shelf: payload.sub_shelf,
    stage,
    recipe_used,
    next_transfer_date,
    recorded_by: payload.recorded_by || "ไม่ระบุชื่อ",
    created_at: today.toISOString(),
  };
  appendRow("Transfers", record);

  return { stage, recipe_used, next_transfer_date };
}

// ===================================================================
// Handler: Dashboard
// ===================================================================
function getDashboard() {
  const transfers = getSheetData("Transfers");
  const plants = getSheetData("PlantTypes");
  const storage = getSheetData("Storage").filter(s => s.active !== false);

  // เอาแถวล่าสุดต่อตำแหน่ง (plant+shelf+sub_shelf) = สถานะปัจจุบันของขวดนั้น
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

  // ใกล้ถึงกำหนด (ภายใน 3 วัน หรือเลยกำหนดแล้ว) เรียงจากใกล้สุด
  const now = new Date();
  const withDays = activeBottles
    .map(t => ({ ...t, _daysLeft: (parseThaiDateApprox(t.next_transfer_date) - now) / 86400000 }))
    .filter(t => t._daysLeft <= 3)
    .sort((a, b) => a._daysLeft - b._daysLeft)
    .slice(0, 20);

  const due_soon = withDays.map(t => ({
    plant_name: t.plant_name,
    shelf: t.shelf,
    sub_shelf: t.sub_shelf,
    next_transfer_date: t.next_transfer_date,
    next_recipe: nextRecipeFor(plants, t),
  }));

  // ภาพรวมตามชนิดพืช
  const by_plant = plants.filter(p => p.active !== false).map(p => {
    const bottles = activeBottles.filter(t => t.plant_id === p.plant_id);
    const stage_counts = { "ขยาย": 0, "กระตุ้นราก": 0, "พร้อมออกปลูก": 0 };
    bottles.forEach(t => { stage_counts[t.stage] = (stage_counts[t.stage] || 0) + 1; });
    const dominant_stage = Object.keys(stage_counts).reduce((a, b) => stage_counts[a] >= stage_counts[b] ? a : b, "ขยาย");
    return { plant_name: p.plant_name, total: bottles.length, stage_counts, dominant_stage };
  });

  // ภาพรวมที่เก็บ
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
  const nextRound = Number(transferRow.round_no) + 1;
  if (nextRound === 2) return plant.stage2_recipe;
  if (nextRound >= 3) return plant.stage3_recipe;
  return plant.stage1_recipe;
}

/** แปลงวันแบบไทยที่เราเก็บเอง (เช่น "12 ก.ค. 69") กลับเป็น Date คร่าวๆ สำหรับเทียบวัน */
function parseThaiDateApprox(thaiDateStr) {
  const months = { "ม.ค.": 0, "ก.พ.": 1, "มี.ค.": 2, "เม.ย.": 3, "พ.ค.": 4, "มิ.ย.": 5, "ก.ค.": 6, "ส.ค.": 7, "ก.ย.": 8, "ต.ค.": 9, "พ.ย.": 10, "ธ.ค.": 11 };
  const parts = thaiDateStr.split(" ");
  if (parts.length < 3) return new Date(8640000000000000); // ไกลมาก ถ้า parse ไม่ได้
  const day = Number(parts[0]);
  const month = months[parts[1]];
  const buddhistYear2digit = Number(parts[2]);
  const fullYear = 2500 + buddhistYear2digit - 543; // พ.ศ. 2 หลัก -> ค.ศ.
  return new Date(fullYear, month, day);
}

// ===================================================================
// แจ้งเตือน Line — ตั้ง Time-driven trigger ให้รันฟังก์ชันนี้ทุกเช้า (เช่น 7:00)
// Edit > Current project's triggers > Add Trigger > sendDueReminders > Time-driven > Day timer
// ===================================================================
function sendDueReminders() {
  const dashboard = getDashboard();
  // เอาเฉพาะที่ครบกำหนด "วันนี้" หรือ "พรุ่งนี้" (ไม่รวมที่เลยกำหนดไปหลายวันแล้ว เพื่อไม่ให้สแปม)
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
