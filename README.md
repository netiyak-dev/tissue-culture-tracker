# README — วิธีติดตั้งระบบให้ใช้งานได้จริง

ระบบนี้ประกอบด้วย 3 ส่วนที่ต้องเชื่อมกัน:
1. **Google Sheet + Google Apps Script** (ฐานข้อมูล + backend) — ไฟล์ใน `backend/`
2. **เว็บฟอร์ม LIFF** (เปิดผ่าน Line) — ไฟล์ใน `liff-app/`
3. **Dashboard** (เว็บสรุปข้อมูล) — ไฟล์ใน `dashboard/`

ทำตามลำดับขั้นตอนนี้ครับ (ใช้เวลาประมาณ 30–45 นาทีในการตั้งครั้งแรก)

---

## ขั้นที่ 1 — สร้าง Google Sheet + Apps Script

1. ไปที่ [sheets.google.com](https://sheets.google.com) สร้างไฟล์ใหม่ ตั้งชื่อ เช่น "ฐานข้อมูลเพาะเลี้ยงเนื้อเยื่อ"
2. คัดลอก **Spreadsheet ID** จาก URL (ส่วนที่อยู่ระหว่าง `/d/` กับ `/edit`)
3. เปิด **Extensions > Apps Script** จากในชีตนั้น
4. ลบโค้ดเดิมในไฟล์ `Code.gs` ทิ้ง แล้ววางโค้ดจากไฟล์ `backend/Code.gs` ของโปรเจกต์นี้ทั้งหมด
5. กดไอคอนเฟือง "Project Settings" → ที่ "Script Properties" กด "Add script property" เพิ่ม 2 ค่า:
   - `SPREADSHEET_ID` = ID ที่คัดลอกไว้ในขั้นตอนที่ 2
   - `LINE_CHANNEL_ACCESS_TOKEN` = (ทำขั้นที่ 2 ก่อนแล้วย้อนมาใส่ตรงนี้)
6. กลับมาที่หน้าโค้ด เลือกฟังก์ชัน `setupSheets` ที่ dropdown ด้านบน แล้วกด ▶ Run ครั้งหนึ่ง (จะขอสิทธิ์เข้าถึง Google Sheet ของคุณ — กด "อนุญาต")
   - หลังรันแล้ว เปิดกลับไปดูที่ Google Sheet จะเห็นชีต `PlantTypes`, `Storage`, `Transfers`, `LineUsers` ถูกสร้างขึ้นพร้อมหัวคอลัมน์
7. Deploy เป็น Web App: กด **Deploy > New deployment** → เลือกประเภท "Web app"
   - Execute as: **Me**
   - Who has access: **Anyone**
   - กด Deploy แล้วคัดลอก **Web app URL** ที่ได้ (จะใช้ในขั้นที่ 3)

> ทุกครั้งที่แก้โค้ดใน Code.gs ต้อง Deploy ใหม่ (Manage deployments > แก้ไข > Version ใหม่) ไม่งั้นของเก่าจะยังทำงานอยู่

---

## ขั้นที่ 2 — สร้าง LINE Official Account + LIFF

1. ไปที่ [LINE Developers Console](https://developers.line.biz/console/) เข้าสู่ระบบด้วยบัญชี Line
2. สร้าง **Provider** ใหม่ (ถ้ายังไม่มี) แล้วสร้าง **Channel** ประเภท "Messaging API"
3. ในหน้า Channel: ไปที่แท็บ **Messaging API** → คัดลอก **Channel access token** (กด Issue ถ้ายังไม่มี) → เอาไปใส่ใน Script Properties `LINE_CHANNEL_ACCESS_TOKEN` ตามขั้นที่ 1.5
4. ไปที่แท็บ **LIFF** → กด **Add** สร้างแอป LIFF ใหม่:
   - Endpoint URL: ใส่ URL ของหน้า `liff-app/index.html` ที่คุณจะ host (ดูขั้นที่ 3 ก่อน แล้วย้อนมากรอกตรงนี้)
   - Size: **Full**
   - Scope: `profile`
   - กด Add → จะได้ **LIFF ID**
5. นำ LIFF ID ไปแก้ในไฟล์ `liff-app/js/api.js` ที่ตัวแปร `LIFF_ID`
6. (แนะนำ) สร้าง **Rich Menu** ในแท็บ "Messaging API" หรือผ่าน Line Official Account Manager — ใส่ปุ่ม 3 ปุ่ม ลิงก์ไปที่:
   - `https://liff.line.me/<LIFF_ID>` (หน้าหลัก)
   - `https://liff.line.me/<LIFF_ID>/record.html` (ถ้าต้องการเข้าฟอร์มตรง)
   - `https://liff.line.me/<LIFF_ID>/settings.html`

---

## ขั้นที่ 3 — Host ไฟล์หน้าเว็บ (liff-app + dashboard)

ไฟล์ในโฟลเดอร์ `liff-app/` และ `dashboard/` เป็น HTML/CSS/JS ธรรมดา ต้อง host ไว้ที่ใดที่หนึ่งที่เป็น **HTTPS** (LIFF บังคับ HTTPS) วิธีที่ง่ายที่สุดสำหรับมือใหม่:

**ตัวเลือก: GitHub Pages (ฟรี)**
1. สร้าง repository ใหม่บน GitHub อัปโหลดทั้งโฟลเดอร์โปรเจกต์ (รวม `shared/`, `liff-app/`, `dashboard/`)
2. ไปที่ Settings > Pages ของ repo → เปิดใช้ GitHub Pages จาก branch หลัก
3. จะได้ URL แบบ `https://ชื่อบัญชี.github.io/ชื่อrepo/liff-app/index.html`
4. นำ URL นี้ไปใส่เป็น Endpoint URL ของ LIFF ในขั้นที่ 2.4

**ตัวเลือกอื่น**: Firebase Hosting, Netlify, Vercel — หลักการเดียวกันคือ host ไฟล์ static แล้วได้ HTTPS URL

---

## ขั้นที่ 4 — แก้ URL ของ backend ในหน้าเว็บ

เปิดไฟล์ `liff-app/js/api.js` แก้ 2 ค่านี้ให้ตรงกับของจริง:

```js
const LIFF_ID = "ใส่ LIFF ID จากขั้นที่ 2.4";
const API_BASE_URL = "ใส่ Web app URL จากขั้นที่ 1.7";
```

แล้วอัปโหลดไฟล์ที่แก้แล้วขึ้น hosting อีกครั้ง (ขั้นที่ 3)

---

## ขั้นที่ 5 — เพิ่มชนิดพืชแรกและตั้งแจ้งเตือนรายวัน

1. เปิด Line OA ของคุณ → กดเข้าหน้า "ตั้งค่า" → เพิ่มชนิดพืชอย่างน้อย 1 ชนิด พร้อมจำนวนวัน/สูตรอาหารแต่ละระยะ
2. กลับไปที่ Apps Script editor → เปิดแท็บ **Triggers** (รูปนาฬิกาทางซ้าย) → **Add Trigger**:
   - Function: `sendDueReminders`
   - Event source: Time-driven
   - Type: Day timer
   - เวลา: เลือกช่วงเช้าที่ต้องการ (เช่น 7:00–8:00)
3. เสร็จแล้ว ทุกคนที่เคยเปิดแอปผ่าน Line อย่างน้อย 1 ครั้ง จะได้รับแจ้งเตือนอัตโนมัติ

---

## หมายเหตุสำคัญเรื่องการออกแบบ (โปรดอ่าน)

ระบบระบุว่า "นี่คือขวดเดิมที่ทำต่อ" หรือ "ขวดใหม่" โดยดูจาก**ชนิดพืช + ชั้นเก็บ + ช่องเก็บเดียวกัน**:
- ถ้าตำแหน่งนี้เคยมีบันทึกที่ยังไม่ถึงระยะ "พร้อมออกปลูก" → รอบนี้นับเป็นรอบถัดไปของขวดเดิม (เลื่อนระยะอัตโนมัติ)
- ถ้าตำแหน่งนี้ว่าง หรือขวดก่อนหน้าถึงระยะสุดท้ายไปแล้ว → เริ่มนับเป็นขวดใหม่ รอบที่ 1

**ถ้าในการทำงานจริง มีหลายขวดวางตำแหน่งเดียวกันพร้อมกัน หรือย้ายขวดไปตำแหน่งอื่นระหว่างรอบ** ตรรกะนี้จะนับผิดได้ — ถ้าเจอกรณีนี้บ่อย แจ้งได้เลยครับ จะปรับให้มีรหัสขวด (bottle ID) แยกต่างหาก แทนการอิงตำแหน่งเก็บ

## แก้ไขเพิ่มเติมในอนาคต
ทุกครั้งที่จะแก้ไขหน้าตา/ข้อความ/กฎการทำงาน ให้ดู `brand.md` (น้ำเสียง), `CLAUDE.md` (กฎเหล็ก), `PRD.md` (สเปค) ก่อนแก้โค้ดเสมอ
