/* ===================================================================
   liff-app/js/api.js
   - เชื่อมต่อ LIFF SDK (ทำงานเมื่อเปิดผ่าน Line)
   - เรียก API ไปยัง Google Apps Script Web App

   *** ต้องแก้ 2 ค่านี้ก่อนใช้งานจริง ***
   1. LIFF_ID            -> ได้จาก LINE Developers Console > LIFF
   2. API_BASE_URL       -> URL ของ Apps Script Web App ที่ deploy แล้ว
   =================================================================== */

const LIFF_ID = "2010453502-yUgNYcUn";
const API_BASE_URL = "https://script.google.com/macros/s/AKfycbyvQDmkOIrFK_BZ1WXExuxrqTEyxGiSdCsVFqraNPS7TeyPzu0we35UjSKvwRiekcbvPw/exec";

const TissueApp = {
  profile: null,

  /** เรียกตอนเริ่มทุกหน้า: เช็คว่าเปิดผ่าน Line หรือเปิดผ่านเบราว์เซอร์ทั่วไป (เผื่อทดสอบ) */
  async init() {
    try {
      if (window.liff) {
        await liff.init({ liffId: LIFF_ID });
        if (!liff.isLoggedIn() && !liff.isInClient()) {
          liff.login();
          return;
        }
        if (liff.isLoggedIn() || liff.isInClient()) {
          this.profile = await liff.getProfile().catch(() => null);
          if (this.profile) {
            // ลงทะเบียนผู้ใช้ไว้รับแจ้งเตือน (ไม่บล็อกการใช้งาน ถ้า fail ก็ไม่เป็นไร)
            this.call("registerLineUser", {
              line_user_id: this.profile.userId,
              display_name: this.profile.displayName,
            }).catch(() => {});
          }
        }
      }
    } catch (err) {
      console.warn("LIFF init ไม่สำเร็จ (อาจเป็นเพราะเปิดทดสอบนอก Line):", err);
    }
  },

  /** ปิดหน้าต่าง LIFF (ใช้ตอนกด "เสร็จแล้ว" ถ้าต้องการ) */
  closeWindow() {
    if (window.liff && liff.isInClient()) {
      liff.closeWindow();
    } else {
      window.history.back();
    }
  },

  /** เรียก backend (Google Apps Script) ด้วยชื่อ action + ข้อมูล */
  async call(action, payload = {}) {
    const res = await fetch(API_BASE_URL, {
      method: "POST",
      headers: { "Content-Type": "text/plain;charset=utf-8" }, // กัน CORS preflight ของ Apps Script
      body: JSON.stringify({ action, payload }),
    });
    if (!res.ok) throw new Error("เรียกข้อมูลไม่สำเร็จ (" + res.status + ")");
    const data = await res.json();
    if (data.ok === false) throw new Error(data.message || "เกิดข้อผิดพลาด");
    return data.result;
  },

  /** ดึงค่าจาก query string เช่น ?plant=xxx */
  getParam(name) {
    return new URLSearchParams(window.location.search).get(name);
  },

  /** แสดง loading แบบง่ายๆ บนปุ่ม */
  setLoading(btn, isLoading, loadingText = "กำลังบันทึก...") {
    if (!btn) return;
    if (isLoading) {
      btn.dataset.originalText = btn.textContent;
      btn.textContent = loadingText;
      btn.disabled = true;
    } else {
      btn.textContent = btn.dataset.originalText || btn.textContent;
      btn.disabled = false;
    }
  },
};
