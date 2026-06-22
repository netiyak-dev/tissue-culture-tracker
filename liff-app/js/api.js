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

const LABS = ["KA", "NJ", "CC"];

const TissueApp = {
  profile: null,
  myLab: "",
  myRole: "user",

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
            // ลงทะเบียนผู้ใช้ไว้รับแจ้งเตือน + เช็คว่าระบุ Lab ไว้แล้วหรือยัง
            try {
              const result = await this.call("registerLineUser", {
                line_user_id: this.profile.userId,
                display_name: this.profile.displayName,
              });
              this.myRole = result.role || "user";
              this.myLab = result.lab || "";
              if (!this.myLab) {
                this.myLab = await this._askMyLab();
              }
            } catch (err) {
              console.warn("ลงทะเบียนผู้ใช้ไม่สำเร็จ:", err);
            }
          }
        }
      }
    } catch (err) {
      console.warn("LIFF init ไม่สำเร็จ (อาจเป็นเพราะเปิดทดสอบนอก Line):", err);
    }
  },

  isAdmin() {
    return this.myRole === "admin";
  },

  /** หน้าจอบังคับเลือก Lab ตอนเปิดแอปครั้งแรก (ตั้งได้ครั้งเดียว เปลี่ยนทีหลังต้องให้แอดมินตั้งให้) */
  _askMyLab() {
    return new Promise(resolve => {
      const overlay = document.createElement("div");
      overlay.style.cssText = "position:fixed;inset:0;background:var(--bg);z-index:9999;display:flex;align-items:center;justify-content:center;padding:24px;";
      overlay.innerHTML = `
        <div style="width:100%;max-width:380px;">
          <h1 style="margin-bottom:8px;">เลือก Lab ของคุณ</h1>
          <p class="muted" style="margin-bottom:20px;">เลือกได้ครั้งเดียว ถ้าต้องการเปลี่ยนทีหลังให้แจ้งแอดมิน</p>
          <div class="chip-grid" id="lab-pick-grid" style="grid-template-columns:repeat(3,1fr);"></div>
        </div>
      `;
      const grid = overlay.querySelector("#lab-pick-grid");
      LABS.forEach(lab => {
        const chip = document.createElement("div");
        chip.className = "chip chip-lg";
        chip.textContent = lab;
        chip.addEventListener("click", async () => {
          chip.textContent = "กำลังบันทึก...";
          try {
            const result = await this.call("setMyLab", { line_user_id: this.profile.userId, lab });
            overlay.remove();
            resolve(result.lab);
          } catch (err) {
            alert("เลือก Lab ไม่สำเร็จ: " + err.message);
            chip.textContent = lab;
          }
        });
        grid.appendChild(chip);
      });
      document.body.appendChild(overlay);
    });
  },

  /** ปิดหน้าต่าง LIFF (ใช้ตอนกด "เสร็จแล้ว" ถ้าต้องการ) */
  closeWindow() {
    if (window.liff && liff.isInClient()) {
      liff.closeWindow();
    } else {
      window.history.back();
    }
  },

  /** เรียก backend (Google Apps Script) ด้วยชื่อ action + ข้อมูล — แนบ line_user_id ของผู้เรียกให้อัตโนมัติทุกครั้ง (ใช้เช็คสิทธิ์/Lab ฝั่ง backend) */
  async call(action, payload = {}) {
    const fullPayload = { ...payload };
    if (this.profile && this.profile.userId && fullPayload.line_user_id === undefined) {
      fullPayload.line_user_id = this.profile.userId;
    }
    const res = await fetch(API_BASE_URL, {
      method: "POST",
      headers: { "Content-Type": "text/plain;charset=utf-8" }, // กัน CORS preflight ของ Apps Script
      body: JSON.stringify({ action, payload: fullPayload }),
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
