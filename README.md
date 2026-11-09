# LabAlert

ระบบติดตามผล Lab ผู้ป่วยอายุรกรรม
- **หน้าเว็บ** (`index.html`) โฮสต์บน GitHub Pages ลิงก์เดียว ใช้กล้องสแกนบาร์โค้ดได้ในหน้าเดียวกัน
- **ฐานข้อมูลและตัวแจ้งเตือน** (`Code.gs`) ทำงานบน Google Apps Script + Google Sheets และส่งข้อความเข้ากลุ่ม LINE

## ความสามารถ
- สแกน Barcode/QR ของ HN แล้วเลือกเวลาแจ้งเตือน 45 นาที / 1 ชม. / 1.5 ชม. / 2 ชม. / กำหนดเอง
- แสดงเฉพาะ HN ไม่แสดงชื่อ แยก 3 แท็บ: **ครบเวลา** (แดง สถานะกระพริบ) / **รอผล** (เหลือง) / **เสร็จสิ้น** (พิมพ์ผลแล้ว)
- Time stamp ลงทะเบียน ตามผล (หลายครั้ง) พิมพ์ผล ซ่อนเป็น "ประวัติเวลา" รายเคส หน้าจอกว้าง (Admin) เห็นเวลารอตอบรับ
- แจ้งเตือนเข้ากลุ่ม LINE เมื่อครบเวลา 1 เคสต่อ 1 กล่องข้อความ แจ้งซ้ำทุก 10 นาที เมื่อกดตามผลจะนับ 10 นาทีใหม่
- เข้าผ่าน LINE: ระบุตัวตนด้วยบัญชี LINE / เข้าผ่านเว็บ: รหัสแผนก (ตัวเลข 4-8 หลัก)

## ติดตั้ง

### 1. ฝั่ง Apps Script
1. สร้าง Google Sheet ใหม่ → Extensions → Apps Script
2. วาง `Code.gs` และ `appsscript.json` (เปิด "Show appsscript.json" ที่ Project Settings) ไม่ต้องมีไฟล์ `Index` แล้ว ถ้ามีอยู่เดิมให้ลบทิ้ง
3. Project Settings → Script Properties ใส่ `DEPT_PIN` (รหัสแผนก) และภายหลัง `LINE_TOKEN`, `GROUP_ID`, `LIFF_ID`, `LOGIN_CHANNEL_ID` (ห้ามใส่ในโค้ดหรือ GitHub)
4. รันฟังก์ชัน `setup` หนึ่งครั้ง
5. Deploy → New deployment → Web app (Execute as: Me, Who has access: Anyone) แล้วคัดลอก Web app URL (ลงท้าย `/exec`) ถ้าแก้โค้ดทีหลังต้อง Deploy → Manage deployments → New version

### 2. ฝั่งหน้าเว็บ (GitHub Pages)
1. Repo Settings → Pages → Deploy from a branch → `main` / `/ (root)` → Save
2. เปิด `https://<ชื่อบัญชี>.github.io/LabAlert/` ครั้งแรกจะให้วาง Web app URL ของ Apps Script (จำไว้ในเครื่อง) แล้วใส่รหัสแผนก
3. กดสแกน → อนุญาตกล้อง

### ติดตั้งเป็นแอปบนหน้าจอหลัก (PWA)
- **Android (Chrome):** เปิดลิงก์ แล้วกดปุ่ม "ติดตั้ง" ที่ขึ้นในแอป (หรือเมนูสามจุด → ติดตั้งแอป / เพิ่มไปยังหน้าจอหลัก)
- **iPhone (Safari):** แตะปุ่มแชร์ → เพิ่มไปยังหน้าจอโฮม
- **PC (Chrome/Edge):** กดไอคอนติดตั้งที่ช่องที่อยู่ หรือเมนู → ติดตั้ง LabAlert
- ไอคอนและชื่อมาจาก `manifest.webmanifest` และไฟล์ `icon-*.png` ตัวแอปโหลดหน้าล่าสุดจากเน็ตก่อนเสมอ (`sw.js`) ไม่เก็บข้อมูลผู้ป่วยในแคช ใช้งานโดยไม่มีเน็ตไม่ได้

### 3. LINE (ภายหลัง)
- สร้าง LINE Login channel แล้วสร้าง LIFF โดยตั้ง Endpoint URL เป็นลิงก์ GitHub Pages ข้างบน เปิด scope `profile`
- ใส่ `LIFF_ID`, `LOGIN_CHANNEL_ID`, `LINE_TOKEN`, `GROUP_ID` ใน Script Properties แล้วรัน `testPush`, `postMenu`
- ลิงก์ในข้อความ LINE แนบลิงก์ Apps Script มาด้วย หน้าเว็บจะถามยืนยันก่อนเชื่อมต่อครั้งแรก

## หมายเหตุ
- หน้าเว็บบน GitHub Pages เป็นสาธารณะแต่ไม่มีข้อมูลผู้ป่วยและไม่ฝังลิงก์ Apps Script ข้อมูลอยู่ใน Google Sheet และทุกคำสั่งต้องผ่านรหัสแผนกหรือบัญชี LINE
- ใส่รหัสผิดเกิน 5 ครั้งจะถูกล็อก 10 นาที
- 1 เคสต่อ 1 กล่องใช้โควตาข้อความ LINE มากกว่าการรวมกล่อง ควรเช็คโควตาแพ็กเกจก่อนใช้จริง
- LINE ไม่มี API แก้ไขหรือลบข้อความที่บอทส่งแล้ว
- การพิมพ์ผลยังกดได้ทันที เมื่อมีรายการ Lab ให้ตรวจว่าผลออกครบก่อน (ดู TODO ใน `apiPrint`)
