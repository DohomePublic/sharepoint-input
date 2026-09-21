# DOHOME KYC - SharePoint Graph API

แพ็กเกจเว็บฟอร์ม KYC สำหรับส่งข้อมูลเข้า SharePoint List `DemoApp`

## ไฟล์ในแพ็กเกจ
- `index.html` เว็บฟอร์มพร้อม Microsoft Login และ Microsoft Graph API
- `README.md` คู่มือติดตั้ง
- `serve-local.bat` เปิดเว็บทดสอบบน Windows ด้วย Python
- `serve-local.sh` เปิดเว็บทดสอบบน macOS/Linux ด้วย Python

## SharePoint ที่ตั้งค่าไว้
- Host: `dohomegroup.sharepoint.com`
- Site: `/sites/AC-Accounting`
- List: `DemoApp`

## Microsoft Entra ID ที่ต้องตั้งค่า
1. เปิด App registration ของ Client ID ที่ระบุใน `index.html`
2. Authentication > Add platform > Single-page application
3. เพิ่ม Redirect URI ให้ตรงกับ URL ที่เปิดเว็บจริง เช่น GitHub Pages หรือ `http://localhost:5500/index.html`
4. API permissions > Microsoft Graph > Delegated permissions
   - `User.Read`
   - `Sites.ReadWrite.All`
5. ให้ผู้ดูแลระบบ Grant admin consent

สำคัญ: ห้ามใส่ Client Secret ใน `index.html` เพราะเป็นเว็บ SPA และ Source Code เปิดดูได้

## ทดสอบในเครื่อง
### Windows
ดับเบิลคลิก `serve-local.bat` แล้วเปิด `http://localhost:5500/index.html`

### macOS/Linux
รัน:
```bash
chmod +x serve-local.sh
./serve-local.sh
```
จากนั้นเปิด `http://localhost:5500/index.html`

ไม่ควรเปิดไฟล์ด้วย `file://` เพราะ Microsoft Login ต้องใช้ Redirect URI ผ่าน HTTP/HTTPS

## ใช้งาน
1. กด “เข้าสู่ระบบ Microsoft”
2. เข้าด้วยบัญชีองค์กรที่มีสิทธิ์ SharePoint
3. กรอกข้อมูล KYC
4. กด “ส่งเข้า SharePoint”
5. ระบบจะแสดงเลขที่รายการเมื่อบันทึกสำเร็จ

## หากบันทึกไม่สำเร็จ
- `redirect_uri` ไม่ตรง: เพิ่ม URL ที่ใช้งานจริงใน SPA Redirect URI
- `Consent required` หรือ `403`: ตรวจ `Sites.ReadWrite.All` และ Grant admin consent
- `Field ... is not recognized`: ตรวจ Internal Name ของคอลัมน์ SharePoint โดยเฉพาะคอลัมน์ที่มีช่องว่าง
- `ไม่พบ SharePoint List DemoApp`: ตรวจสิทธิ์บัญชีและชื่อ List
