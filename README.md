# DOHOME SharePoint KYC Input — ตัวอย่างระบบพร้อมทดสอบ

ระบบนี้มี 2 ส่วนที่แยกกันโดยสิ้นเชิง อย่าสับสนสิทธิ์ของสองส่วนนี้:

| ส่วน | ใช้ทำอะไร | วิธียืนยันตัวตน | ต้องใช้ Client Secret? |
|---|---|---|---|
| **`index.html` (SPA บน GitHub Pages)** | ฟอร์มกรอก KYC จริง ที่พนักงานใช้บันทึกข้อมูลลง SharePoint List | ผู้ใช้ login ด้วยบัญชีตัวเอง (delegated, PKCE) | **ไม่ต้อง** — SPA ไม่มี secret |
| **`scripts/test-connection.js` (GitHub Actions)** | สคริปต์ตรวจสุขภาพ ยืนยันว่า Secrets/สิทธิ์ตั้งไว้ถูกต้อง โดยไม่ต้องมีคนกด login | app-only (client credentials) | **ต้องใช้** — ตามที่เก็บใน GitHub Secrets |

## 1) ติดตั้งไฟล์ลง repository

1. อัปโหลดไฟล์ทั้งหมดในนี้ไปที่ root ของ `DohomePublic/sharepoint-input` แล้ว commit ไปที่ `main`
   (โครงสร้าง: `index.html`, `config.js`, `README.md`, `.github/workflows/test-connection.yml`, `scripts/test-connection.js`)
2. GitHub → Settings → Pages → Deploy from a branch → `main` / `(root)`
3. รอจน Pages ขึ้นสถานะ published แล้วเปิด `https://dohomepublic.github.io/sharepoint-input/`

## 2) ตั้งค่า Microsoft Entra ID App registration

App registration ตัวนี้ต้องมีสิทธิ์ **สองชุด** เพราะมีสองการยืนยันตัวตนคนละแบบ:

### ก) สำหรับ SPA (delegated) — ใช้ตอนพนักงาน login
- Authentication → Platform: **Single-page application**
- Redirect URI: `https://dohomepublic.github.io/sharepoint-input/`
- API permissions → Delegated permissions: `User.Read`, `Sites.ReadWrite.All` → Grant admin consent

### ข) สำหรับ GitHub Actions (application/app-only) — ใช้ตอนรัน test script
- API permissions → **Application permissions**: `Sites.ReadWrite.All` → Grant admin consent
  (นี่คนละ checkbox กับ Delegated — ต้องเพิ่มเพิ่มต่างหาก มิฉะนั้น `test-connection.js` จะได้ `403`)
- Certificates & secrets → New client secret → คัดลอกค่า **Value** (ไม่ใช่ Secret ID) ไปใส่ใน GitHub Secret `AZ_CLIENT_SECRET` ทันที เพราะจะดูค่าซ้ำไม่ได้อีก

## 3) ตั้งค่า GitHub Secrets และ Variables

Settings → Secrets and variables → Actions

**Secrets** (encrypted, ห้าม commit ลงโค้ด):
| ชื่อ | ค่า |
|---|---|
| `AZ_CLIENT_ID` | `a37bd62d-e74d-4ea0-9546-1eb5aa96f604` |
| `AZ_TENANT_ID` | `7f8918d9-718a-495b-ac9a-17cba381c4a0` |
| `AZ_CLIENT_SECRET` | ค่า Client Secret จากขั้นตอน 2ข |

**Variables** (ไม่บังคับ — มี default ในสคริปต์อยู่แล้ว ใส่เฉพาะถ้าต่างจาก default):
| ชื่อ | Default |
|---|---|
| `SP_HOSTNAME` | `dohomegroup.sharepoint.com` |
| `SP_SITE_PATH` | `/sites/AC-Accounting` |
| `SP_LIST_NAME` | `DemoApp ต้องการบึนทึกเข้า Sharepoint` |

## 4) ทดสอบจริง — ทีละส่วน

### ทดสอบส่วนที่ 1: Secrets/สิทธิ์ถูกต้องหรือไม่ (ไม่ต้องเปิดหน้าเว็บ)
1. ไปที่ tab **Actions** ของ repo → เลือก workflow **"Test SharePoint connection"**
2. กด **Run workflow** (workflow_dispatch)
3. เปิด log แล้วดูผลลัพธ์:
   - ✓ ควรเห็นว่าอ่านค่าจากตัวแปรชื่อ `AZ_CLIENT_ID` / `AZ_TENANT_ID` / `AZ_CLIENT_SECRET` (ไม่มีการ echo ค่าจริง)
   - ✓ ได้ access token
   - ✓ พบ site และ list
   - ถ้า **403** ที่ขั้นอ่าน site → แปลว่ายังไม่ได้ grant admin consent สิทธิ์ **Application** `Sites.ReadWrite.All` (คนละอันกับ Delegated ในข้อ 2ก)
   - ถ้า **list ไม่พบ** → ชื่อใน `SP_LIST_NAME` ไม่ตรงกับชื่อ List จริงใน SharePoint เป๊ะๆ (รวมช่องว่าง)

### ทดสอบส่วนที่ 2: ฟอร์มจริงบน GitHub Pages
1. เปิด `https://dohomepublic.github.io/sharepoint-input/`
2. กด "เข้าสู่ระบบด้วย Microsoft" → login ด้วยบัญชีที่มีสิทธิ์เขียนลง SharePoint site นั้น
   - ถ้า login แล้วเด้ง error `AADSTS50011` (redirect URI mismatch) → Redirect URI ในขั้นตอน 2ก พิมพ์ไม่ตรงกับ URL จริง (ต้องมี `/` ปิดท้ายให้ตรงกัน)
3. กรอกฟอร์ม แล้วกด "บันทึกเข้า SharePoint"
4. ควรเห็นข้อความ "บันทึกสำเร็จ · Item ID ..." แล้วเข้าไปเช็คใน SharePoint List ว่ามีแถวใหม่จริง

## 5) ก่อนใช้งานจริง — สิ่งที่ต้องแก้ให้ตรงกับของจริง

- **`config.js` → `fieldMapping`**: ตอนนี้ map เป็นชื่อ column สมมติ (`CustomerType`, `IDNumber`, ...) ต้องแก้ให้ตรงกับ **internal name** จริงของคอลัมน์ใน List (ดูวิธีหาใน comment ของไฟล์) ไม่ใช่ชื่อที่เห็นบนหน้าจอ ถ้า map ผิดชื่อ Graph จะตอบ `400 Bad Request` ตอน POST
- **`config.js` → `spListName`**: ต้องตรงกับชื่อ List จริงเป๊ะ
- ฟิลด์ในฟอร์ม (ชื่อ, เลขบัตร, ที่อยู่ ฯลฯ) เป็นตัวอย่างเบื้องต้นสำหรับ KYC — เพิ่ม/ลด/ปรับ validation ได้ตาม requirement จริงของบัญชี AC-Accounting

## โครงสร้างไฟล์
```
/
├── index.html                          ← ฟอร์ม KYC (SPA, MSAL.js, เขียนลง SharePoint ผ่าน Graph)
├── config.js                           ← ค่า config สาธารณะ + field mapping (ไม่มี secret)
├── scripts/test-connection.js          ← สคริปต์ทดสอบ app-only auth + อ่าน site/list
└── .github/workflows/test-connection.yml  ← workflow รันสคริปต์ข้างบน
```
