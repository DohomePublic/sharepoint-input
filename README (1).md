# ส่งข้อมูลเข้า SharePoint List `DemoApp` ผ่าน GitHub Actions

ส่ง (insert/upsert) ข้อมูลจากไฟล์ CSV หรือ JSON เข้า SharePoint Online list
`DemoApp` บนไซต์ `https://dohomegroup.sharepoint.com/sites/AC-Accounting`
โดยใช้ Microsoft Graph + Azure AD app registration (client credentials flow)

---

มี 2 ช่องทางการส่งข้อมูล:
1. **หน้าเว็บกรอกข้อมูล** (`webapp/server.py`) — กรอกทีละรายการ + ทดสอบการเชื่อมต่อ
2. **ส่งเป็นชุดจากไฟล์** (`scripts/sp_upload.py` + GitHub Actions) — CSV/JSON หลายรายการ

---

## 1. โครงสร้างไฟล์

```
.github/workflows/upload-to-sharepoint.yml   GitHub Actions workflow
webapp/index.html                             หน้าเว็บเพิ่มข้อมูล (HTML/CSS/JS)
webapp/server.py                              web server + JSON API (stdlib ล้วน ไม่ต้องใช้ Flask)
scripts/sp_upload.py                          สคริปต์ส่งข้อมูลเป็นชุด (Python 3.9+)
config/field_map.json                         map ชื่อคอลัมน์ไฟล์ -> internal name ของ list
data/upload.csv                               ไฟล์ข้อมูลตัวอย่าง (ใช้คอลัมน์จริง)
docs/form-preview.html                        สำเนาหน้าฟอร์มไว้ดูหน้าตาออฟไลน์
run_web.sh / run_web.bat                      สคริปต์เปิดหน้าเว็บ (Linux-macOS / Windows)
.env.example                                  ตัวอย่างไฟล์ตั้งค่า (คัดลอกเป็น .env)
requirements.txt                              requests
.gitignore                                    กัน .env / secret หลุดเข้า repo
```

---

## 1.1 หน้าเว็บเพิ่มข้อมูล (ทดสอบส่งเข้า SharePoint)

```bash
cp .env.example .env          # ใส่ AZ_CLIENT_SECRET ลงไป
./run_web.sh                  # Windows: run_web.bat
# หรือทดสอบหน้าตาก่อนโดยไม่เขียนจริง
./run_web.sh --dry-run
```

เปิดเบราว์เซอร์ที่ **http://127.0.0.1:8000**

> ⚠️ **ต้องเปิดผ่าน server เท่านั้น** — อย่า double-click ไฟล์ `webapp/index.html`
> เพราะหน้าเว็บต้องเรียก API ที่ `server.py` ให้บริการ ถ้าเปิดแบบ `file://`
> ฟอร์มจะแสดงได้แต่กดบันทึกไม่ได้ และจะมีข้อความเตือนสีแดงบอกให้เปิดผ่าน server

### ถ้าไม่มี `run_web.sh` หรือรันไม่ได้ ให้สั่งตรง ๆ

```bash
pip install -r requirements.txt
python webapp/server.py --port 8000
```

เมื่อ server ทำงานจะขึ้นข้อความ:

```
[web] เปิดเบราว์เซอร์ที่  http://127.0.0.1:8000   (Ctrl+C เพื่อหยุด)
```

| ปุ่ม | การทำงาน |
|---|---|
| **บันทึกเข้า SharePoint (Upsert)** | ถ้ามี `Customer_id` นี้อยู่แล้ว → อัปเดต, ไม่มี → สร้างใหม่ |
| **สร้างรายการใหม่เสมอ** | สร้าง item ใหม่ทุกครั้ง (insert) |
| **ตรวจสอบอย่างเดียว** | แสดง JSON ที่จะส่งพร้อม internal name โดยไม่เขียน |

ฟีเจอร์ของหน้าเว็บ:
- ฟิลด์ตามคอลัมน์จริงของ `DemoApp` พร้อม **dropdown ค่าที่ถูกต้อง** (Status, Type_Request, Type1, type_teams, business_type, Estimated_annual_income) ลดการพิมพ์ผิด
- ตรวจฟิลด์บังคับ: `Title`, `Registered_Name`, `Customer_id`, `Status`
- แปลงชื่อคอลัมน์เป็น internal name อัตโนมัติ (เช่น `Customer Name` → `Customer_x0020_Name`)
- ตัดลูกน้ำในจำนวนเงินให้ (`5,000,000` → `5000000`)
- เมื่อบันทึกสำเร็จจะแสดง **item ID** พร้อมลิงก์เปิดรายการใน SharePoint และล้างฟอร์มให้
- ตารางประวัติการส่งของ session

### Endpoints ของ server

| Method | Path | หน้าที่ |
|---|---|---|
| GET | `/` | หน้าเว็บ (`webapp/index.html`) |
| GET | `/api/fields` | นิยามฟิลด์บนฟอร์ม — แก้ที่ `FORM_FIELDS` ใน `server.py` ที่เดียว |
| GET | `/api/health` | สถานะการเชื่อมต่อ, site id, list id |
| GET | `/api/schema` | คอลัมน์ที่เขียนได้จริง (internal name) |
| POST | `/api/submit` | บันทึกข้อมูล — body `{"mode":"upsert\|insert\|validate","fields":{...}}` |

ทดสอบด้วย curl ได้:

```bash
curl http://127.0.0.1:8000/api/health
curl -X POST http://127.0.0.1:8000/api/submit -H "Content-Type: application/json" \
  -d '{"mode":"validate","fields":{"Title":"T","Registered_Name":"R","Customer_id":"101047587","Status":"Draft","limit":"5,000,000"}}'
```

หน้าเว็บนี้ bind กับ `127.0.0.1` เท่านั้น (ไม่เปิดสู่ภายนอก) เปลี่ยนได้ด้วย
`python webapp/server.py --host 0.0.0.0 --port 8000` — ถ้าเปิดสู่เครือข่ายควรวางหลัง
reverse proxy ที่มีการยืนยันตัวตน เพราะหน้านี้ยังไม่มีระบบ login

---

## 2. ตั้งค่า GitHub Secrets

ไปที่ **Settings → Secrets and variables → Actions → Secrets → New repository secret**

| Secret | Value |
|---|---|
| `AZ_CLIENT_ID` | `a37bd62d-e74d-4ea0-9546-1eb5aa96f604` |
| `AZ_TENANT_ID` | `7f8918d9-718a-495b-ac9a-17cba381c4a0` |
| `AZ_CLIENT_SECRET` | ค่า Client Secret จาก Azure AD — **ห้าม commit ลงโค้ดเด็ดขาด** |

สคริปต์รองรับชื่อ secret ทั้งสองแบบ และเลือกตัวที่ "มีค่า" ให้อัตโนมัติ:

| ชื่อที่ใช้จริง | ชื่อเดิม (legacy) |
|---|---|
| `AZ_CLIENT_ID` | `AZURE_CLIENT_ID` |
| `AZ_TENANT_ID` | `AZURE_TENANT_ID` |
| `AZ_CLIENT_SECRET` | `AZURE_CLIENT_SECRET` |

log จะพิมพ์ **ชื่อ** ตัวแปรที่อ่านค่ามาได้ เช่น `client id <- AZ_CLIENT_ID`
และสำหรับ secret จะแสดงเพียง `<hidden, N chars>` — **ไม่แสดงค่าจริง**

### ตัวแปรเสริม (Variables — ไม่บังคับ)

**Settings → Secrets and variables → Actions → Variables**

| Variable | Default ในสคริปต์ |
|---|---|
| `SP_HOSTNAME` | `dohomegroup.sharepoint.com` |
| `SP_SITE_PATH` | `/sites/AC-Accounting` |
| `SP_LIST_NAME` | `DemoApp` |

### ข้อมูล App registration ที่ใช้

| รายการ | ค่า |
|---|---|
| Application (client) ID | `a37bd62d-e74d-4ea0-9546-1eb5aa96f604` |
| Object ID | `f4e84724-e3f8-444b-981b-74ead3130171` |
| Directory (tenant) ID | `7f8918d9-718a-495b-ac9a-17cba381c4a0` |

Client ID / Tenant ID / Object ID **ไม่ใช่ความลับ** (เป็นตัวระบุแอป)
แต่ **Client Secret เป็นความลับ** ต้องเก็บใน GitHub Secrets เท่านั้น

---

## 3. สิทธิ์ที่ App registration ต้องมี (สำคัญ)

ใน Azure Portal → App registrations → **API permissions** → Microsoft Graph →
**Application permissions** เลือกอย่างใดอย่างหนึ่ง แล้วกด **Grant admin consent**:

| ทางเลือก | Permission | หมายเหตุ |
|---|---|---|
| แนะนำ (least privilege) | `Sites.Selected` | ต้องให้ admin grant สิทธิ์ `write` เฉพาะไซต์ AC-Accounting เพิ่มด้วย |
| ง่ายกว่าแต่กว้าง | `Sites.ReadWrite.All` | เขียนได้ทุกไซต์ใน tenant |

กรณีใช้ `Sites.Selected` ให้ admin รันคำสั่งนี้ครั้งเดียวเพื่อผูกสิทธิ์กับไซต์:

```http
POST https://graph.microsoft.com/v1.0/sites/{site-id}/permissions
{
  "roles": ["write"],
  "grantedToIdentities": [{
    "application": { "id": "a37bd62d-e74d-4ea0-9546-1eb5aa96f604",
                     "displayName": "DemoApp Uploader" }
  }]
}
```

หา `{site-id}` ได้จาก
`GET https://graph.microsoft.com/v1.0/sites/dohomegroup.sharepoint.com:/sites/AC-Accounting`

> ถ้าได้ error `403 accessDenied` ตอนรัน แปลว่าขั้นตอนนี้ยังไม่ครบ

---

## 4. วิธีรัน

### รันจากหน้า GitHub (Actions → Upload to SharePoint (DemoApp) → Run workflow)

| Input | ความหมาย |
|---|---|
| `input_file` | path ไฟล์ CSV/JSON ใน repo (default `data/upload.csv`) |
| `mode` | `upsert` = มีอยู่แล้วให้อัปเดต, `insert` = สร้างใหม่เสมอ |
| `key_column` | คอลัมน์ที่ใช้จับคู่รายการเดิม (default `Customer_id`) |
| `dry_run` | `true` = ตรวจสอบอย่างเดียว ไม่เขียนจริง (ค่าเริ่มต้น) |

workflow ยังทำงานอัตโนมัติเมื่อ push ไฟล์ใน `data/**` เข้า `main`
และตามตาราง cron `0 1 * * *` (08:00 เวลาไทย)

### รันบนเครื่องตัวเอง

```bash
pip install -r requirements.txt

export AZ_CLIENT_ID=a37bd62d-e74d-4ea0-9546-1eb5aa96f604
export AZ_TENANT_ID=7f8918d9-718a-495b-ac9a-17cba381c4a0
export AZ_CLIENT_SECRET='<client secret>'      # อย่าใส่ลงไฟล์ที่ commit

# ดูคอลัมน์ที่เขียนได้จริงใน list (internal name)
python scripts/sp_upload.py --show-schema

# ตรวจสอบก่อนส่งจริง
python scripts/sp_upload.py --input data/upload.csv --key-column Customer_id --dry-run

# ส่งจริงแบบ upsert
python scripts/sp_upload.py --input data/upload.csv --key-column Customer_id --mode upsert
```

### options ทั้งหมด

| Option | ความหมาย |
|---|---|
| `--input` | ไฟล์ CSV หรือ JSON (`[...]` หรือ `{"items":[...]}`) |
| `--field-map` | ไฟล์ map header → internal field (default `config/field_map.json`) |
| `--key-column` | คอลัมน์ key สำหรับ upsert |
| `--mode` | `upsert` (default) / `insert` |
| `--limit N` | ส่งเฉพาะ N แถวแรก (ใช้ทดสอบ) |
| `--keep-raw-numbers` | ไม่ตัดลูกน้ำในจำนวนเงิน (ปกติ `5,000,000` → `5000000`) |
| `--batch-sleep` | หน่วงเวลาระหว่างแถว (วินาที) ลดโอกาสโดน throttle |
| `--dry-run` | ไม่เขียนจริง — ถ้าไม่มี secret จะตรวจแค่ไฟล์ input |
| `--show-schema` | แสดงคอลัมน์ที่เขียนได้ของ list แล้วจบ |

---

## 5. การ map คอลัมน์

SharePoint ใช้ **internal name** ซึ่งต่างจากชื่อที่แสดง เช่น
`Customer Name` → `Customer_x0020_Name`, `Request TimeStamp` → `Request_x0020_TimeStamp`

สคริปต์จัดการให้ 3 ชั้นตามลำดับ:
1. ใช้ค่าที่ระบุใน `config/field_map.json`
2. จับคู่กับ internal name ตรงตัว
3. จับคู่กับ display name (ไม่สนตัวพิมพ์เล็ก/ใหญ่)

คอลัมน์ที่หาไม่เจอหรือเป็น read-only (`Created`, `Modified`, `Author`, `Editor`, `ID`, …)
จะถูกข้ามและสรุปไว้ท้าย log ว่า *columns ignored*

ให้รัน `--show-schema` ก่อนเสมอ แล้วแก้ `config/field_map.json` ให้ตรงกับ list จริง

### คอลัมน์หลักของ `DemoApp` (416 รายการ, 50 คอลัมน์)

`Title`, `Customer_id`, `Type1`, `type_teams`, `Typr_Distribution`, `Typr_Retail`,
`Customer Name`, `branch`, `Request TimeStamp`, `Status`, `Type_Request`, `limit`,
`CraditApprove`, `1addmonney`, `1CreditApprove`, `Owner`, `Data`, `registration_number`,
`building_road`, `county`, `district`, `province`, `post_office`, `telephone`,
`Registered_Name`, `business_type`, `Estimated_annual_income`, `contact_name`, `position`,
`contact_number`, `Wholesale_retail_stores`, `credit_semester1`, `Margin_type1`, `value`,
`limit_other`, `credit_semester2`, `Margin_type2`, `value2`, `limit_OD`, `Bank1`,
`insurance_limit`, `Bank2`, `leasing_limit`, `Bank3`, `Other_limits`, `Bank4`, `land`,
`Status_1`, `other_property` [doc:turn1doc1]

ค่าที่ใช้ได้ของคอลัมน์สำคัญ (อ้างอิงข้อมูลจริงใน list):

| คอลัมน์ | ค่าที่พบ |
|---|---|
| `Status` | ไม่ผ่านการพิจารณาเบื้องต้น, Draft, ไม่อนุมัติ, อนุมัติ-KYC, ผ่านการพิจารณาเบื้องต้น, รอการพิจารณาเบื้องต้น |
| `Type_Request` | คำขอเปิดวงเงินลูกค้าใหม่, C.ติดตามชุดเปิดตัวจริง, คำขอเพิ่มวงเงิน, เพิ่มวงเงิน, เปิดวงเงินลูกค้าใหม่ |
| `Type1` | Existing, Lead, ค้าปลีก |
| `type_teams` | Store Operation, Wholesales (WS), Project Sales (PS), Retail, Steel Key Account, ค้าปลีก, Key Account, ตะวันตก, ผู้แทนขาย |
| `Estimated_annual_income` | 10 - 50 ล้านบาท, มากกว่า 50 ล้านบาท |

> ข้อควรระวัง: ตอนนี้คอลัมน์จำนวนเงินอย่าง `limit` / `CraditApprove` เก็บเป็น **Text**
> และมีทั้งรูปแบบ `5,000,000` และ `500000` ปนกัน ถ้าจะนำไปคำนวณควรทำความสะอาดก่อนส่ง
> หรือแก้ชนิดคอลัมน์ใน SharePoint เป็น Number [doc:turn1doc1]

---

## 6. รูปแบบไฟล์ข้อมูล

CSV (แถวแรกเป็น header ตรงกับชื่อคอลัมน์ / display name) — ตัวอย่างใน `data/upload.csv`
ใช้คอลัมน์จริงตามที่ใช้งาน:

```csv
Title,Registered_Name,Customer_id,Type1,Status,Type_Request,business_type,Customer Name,branch,Owner,province,telephone,limit
Rungroj Chatch Limited Partnership...,Rungroj Chatchawal Limited Partnership,101047587,,Draft,Fee for requesting to open a credit line for new customers,Construction contractor (add),Rungroj Chatchawal,HQ,phongsapan.mar@dohome.co.th,กรุงเทพมหานคร,021234567,1000000
```

JSON:

```json
{ "items": [
  { "Title": "CR-2026-0001", "Customer_id": "C100001", "Status": "Draft" }
] }
```

เซลล์ว่างจะถูกข้าม (ไม่ส่งไปทับค่าที่มีอยู่เดิมตอน update)
ไฟล์ต้องเป็น UTF-8 (รองรับ BOM) เพื่อให้ภาษาไทยไม่เพี้ยน

---

## 7. ความปลอดภัย

- Client Secret อยู่ใน GitHub Secrets เท่านั้น ไม่มีการ commit และไม่พิมพ์ลง log
- log แสดงเฉพาะ **ชื่อ** ตัวแปรที่อ่านมา และความยาวของ secret
- `.gitignore` กัน `.env` / `secrets.json` ไว้แล้ว
- workflow ใช้ `permissions: contents: read` และ `environment: sharepoint`
  (ไปตั้ง required reviewers ได้ที่ Settings → Environments)
- `concurrency` ป้องกันการรันซ้อนจนเกิดข้อมูลซ้ำ
- ตั้งวันหมดอายุของ Client Secret ให้สั้น (6–12 เดือน) และจดวันหมุนเวียนไว้
  หรือย้ายไปใช้ **OIDC federated credential** เพื่อเลิกใช้ secret ถาวร

---

## 8. แก้ปัญหาที่พบบ่อย

| อาการ | สาเหตุ / วิธีแก้ |
|---|---|
| `missing credentials: ...` | ยังไม่ได้ตั้ง secret ใน repo หรือสะกดชื่อผิด |
| `token request failed (401) AADSTS7000215` | Client Secret ผิด/หมดอายุ → สร้างใหม่แล้วอัปเดต secret |
| `cannot resolve site (403)` | ยังไม่ได้ grant `Sites.Selected` / `Sites.ReadWrite.All` + admin consent |
| `list 'DemoApp' not found` | ชื่อ list ไม่ตรง → ตั้ง variable `SP_LIST_NAME` ให้ถูก |
| `row N failed (400) invalidRequest` | ชื่อ field หรือชนิดข้อมูลไม่ตรง → รัน `--show-schema` แล้วแก้ `field_map.json` |
| `columns ignored: ...` | header ในไฟล์ไม่ตรงกับ list หรือเป็นคอลัมน์ read-only |
| เปิดหน้าเว็บแล้วขึ้น "เรียก API ไม่ได้" | เปิดไฟล์ `index.html` แบบ `file://` → ต้องรัน `python webapp/server.py` แล้วเข้า `http://127.0.0.1:8000` |
| `ไม่พบ index.html ใน ...` | แตกไฟล์ zip ไม่ครบ ให้แตกใหม่ทั้งโฟลเดอร์ |
| `เปิดพอร์ต 8000 ไม่ได้` | พอร์ตถูกใช้อยู่ → ใช้ `--port 8080` |
| `ModuleNotFoundError: requests` | ยังไม่ติดตั้ง dependency → `pip install -r requirements.txt` |
| โดน throttle (429) | สคริปต์ retry ให้อัตโนมัติ; เพิ่ม `--batch-sleep 0.5` ถ้าข้อมูลเยอะ |

---

## 9. สถานะการทดสอบ

ทดสอบแล้วในสภาพแวดล้อมนี้ (รันจริงกับ mock Microsoft Graph server):

| รายการทดสอบ | ผล |
|---|---|
| ไวยากรณ์ Python + YAML ของ workflow | ผ่าน |
| เลือกชื่อ secret อัตโนมัติ (ตั้ง `AZURE_CLIENT_ID` ปนกับ `AZ_TENANT_ID`) | เลือกถูกทั้งคู่ และ secret แสดงเป็น `<hidden, N chars>` |
| `--dry-run` ไม่มี credential | ตรวจไฟล์ได้ + จับ key column ผิดได้ |
| ส่งเป็นชุดจาก CSV | รอบแรก `created`, รอบสอง `updated` — upsert ถูกต้อง |
| `GET /` เสิร์ฟ `index.html` | HTTP 200, `text/html; charset=utf-8` |
| `GET /api/fields` | คืน 19 ฟิลด์ บังคับ 4 ฟิลด์ |
| `POST /api/submit` upsert ครั้งแรก | สร้างสำเร็จ คืน item ID + ลิงก์ DispForm |
| `POST /api/submit` upsert ซ้ำ `Customer_id` เดิม | อัปเดต item เดิม (ไม่สร้างซ้ำ) |
| โหมดตรวจสอบอย่างเดียว | คืน `Customer_x0020_Name` และ `limit: "5000000"` |
| กรอกไม่ครบ / JSON ผิดรูปแบบ | HTTP 400 พร้อมข้อความภาษาไทย |
| `/api/health`, `/api/schema` | คืน JSON ถูกต้อง (503 เมื่อยังไม่เชื่อมต่อ) |
| โหมด `--dry-run` | ไม่เขียนจริง แสดง payload ให้ตรวจ |
| network/TLS error | แจ้งข้อความชัดเจน ไม่ crash เป็น traceback |

**ทดสอบฝั่งเบราว์เซอร์** (รัน JS ของ `index.html` จริงด้วย DOM shim + ยิงไป server จริง) — ผ่าน 13/13:
render ฟอร์ม 19 ฟิลด์, dropdown 6 ตัว, option ภาษาไทยถูกต้อง, เครื่องหมาย `*` 4 จุด,
แสดงสถานะเชื่อมต่อ, เตือนเมื่อกรอกไม่ครบ, แสดง payload ในโหมดตรวจสอบ,
map `Customer Name` → `Customer_x0020_Name`, ตัดลูกน้ำ `5,000,000` → `5000000`,
บันทึกแล้วแสดง item ID + ลิงก์ DispForm, บันทึกประวัติ, ล้างฟอร์มหลังสำเร็จ

สิ่งที่ยังต้องทำบน tenant จริง: ใส่ Client Secret ใน GitHub Secrets และ grant
API permission + admin consent ตามข้อ 3
