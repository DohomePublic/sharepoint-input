# DOHOME SharePoint KYC Input

พร้อมอัปโหลดทับใน repository `DohomePublic/sharepoint-input`

## ติดตั้ง
1. แตก ZIP
2. อัปโหลดไฟล์ทั้งหมดไปยัง root ของ repository
3. Commit ไปที่ branch `main`
4. GitHub > Settings > Pages > Deploy from a branch > `main` / `(root)`
5. Microsoft Entra ID > App registration > Authentication > Single-page application
6. เพิ่ม Redirect URI: `https://dohomepublic.github.io/sharepoint-input/`
7. เพิ่ม Delegated permissions: `User.Read`, `Sites.ReadWrite.All` และ Grant admin consent

## SharePoint
- Site: `/sites/AC-Accounting`
- List: `DemoApp`

## ความปลอดภัย
ห้ามใส่ Client Secret ในไฟล์หน้าเว็บ
