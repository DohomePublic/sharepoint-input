# ติดตั้ง Admin KYC Form

1. อัปโหลดไฟล์ทั้งหมดในโฟลเดอร์นี้ไปยัง Repository `DohomePublic/sharepoint-web` ที่ branch `main`
2. ใน Microsoft Entra App Registration เพิ่มแพลตฟอร์ม Single-page application และ Redirect URI:
   `https://dohomepublic.github.io/sharepoint-web/`
3. เพิ่ม Microsoft Graph Delegated permission `Sites.ReadWrite.All` และ Grant admin consent
4. ที่ GitHub ให้ไป Settings > Pages > Source แล้วเลือก GitHub Actions
5. Push หรือเปิด Actions แล้วรัน `Deploy GitHub Pages`
6. เปิด `https://dohomepublic.github.io/sharepoint-web/`

## GitHub Secrets สำหรับ Connection Test

เพิ่ม Secrets ต่อไปนี้ใน Settings > Secrets and variables > Actions:

- `AZURE_CLIENT_ID`
- `AZURE_TENANT_ID`
- `AZURE_CLIENT_SECRET`

Client Secret ใช้เฉพาะ Workflow ทดสอบการเชื่อมต่อ และห้ามใส่ใน HTML/JavaScript
