// config.js
// ค่าทั้งหมดในไฟล์นี้ "ไม่ใช่ความลับ" (Client ID / Tenant ID เป็นตัวระบุแอปที่เปิดเผยได้)
// ห้ามใส่ Client Secret ในไฟล์นี้หรือไฟล์ใดๆ ที่อยู่ใน repository เด็ดขาด

window.KYC_CONFIG = {
  // --- Microsoft Entra ID (Azure AD) App registration ---
  clientId: "a37bd62d-e74d-4ea0-9546-1eb5aa96f604",
  tenantId: "7f8918d9-718a-495b-ac9a-17cba381c4a0",
  // ต้องตรงกับ Redirect URI ที่ลงทะเบียนไว้ใน Authentication > Single-page application ทุกตัวอักษร
  redirectUri: "https://dohomepublic.github.io/sharepoint-input/",

  // Delegated scopes ที่ขอตอน login (ต้อง grant admin consent ไว้แล้ว)
  scopes: ["User.Read", "Sites.ReadWrite.All"],

  // --- SharePoint target ---
  spHostname: "dohomegroup.sharepoint.com",
  spSitePath: "/sites/AC-Accounting",
  // ชื่อ List ตามที่แสดงใน SharePoint (Display name) — แก้ให้ตรงกับของจริง
  spListName: "DemoApp ต้องการบึนทึกเข้า Sharepoint",

  // --- การ mapping ฟิลด์ในฟอร์ม -> ชื่อคอลัมน์ภายใน (internal name) ของ SharePoint List ---
  // "internal name" ไม่ใช่ชื่อที่เห็นบนหน้าจอ SharePoint เสมอไป (เช่นภาษาไทย/เว้นวรรค SharePoint จะเปลี่ยนเป็น _x0020_ ฯลฯ)
  // วิธีหา internal name จริง: ไปที่ List Settings > คลิกที่ชื่อคอลัมน์ > ดู "Field" ใน URL (ParamField=...)
  // หรือเรียก GET /sites/{siteId}/lists/{listId}/columns ผ่าน Graph Explorer แล้วดู field "name"
  fieldMapping: {
    fullName: "Title", // คอลัมน์ Title มีอยู่แล้วในทุก List ของ SharePoint
    customerType: "CustomerType",
    idNumber: "IDNumber",
    address: "Address",
    phone: "Phone",
    email: "Email",
    notes: "Notes",
  },
};
