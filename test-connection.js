// scripts/test-connection.js
// ทดสอบว่า GitHub Secrets ที่ตั้งไว้ใช้งานได้จริง โดยขอ token แบบ app-only (client credentials)
// แล้วลองอ่านข้อมูล site + list จาก SharePoint ผ่าน Microsoft Graph
//
// รองรับชื่อ Secret สองแบบ: AZ_* (ใช้จริง) และ AZURE_* (ชื่อเดิม) — เลือกตัวที่มีค่าให้อัตโนมัติ
// รัน: node scripts/test-connection.js
// ต้องการ Node.js 18+ (มี fetch ในตัว)

function pickEnv(azName, azureName) {
  if (process.env[azName]) return { value: process.env[azName], source: azName };
  if (process.env[azureName]) return { value: process.env[azureName], source: azureName };
  return { value: null, source: null };
}

const clientId = pickEnv("AZ_CLIENT_ID", "AZURE_CLIENT_ID");
const tenantId = pickEnv("AZ_TENANT_ID", "AZURE_TENANT_ID");
const clientSecret = pickEnv("AZ_CLIENT_SECRET", "AZURE_CLIENT_SECRET");

const spHostname = process.env.SP_HOSTNAME || "dohomegroup.sharepoint.com";
const spSitePath = process.env.SP_SITE_PATH || "/sites/AC-Accounting";
const spListName =
  process.env.SP_LIST_NAME || "DemoApp ต้องการบึนทึกเข้า Sharepoint";

function fail(msg) {
  console.error(`✗ ${msg}`);
  process.exit(1);
}

async function main() {
  console.log("== ตรวจสอบตัวแปรที่อ่านได้ (ไม่แสดงค่าจริง) ==");
  for (const [label, v] of [
    ["Client ID", clientId],
    ["Tenant ID", tenantId],
    ["Client Secret", clientSecret],
  ]) {
    if (!v.value) fail(`ไม่พบค่าสำหรับ ${label} (ตรวจสอบ Settings > Secrets and variables > Actions)`);
    console.log(`  ${label}: อ่านจากตัวแปร ${v.source} (มีค่าแล้ว)`);
  }

  console.log("\n== ขอ access token (client credentials flow) ==");
  const tokenRes = await fetch(
    `https://login.microsoftonline.com/${tenantId.value}/oauth2/v2.0/token`,
    {
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body: new URLSearchParams({
        client_id: clientId.value,
        client_secret: clientSecret.value,
        scope: "https://graph.microsoft.com/.default",
        grant_type: "client_credentials",
      }),
    }
  );
  const tokenBody = await tokenRes.json();
  if (!tokenRes.ok) {
    fail(`ขอ token ไม่สำเร็จ (HTTP ${tokenRes.status}): ${tokenBody.error_description || JSON.stringify(tokenBody)}`);
  }
  console.log("  ✓ ได้ access token แล้ว");
  const token = tokenBody.access_token;

  console.log("\n== อ่านข้อมูล site ==");
  const siteRes = await fetch(
    `https://graph.microsoft.com/v1.0/sites/${spHostname}:${spSitePath}`,
    { headers: { Authorization: `Bearer ${token}` } }
  );
  const site = await siteRes.json();
  if (!siteRes.ok) {
    fail(`อ่าน site ไม่สำเร็จ (HTTP ${siteRes.status}): ${JSON.stringify(site)}. ตรวจสอบว่าให้ admin consent สิทธิ์ Sites.ReadWrite.All (application) แล้ว`);
  }
  console.log(`  ✓ พบ site: ${site.displayName || site.name} (id: ${site.id})`);

  console.log("\n== ค้นหา list ==");
  const listsRes = await fetch(
    `https://graph.microsoft.com/v1.0/sites/${site.id}/lists?$filter=displayName eq '${encodeURIComponent(
      spListName
    )}'`,
    { headers: { Authorization: `Bearer ${token}` } }
  );
  const lists = await listsRes.json();
  if (!listsRes.ok) fail(`อ่าน list ไม่สำเร็จ (HTTP ${listsRes.status}): ${JSON.stringify(lists)}`);
  if (!lists.value || lists.value.length === 0) {
    fail(`ไม่พบ List ชื่อ "${spListName}" ใน site นี้ — ตรวจสอบชื่อ list ให้ตรงกับของจริง`);
  }
  const list = lists.value[0];
  console.log(`  ✓ พบ list: ${list.displayName} (id: ${list.id})`);

  console.log("\n== อ่านจำนวนรายการใน list ==");
  const itemsRes = await fetch(
    `https://graph.microsoft.com/v1.0/sites/${site.id}/lists/${list.id}/items?$top=1`,
    { headers: { Authorization: `Bearer ${token}` } }
  );
  const items = await itemsRes.json();
  if (!itemsRes.ok) fail(`อ่าน items ไม่สำเร็จ (HTTP ${itemsRes.status}): ${JSON.stringify(items)}`);

  console.log("\n✓ การเชื่อมต่อทั้งหมดสำเร็จ — Secrets และสิทธิ์ตั้งค่าถูกต้อง");
}

main().catch((e) => fail(e.message));
