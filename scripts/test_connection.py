import os
import requests
from msal import ConfidentialClientApplication

tenant = os.environ["AZURE_TENANT_ID"]
client = os.environ["AZURE_CLIENT_ID"]
secret = os.environ["AZURE_CLIENT_SECRET"]
app = ConfidentialClientApplication(client, authority=f"https://login.microsoftonline.com/{tenant}", client_credential=secret)
token = app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])
if "access_token" not in token:
    raise RuntimeError(token.get("error_description", str(token)))
headers = {"Authorization": "Bearer " + token["access_token"]}
site = requests.get("https://graph.microsoft.com/v1.0/sites/dohomegroup.sharepoint.com:/sites/KYC", headers=headers, timeout=60)
site.raise_for_status()
site_id = site.json()["id"]
lists = requests.get(f"https://graph.microsoft.com/v1.0/sites/{site_id}/lists?$select=id,displayName", headers=headers, timeout=60)
lists.raise_for_status()
match = next((x for x in lists.json().get("value", []) if x.get("displayName", "").lower() == "admin_kyc"), None)
if not match:
    raise RuntimeError("Admin_Kyc list not found")
print("Connection successful. Site and Admin_Kyc list are accessible.")
