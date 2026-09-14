ADMIN KYC FORM
1. Upload index.html to the root of GitHub repository: DohomePublic/sharepoint-web
2. In Microsoft Entra ID > App Registration > Authentication, add SPA redirect URI:
   https://dohomepublic.github.io/sharepoint-web/
3. Add Microsoft Graph Delegated permission: Sites.ReadWrite.All and grant admin consent.
4. GitHub Pages: deploy from main branch / root.
5. Open https://dohomepublic.github.io/sharepoint-web/ and sign in with Microsoft 365.

Important: AZURE_CLIENT_SECRET is not used in index.html and must never be placed in browser code.
