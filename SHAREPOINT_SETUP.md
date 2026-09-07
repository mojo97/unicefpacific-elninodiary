# SharePoint live connection — administrator guide

This dashboard is a browser-based, read-only Microsoft Graph client. It uses delegated sign-in, so the signed-in user's existing Microsoft 365 access determines whether the workbook can be read.

## 1. Confirm the Excel source

- Store the workbook in the intended SharePoint document library.
- Confirm that the operational data are in the Excel table named `ElNinoActivityDiary` on the **Activity Diary** sheet.
- Keep the existing 21 table headers unchanged.
- Keep the source header `UNICEF Unit*` unchanged. The dashboard presents it to users as **UNICEF Sector**.
- Confirm that intended viewers have at least read access to the file.
- Do not use the **Example Entries** sheet as operational data.

## 2. Register the browser application in Microsoft Entra ID

An Entra administrator should:

1. Open **Microsoft Entra admin center → App registrations → New registration**.
2. Use a clear name such as `UNICEF El Nino Activity Diary Dashboard`.
3. Choose accounts in this organisational directory only.
4. Under **Authentication**, add a **Single-page application (SPA)** redirect URI equal to the final GitHub Pages URL, including the trailing slash, for example:

   `https://YOUR-GITHUB-NAME.github.io/unicef-el-nino-diary/`

5. Under **API permissions → Microsoft Graph → Delegated permissions**, add:

   - `Files.Read.All`
   - `Sites.Read.All`

6. Grant administrator consent if required by the UNICEF tenant's policies.
7. Record the **Directory (tenant) ID** and **Application (client) ID**.
8. Do **not** create or place a client secret in this website.

The application requests read permissions only. For tighter least-privilege control, an Entra/SharePoint administrator may replace this design with a tenant-managed solution using site-scoped permissions or SharePoint Framework.

## 3. Confirm the preconfigured workbook location

The dashboard has been configured from the supplied SharePoint workbook link:

```text
Hostname: unicef.sharepoint.com
Site path: /teams/FJI-Program
Document library: Emergency
Workbook path: /2026 El Niño/UNICEF_El_Nino_Activity_Diary_Reporting_Template_v1.xlsx
```

The dashboard resolves the SharePoint Site ID, Drive ID and Item ID automatically after the user signs in. Confirm that `Emergency` is the document library name and that the workbook has not been moved or renamed.

The automatic read sequence is:

```text
GET https://graph.microsoft.com/v1.0/sites/unicef.sharepoint.com:/teams/FJI-Program
GET https://graph.microsoft.com/v1.0/sites/{siteId}/drives
GET https://graph.microsoft.com/v1.0/drives/{driveId}/root:/2026 El Niño/UNICEF_El_Nino_Activity_Diary_Reporting_Template_v1.xlsx
GET https://graph.microsoft.com/v1.0/drives/{driveId}/items/{itemId}/content
```

## 4. Configure the website

Edit `config.js`:

```js
window.DIARY_CONFIG = {
  mode: "graph",
  refreshMinutes: 5,
  sharePoint: {
    tenantId: "PASTE_TENANT_ID",
    clientId: "PASTE_APPLICATION_CLIENT_ID",
    hostname: "unicef.sharepoint.com",
    sitePath: "/teams/FJI-Program",
    libraryName: "Emergency",
    filePath: "/2026 El Niño/UNICEF_El_Nino_Activity_Diary_Reporting_Template_v1.xlsx",
    siteId: "",
    driveId: "",
    itemId: "",
    tableName: "ElNinoActivityDiary"
  }
};
```

Commit the change to GitHub. These identifiers identify resources but are not passwords. Never add passwords, access tokens, or client secrets to the repository.

## 5. Validate

1. Open the GitHub Pages URL in a private browser window.
2. Select **Connect Microsoft 365**.
3. Sign in with an authorised UNICEF account.
4. Confirm that the banner says **Live SharePoint data**.
5. Compare the activity count and one complete record with the Excel table.
6. Add or amend a safe test row in Excel, save it, and select **Refresh data**.
7. Verify filters, timeline order, map country, reach totals, funding total, and the 21-field detail record.
8. Test the embedded version inside SharePoint with a normal viewer account, not only an administrator account.

## 6. How live refresh works

After Microsoft sign-in, the site resolves the configured location and requests the current workbook file:

```text
GET /drives/{driveId}/items/{itemId}/content
```

The page reads the **Activity Diary** sheet locally in the signed-in user's browser, matches the 21 column headers, discards unused blank template rows, and recalculates every view. It refreshes every five minutes while open. The source of truth remains the SharePoint workbook.

## Troubleshooting

| Symptom | Check |
|---|---|
| Sign-in says redirect URI mismatch | The Entra SPA redirect URI must exactly match the GitHub Pages URL, including path and trailing slash |
| Access denied | Viewer lacks workbook access, tenant consent is missing, or Graph delegated permissions are not approved |
| Document library not found | Confirm the library is named `Emergency`; update `libraryName` if its actual name differs |
| Workbook not found | Confirm the `2026 El Niño` folder and workbook filename exactly match `config.js` |
| Activity Diary header not found | Keep the **Activity Diary** sheet and the original 21 column headings |
| Dashboard is empty | Confirm the **Activity Diary** table contains dated rows and the workbook is saved |
| Embed is blocked | Allow the `github.io` domain in SharePoint HTML Field Security or use an approved internal host |
| Sign-in fails inside embed | Allow pop-ups/third-party sign-in, test direct URL, or use a SharePoint Framework deployment |
| Map marker is approximate | Current workbook has no coordinates; markers use country centroids |
