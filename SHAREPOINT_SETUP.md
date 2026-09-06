# SharePoint live connection — administrator guide

This dashboard is a browser-based, read-only Microsoft Graph client. It uses delegated sign-in, so the signed-in user's existing Microsoft 365 access determines whether the workbook can be read.

## 1. Confirm the Excel source

- Store the workbook in the intended SharePoint document library.
- Confirm that the operational data are in the Excel table named `ElNinoActivityDiary` on the **Activity Diary** sheet.
- Keep the existing 21 table headers unchanged.
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

## 3. Find the workbook identifiers

Use Microsoft Graph Explorer or an approved administrative script while signed in to the UNICEF tenant.

You need:

- `siteId` — the SharePoint site containing the document library
- `driveId` — the document library
- `itemId` — the workbook file

Typical Microsoft Graph requests are:

```text
GET https://graph.microsoft.com/v1.0/sites/{hostname}:/{server-relative-site-path}
GET https://graph.microsoft.com/v1.0/sites/{siteId}/drives
GET https://graph.microsoft.com/v1.0/sites/{siteId}/drives/{driveId}/root:/{path-to-workbook}
```

Copy the returned `id` values. The exact site path, document-library drive, and workbook path depend on the UNICEF SharePoint structure, so the public sharing URL by itself should not be guessed into these fields.

## 4. Configure the website

Edit `config.js`:

```js
window.DIARY_CONFIG = {
  mode: "graph",
  refreshMinutes: 5,
  sharePoint: {
    tenantId: "PASTE_TENANT_ID",
    clientId: "PASTE_APPLICATION_CLIENT_ID",
    siteId: "PASTE_SITE_ID",
    driveId: "PASTE_DRIVE_ID",
    itemId: "PASTE_EXCEL_ITEM_ID",
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

The site calls this read endpoint after Microsoft sign-in:

```text
GET /sites/{siteId}/drives/{driveId}/items/{itemId}/workbook/tables/ElNinoActivityDiary/rows
```

The page converts each table row to the matching column header, discards unused blank template rows, and recalculates every view. It refreshes every five minutes while open. The source of truth remains the SharePoint workbook.

## Troubleshooting

| Symptom | Check |
|---|---|
| Sign-in says redirect URI mismatch | The Entra SPA redirect URI must exactly match the GitHub Pages URL, including path and trailing slash |
| Access denied | Viewer lacks workbook access, tenant consent is missing, or Graph delegated permissions are not approved |
| Table not found | Excel table must be named `ElNinoActivityDiary`, not merely a worksheet range |
| Dashboard is empty | Confirm the **Activity Diary** table contains dated rows and the workbook is saved |
| Embed is blocked | Allow the `github.io` domain in SharePoint HTML Field Security or use an approved internal host |
| Sign-in fails inside embed | Allow pop-ups/third-party sign-in, test direct URL, or use a SharePoint Framework deployment |
| Map marker is approximate | Current workbook has no coordinates; markers use country centroids |
