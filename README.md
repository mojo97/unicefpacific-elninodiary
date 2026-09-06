# UNICEF Pacific El Niño Activity Diary

This is a static, GitHub Pages–ready dashboard for the UNICEF Pacific El Niño Activity Diary workbook. It provides an overview, management signals, monthly analysis, a chronological activity timeline, a country-level map, a searchable register, and a complete record view.

The header now includes the UNICEF logo. Use of the UNICEF name and logo should remain limited to an authorised UNICEF deployment and follow current organisational brand guidance.

## Important before you publish

The uploaded workbook's **Activity Diary** table is empty. The included preview uses the five rows from **Example Entries**, which the workbook marks as illustrative. The dashboard therefore displays a prominent **Demonstration** notice.

For real reporting, enter approved rows in the Excel table named `ElNinoActivityDiary` on the **Activity Diary** sheet. Do not rename that table or its 21 column headings.

## What the dashboard uses

Every workbook field is used:

| Workbook field | Where it appears |
|---|---|
| Entry ID | Record identity and detail view |
| Activity Date | Timeline, charts, table, record detail |
| Reporting Month | Record detail and monthly analysis support |
| UNICEF Unit | Filters, charts, timeline, table |
| Country | Filters, map, timeline, table |
| Location / Admin Area | Map popup and record detail |
| El Niño Phase | Filters, phase chart, timeline |
| Activity Type | Timeline and record detail |
| Activity Title | Cards, timeline, table, map popup |
| What Was Done? | Timeline and record detail |
| Result / Output | Latest activity cards and record detail |
| People Reached (Total) | KPI, charts, table, map popup |
| Children Reached | KPI, timeline, record detail |
| Partners | Record detail |
| Implementation Status | Filters, status badges, management signal |
| Funding Used (USD) | KPI, table, record detail |
| Challenges | Management follow-up signal and record detail |
| Next Step | Management follow-up signal and record detail |
| Evidence Link | Record detail |
| Focal Point | Record detail |
| Submission Date | Reporting-lag signal and record detail |

The workbook has no latitude or longitude columns. The map therefore places activities at country centroids and shows the reported Location / Admin Area as text. To show exact locations later, add validated `Latitude` and `Longitude` columns and extend the map logic.

## Part 1 — Preview it on your computer

Because browsers block some local-file data loading, use a tiny local web server instead of double-clicking `index.html`.

1. Install Python 3 from <https://www.python.org/downloads/> if it is not already installed.
2. Unzip this package.
3. Open Terminal (macOS), PowerShell (Windows), or a terminal in VS Code.
4. Change into the unzipped folder, for example:

   ```bash
   cd path/to/UNICEF_El_Nino_Activity_Diary_Site
   ```

5. Start the server:

   ```bash
   python -m http.server 8000
   ```

6. Open <http://localhost:8000> in your browser.

## Part 2 — Publish it with GitHub Pages

1. Create a free GitHub account at <https://github.com/>.
2. Select **New repository**.
3. Name it, for example, `unicef-el-nino-diary`.
4. Choose **Public** only if UNICEF has approved making the site code public. The included demonstration data are illustrative; never put confidential workbook data in `data/activities.json`.
5. Select **Create repository**.
6. Choose **Add file → Upload files**.
7. Upload the *contents* of this folder—`index.html`, `app.js`, `styles.css`, `config.js`, `README.md`, `SHAREPOINT_SETUP.md`, and the `data` folder—not a parent folder around them.
8. Commit the files to the `main` branch.
9. In the repository, open **Settings → Pages**.
10. Under **Build and deployment**, choose **Deploy from a branch**.
11. Select branch `main`, folder `/(root)`, then **Save**.
12. After a few minutes, GitHub shows a URL similar to:

    `https://YOUR-GITHUB-NAME.github.io/unicef-el-nino-diary/`

At this point the site runs in demonstration mode.

## Load an Excel file manually — without SharePoint

Use this option when the workbook is on your computer, attached to an email, stored outside SharePoint, or when the Microsoft 365 connection has not been configured.

1. Open the published dashboard or the local preview.
2. Select **Load Excel** in the top-right corner.
3. Select an `.xlsx`, `.xlsm`, or `.xls` file from your computer.
4. Select **Open**.
5. Wait for the confirmation message showing how many activities were loaded.
6. Confirm that the source indicator in the header changes to `Excel · filename.xlsx`.
7. Review the activity count and one record before using the analysis.

The dashboard first looks for a sheet named **Activity Diary**. It finds the header row automatically, verifies that all 21 required columns are present, ignores blank template rows, and then rebuilds every chart, map, timeline, filter and table.

Important details:

- The selected Excel file is processed inside your browser and is not sent to GitHub or stored by the website.
- The file must retain the original 21 column headings.
- The Activity Diary sheet must contain at least one row with an Activity Date and Activity Title.
- Selecting **Refresh** re-reads the currently selected file.
- If you close or reload the webpage, select **Load Excel** again. Browsers do not allow the page to remember access to a local file permanently.
- This manual method shows the latest information at the moment you select the file. It does not automatically detect later changes made to the Excel file; load the saved file again after editing it.
- For automatic five-minute refresh from a shared source, use the SharePoint/Microsoft Graph setup below.

## Part 3 — Connect the real SharePoint workbook

A private SharePoint file link cannot safely be pasted into a public webpage. The dashboard uses Microsoft sign-in and Microsoft Graph instead. Ask your UNICEF Microsoft 365/Entra administrator to help with the one-time setup in [SHAREPOINT_SETUP.md](SHAREPOINT_SETUP.md).

After the administrator gives you the Tenant ID, App/Client ID, Site ID, Drive ID, and Excel Item ID:

1. Open `config.js` in GitHub.
2. Select the pencil icon to edit.
3. Change `mode: "local"` to `mode: "graph"`.
4. Replace each `YOUR_...` value with the administrator-provided ID.
5. Leave `tableName: "ElNinoActivityDiary"` unchanged.
6. Commit the change.
7. Open the GitHub Pages site and select **Connect Microsoft 365**.
8. Sign in with a UNICEF account that already has permission to the workbook.

The dashboard refreshes the table every five minutes and whenever **Refresh data** is selected. It does not copy the live workbook into the GitHub repository.

## Part 4 — Add the dashboard to a SharePoint page

1. Open the SharePoint communication or team site page.
2. Select **Edit**.
3. Select the **+** where the dashboard should appear.
4. Add the **Embed** web part.
5. Paste the GitHub Pages URL.
6. Make the web part wide or full-width if that layout is available.
7. Publish or republish the SharePoint page.

If SharePoint blocks the URL, ask the site administrator to allow your `github.io` domain in SharePoint HTML Field Security. Users may also need to allow the Microsoft sign-in pop-up. If your organisation prohibits public GitHub Pages or custom browser applications, deploy the same interface as a SharePoint Framework web part instead.

## Regular reporting workflow

1. Staff update approved rows in the SharePoint-hosted Excel table.
2. Keep one activity per row and retain the exact headers.
3. The dashboard reads the saved table after its next refresh.
4. Users can filter by unit, country, phase, status, date-related text, partners, results, challenges, next steps, focal point, and any other field through search.
5. Select any timeline or register record to review all 21 fields.

## Files

- `index.html` — page structure
- `styles.css` — responsive UNICEF-inspired visual design
- `app.js` — analysis, timeline, map, filters, detail view, and Graph sync
- `config.js` — connection settings
- `data/activities.json` — clearly labelled demonstration data only
- `SHAREPOINT_SETUP.md` — administrator setup and security notes

## Support and safety

- The live connection is read-only.
- Workbook access continues to be controlled by Microsoft 365 permissions.
- Do not add a client secret to `config.js`; browser apps must not contain secrets.
- Do not publish personal, sensitive, or programme-confidential content without UNICEF approval.
- Review evidence URLs and personal information before enabling a public-facing deployment.
