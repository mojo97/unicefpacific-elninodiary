# UNICEF Pacific El Niño Activity Diary

This is a static, GitHub Pages–ready dashboard for the UNICEF Pacific El Niño Activity Diary workbook. It provides an overview, management signals, monthly analysis, a chronological activity timeline, a country-level map, a searchable register, a complete record view, and a print-ready country Situation Report.

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
| UNICEF Sector | Filters, charts, timeline and table. The source workbook field remains `UNICEF Unit*` for compatibility. |
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

## Country Situation Report

Open **Situation report** in the navigation, then select one country or **All countries**. An individual-country report is grouped by UNICEF Sector. The All countries report keeps every country separate and then groups its activities by UNICEF Sector, using the hierarchy **Country → UNICEF Sector → interventions, partners, people reached/targeted, gaps and priority actions**.

The report masthead is titled **El Niño Situation Report**. When All countries is selected, the report uses a regional-overview subtitle and labels every sector row with its actual country name, for example **Fiji · UNICEF Sector**. It does not use “All countries” as a country name inside the report.

Reach labels respond to implementation status. Records that are only Planned or On Hold are shown as **people/children targeted**. Records that are only Ongoing or Completed are shown as **people/children reached**. When both groups are present, the label becomes **reached / targeted** and the card notes that reported and planned reach are combined.

Funding labels follow the same rule. Planned or On Hold records show **Funding required**; Ongoing or Completed records show **Funding used**; and a mixed selection shows **Funding used / required**.

| Situation report section | Activity Diary source |
|---|---|
| UNICEF interventions | `Activity Title*` and `What Was Done?*` |
| Partners | `Partners` |
| Number of people and children reached / targeted | `People Reached (Total)` and `Children Reached`, labelled from `Implementation Status*` |
| Gaps & priority actions | `Challenges` and `Next Step` |
| Key figures | Activities, People Reached, Children Reached, Funding Used and Implementation Status |
| Context and assurance | Country, Location, Phase, Activity Date, Submission Date and Evidence Link |

The country selector is specific to the report and begins with **All countries**. The dashboard's UNICEF Sector, Phase, Status and Search filters also narrow the report. Select **Print / Save PDF** to open the browser print window; choose **Save as PDF** to create a landscape A4 briefing page.

The report is a draft synthesized from activity records, not a replacement for a validated humanitarian needs assessment. Review the generated narrative, totals, evidence and any personal information before external circulation. With many sectors or unusually long narratives, the browser may continue the report onto a second page rather than cut content.

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
7. Upload the *contents* of this folder—`index.html`, `app.js`, `styles.css`, `config.js`, `README.md`, `GOOGLE_SHEETS_SETUP.md`, and the `data` folder—not a parent folder around them.
8. Commit the files to the `main` branch.
9. In the repository, open **Settings → Pages**.
10. Under **Build and deployment**, choose **Deploy from a branch**.
11. Select branch `main`, folder `/(root)`, then **Save**.
12. After a few minutes, GitHub shows a URL similar to:

    `https://YOUR-GITHUB-NAME.github.io/unicef-el-nino-diary/`

At this point the site reads the configured live Google Sheet.

## Part 3 — Live Google Sheet

The dashboard is already configured to read this workbook:

```text
https://docs.google.com/spreadsheets/d/1x0kGVCRSWDESMxIMSiLlYj6bwV217AZ_wRXRVMIPxZk/edit
```

Set the Google Sheet's General access to **Anyone with the link — Viewer**. The dashboard downloads the current workbook, reads the **Activity Diary** sheet, and refreshes every five minutes or whenever **Refresh** is selected. See [GOOGLE_SHEETS_SETUP.md](GOOGLE_SHEETS_SETUP.md).

Important: link-accessible Google Sheets are not suitable for confidential or personally identifiable data. Use an authenticated internal application if UNICEF policy requires restricted access.

The site displays **UNICEF Sector** throughout. Keep the workbook header `UNICEF Unit*` unchanged because the template instructs users to preserve column names. The Excel loader also accepts `UNICEF Sector*`, `UNICEF Sectors*`, or `UNICEF Unit / Sector*` if a future workbook version uses one of those headings.

## Part 4 — Add the dashboard to a SharePoint page

1. Open the SharePoint communication or team site page.
2. Select **Edit**.
3. Select the **+** where the dashboard should appear.
4. Add the **Embed** web part.
5. Paste the GitHub Pages URL.
6. Make the web part wide or full-width if that layout is available.
7. Publish or republish the SharePoint page.

If SharePoint blocks the embedded URL, ask the site administrator to allow your `github.io` domain in SharePoint HTML Field Security. The dashboard itself no longer uses SharePoint for data.

## Regular reporting workflow

1. Staff update approved rows in the shared Google Sheet.
2. Keep one activity per row and retain the exact headers.
3. The dashboard reads the saved table after its next refresh.
4. Users can filter by UNICEF sector, country, phase, status, date-related text, partners, results, challenges, next steps, focal point, and any other field through search.
5. Select any timeline or register record to review all 21 fields.

## Files

- `index.html` — page structure
- `styles.css` — responsive UNICEF-inspired visual design
- `app.js` — analysis, timeline, map, filters, detail view, and Google Sheets sync
- `config.js` — connection settings
- `data/activities.json` — clearly labelled demonstration data only
- `GOOGLE_SHEETS_SETUP.md` — live-source setup, update steps and security notes

## Support and safety

- The dashboard's live connection is read-only; editors change data directly in Google Sheets.
- Anyone who can access the public dashboard may be able to retrieve the link-accessible source workbook.
- Do not publish personal, sensitive, or programme-confidential content without UNICEF approval.
- Review evidence URLs and personal information before enabling a public-facing deployment.
