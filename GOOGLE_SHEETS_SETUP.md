# Google Sheets live data — setup and update guide

The dashboard reads this Google workbook automatically:

`https://docs.google.com/spreadsheets/d/1x0kGVCRSWDESMxIMSiLlYj6bwV217AZ_wRXRVMIPxZk/edit`

## Required sharing setting

The GitHub Pages dashboard cannot sign in to Google. The workbook must therefore be available as **Anyone with the link — Viewer**, or be published to the web. Do not use this setup for confidential or personally identifiable data.

1. Open the Google Sheet.
2. Select **Share**.
3. Under **General access**, select **Anyone with the link**.
4. Choose **Viewer**.
5. Select **Done**.

If UNICEF policy does not allow link-accessible operational data, do not publish the dashboard publicly. Use an authenticated internal application instead.

## Updating the live dashboard

1. Edit or paste approved records into the **Activity Diary** sheet.
2. Keep all 21 required column headings unchanged.
3. Keep one activity per row and ensure Activity Date and Activity Title are completed.
4. Keep the Key Messages headers as `Key Messages` / `Link`, and the country-plan headers as `Country` / `Response Plan Link`. The link cells may be full web addresses or Google Sheets hyperlink/smart-chip cells. You may add `Sector` to Key Messages; otherwise the site infers it from the title and filename.
5. On the UNICEF sector response-plan tab (`gid=1281880942`), retain the existing headers for Sector, Country, Indicators, priority activity/location/group, people in need, targets, demographic breakdown, target unit, non-person target, unit cost, total funding required and notes. Enter numeric values without words in target, cost and funding columns.
6. On the Emergency Declaration tab (`gid=1366090099`), keep the headers `Country`, `Date` and `Source`. Use one declaration per row, a readable date, and the official source URL.
7. Wait for Google Sheets to save the changes.
8. Open the dashboard and select **Refresh**. The dashboard also checks for updates every five minutes.

The dashboard has one data control: **Refresh**. It always fetches the latest saved version of the shared Google Sheet.

## Troubleshooting

| Message or issue | What to check |
|---|---|
| Google Sheets returned 401 or 403 | Change General access to **Anyone with the link — Viewer** |
| Activity Diary header not found | Keep the sheet name **Activity Diary** and the required headings |
| Missing-column message | Restore the exact workbook column headings |
| Old data remain visible | Wait for Google to save, then select **Refresh** and hard-refresh the browser |
| Live data unavailable | The live sheet could not be read; check sharing and the browser/network policy |
| Funding shows “Not reported” | Complete `Total Funding required (USD)` for that response-plan row; a blank is intentionally not treated as zero |

The loader tolerates harmless heading differences such as `El Niño Phase*`, `El Nino Phase`, extra spaces, and missing asterisks. It still requires every reporting field to be present.
