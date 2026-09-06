/* Data connection settings. See SHAREPOINT_SETUP.md before enabling live mode. */
window.DIARY_CONFIG = {
  mode: "local",
  refreshMinutes: 5,
  sharePoint: {
    tenantId: "YOUR_TENANT_ID",
    clientId: "YOUR_ENTRA_APP_CLIENT_ID",
    siteId: "YOUR_SHAREPOINT_SITE_ID",
    driveId: "YOUR_DOCUMENT_LIBRARY_DRIVE_ID",
    itemId: "YOUR_EXCEL_FILE_ITEM_ID",
    tableName: "ElNinoActivityDiary"
  }
};
