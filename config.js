/* Data connection settings. See SHAREPOINT_SETUP.md before enabling live mode. */
window.DIARY_CONFIG = {
  mode: "graph",
  refreshMinutes: 5,
  sharePoint: {
    tenantId: "YOUR_TENANT_ID",
    clientId: "YOUR_ENTRA_APP_CLIENT_ID",
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
