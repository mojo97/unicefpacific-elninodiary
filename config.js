/* Live data connection. See GOOGLE_SHEETS_SETUP.md for sharing and update steps. */
window.DIARY_CONFIG = {
  mode: "google-sheets",
  refreshMinutes: 5,
  googleSheets: {
    workbookUrl: "https://docs.google.com/spreadsheets/d/1x0kGVCRSWDESMxIMSiLlYj6bwV217AZ_wRXRVMIPxZk/export?format=xlsx",
    dataUrl: "https://docs.google.com/spreadsheets/d/1x0kGVCRSWDESMxIMSiLlYj6bwV217AZ_wRXRVMIPxZk/gviz/tq?tqx=out:csv&sheet=Activity%20Diary",
    displayName: "UNICEF El Niño Activity Diary",
    keyMessages: {
      dataUrl: "https://docs.google.com/spreadsheets/d/1x0kGVCRSWDESMxIMSiLlYj6bwV217AZ_wRXRVMIPxZk/gviz/tq?tqx=out:csv&gid=1470373217",
      sourceUrl: "https://docs.google.com/spreadsheets/d/1x0kGVCRSWDESMxIMSiLlYj6bwV217AZ_wRXRVMIPxZk/edit?gid=1470373217#gid=1470373217"
    },
    responsePlans: {
      dataUrl: "https://docs.google.com/spreadsheets/d/1x0kGVCRSWDESMxIMSiLlYj6bwV217AZ_wRXRVMIPxZk/gviz/tq?tqx=out:csv&gid=542411509",
      sourceUrl: "https://docs.google.com/spreadsheets/d/1x0kGVCRSWDESMxIMSiLlYj6bwV217AZ_wRXRVMIPxZk/edit?gid=542411509#gid=542411509"
    }
  }
};
