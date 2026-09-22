/* Live data connection. See GOOGLE_SHEETS_SETUP.md for sharing and update steps. */
window.DIARY_CONFIG = {
  mode: "google-sheets",
  refreshMinutes: 5,
  googleSheets: {
    dataUrl: "https://docs.google.com/spreadsheets/d/1x0kGVCRSWDESMxIMSiLlYj6bwV217AZ_wRXRVMIPxZk/gviz/tq?tqx=out:csv&sheet=Activity%20Diary",
    displayName: "UNICEF El Niño Activity Diary"
  }
};
