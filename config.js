/* Live data connection. See GOOGLE_SHEETS_SETUP.md for sharing and update steps. */
window.DIARY_CONFIG = {
  mode: "google-sheets",
  refreshMinutes: 5,
  /* Optional public ArcGIS embed URL. Leave blank until the approved map link is available. */
  arcgisMapUrl: "https://services9.arcgis.com/RHVPKKiFTONKtxq3/arcgis/rest/services/Active_Hurricanes_v1/FeatureServer",
  /* Optional approved endpoint returning JSON such as { "active": 12 }. */
  viewerCounterUrl: "",
  hazardEvents: [
    {
      country: "Cook Islands",
      location: "Northern Cook Islands",
      event: "Severe Tropical Cyclone Xavier",
      date: "2026-11-19",
      category: "Category 4",
      latitude: -10.9,
      longitude: -165.8,
      source: "https://cookislands.travel/tropical-cyclone-faqs"
    }
  ],
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
    },
    sectorResponsePlan: {
      dataUrl: "https://docs.google.com/spreadsheets/d/1x0kGVCRSWDESMxIMSiLlYj6bwV217AZ_wRXRVMIPxZk/gviz/tq?tqx=out:csv&gid=1281880942"
    },
    emergencyDeclarations: {
      dataUrl: "https://docs.google.com/spreadsheets/d/1x0kGVCRSWDESMxIMSiLlYj6bwV217AZ_wRXRVMIPxZk/gviz/tq?tqx=out:csv&gid=1366090099"
    }
  }
};
