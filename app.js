const FIELD_ORDER = [
  "Entry ID", "Activity Date*", "Reporting Month", "UNICEF Unit*", "Country*",
  "Location / Admin Area", "El Niño Phase*", "Activity Type*", "Activity Title*",
  "What Was Done?*", "Result / Output", "People Reached (Total)", "Children Reached",
  "Partners", "Implementation Status*", "Funding Used (USD)", "Challenges", "Next Step",
  "Evidence Link", "Focal Point*", "Submission Date*"
];
const FIELD_ALIASES = {
  "UNICEF Unit*": ["UNICEF Sector*", "UNICEF Sectors*", "UNICEF Unit / Sector*"],
  "Submission Date*": ["Reporting Date", "Reporting Date*", "Date Reported", "Date Reported*"]
};

const COUNTRY_COORDS = {
  "Pacific region / Multi-country": [-8, 168], "Australia": [-25.27, 133.78],
  "Cook Islands": [-21.24, -159.78], "Federated States of Micronesia": [7.43, 150.55],
  "Fiji": [-17.71, 178.07], "French Polynesia": [-17.68, -149.41], "Guam": [13.44, 144.79],
  "Kiribati": [1.87, -157.36], "Marshall Islands": [7.13, 171.18], "Nauru": [-0.52, 166.93],
  "New Caledonia": [-20.90, 165.62], "New Zealand": [-40.90, 174.89], "Niue": [-19.05, -169.87],
  "Northern Mariana Islands": [15.10, 145.67], "Palau": [7.51, 134.58], "Papua New Guinea": [-6.31, 143.96],
  "Samoa": [-13.76, -172.10], "Solomon Islands": [-9.65, 160.16], "Tokelau": [-9.20, -171.85],
  "Tonga": [-21.18, -175.20], "Tuvalu": [-7.11, 177.65], "Vanuatu": [-15.38, 166.96],
  "Wallis and Futuna": [-13.77, -177.16], "American Samoa": [-14.27, -170.13]
};

const PHASE_COLORS = { "Preparedness": "#38bdf8", "Anticipatory Action": "#f7bf45", "Response": "#ef6f61", "Recovery": "#43b982" };
const SECTOR_COLORS = {
  "education": "#1f77b4",
  "wash": "#00a7b5",
  "health": "#d64b4b",
  "nutrition": "#f28e2b",
  "child protection": "#8b5fbf",
  "social protection": "#59a14f",
  "communication": "#e3b341",
  "communications": "#e3b341",
  "programme effectiveness": "#4e6e81",
  "supply": "#d45087"
};
const SECTOR_PALETTE = ["#008dc9","#00a7b5","#ef6f61","#f28e2b","#8b5fbf","#59a14f","#e3b341","#4e6e81","#d45087","#7a9e9f"];
const config = window.DIARY_CONFIG || { mode: "local", refreshMinutes: 5 };
let allActivities = [], filteredActivities = [], keyMessages = [], responsePlans = [], sectorPlans = [], emergencyDeclarations = [], metadata = { isDemo: true }, map, markerLayer, supportingWorkbookPromise;

const $ = (id) => document.getElementById(id);
const fmtNum = new Intl.NumberFormat("en-US", { maximumFractionDigits: 0 });
const fmtUSD = new Intl.NumberFormat("en-US", { style: "currency", currency: "USD", maximumFractionDigits: 0 });
const fmtDate = new Intl.DateTimeFormat("en-GB", { day: "numeric", month: "short", year: "numeric" });
const formatDateValue = value => {
  const date = new Date(value);
  return !value || Number.isNaN(date.getTime()) ? "Not reported" : fmtDate.format(date);
};
const escapeHtml = (value) => String(value ?? "—").replace(/[&<>'"]/g, c => ({"&":"&amp;","<":"&lt;",">":"&gt;","'":"&#39;",'"':"&quot;"}[c]));
const phaseColor = (phase) => PHASE_COLORS[phase] || "#91a6af";
const sectorColor = sector => {
  const name = String(sector || "Sector not reported").trim().toLowerCase();
  if (SECTOR_COLORS[name]) return SECTOR_COLORS[name];
  const known = Object.keys(SECTOR_COLORS).find(key => name.includes(key));
  if (known) return SECTOR_COLORS[known];
  const hash = [...name].reduce((total,char) => ((total * 31) + char.charCodeAt(0)) >>> 0, 0);
  return SECTOR_PALETTE[hash % SECTOR_PALETTE.length];
};
const canonicalHeader = value => String(value ?? "")
  .normalize("NFD")
  .replace(/[\u0300-\u036f]/g, "")
  .replace(/&/g, " and ")
  .replace(/\*/g, "")
  .replace(/[^a-z0-9]+/gi, " ")
  .trim()
  .toLowerCase();

function excelDate(value) {
  if (!value && value !== 0) return "";
  if (typeof value === "number" || /^\d{5}(\.\d+)?$/.test(String(value))) {
    const date = new Date(Math.round((Number(value) - 25569) * 86400 * 1000));
    return date.toISOString().slice(0, 10);
  }
  const date = new Date(value);
  return Number.isNaN(date.getTime()) ? String(value) : date.toISOString().slice(0, 10);
}

function normalizeActivity(raw) {
  const activity = {};
  FIELD_ORDER.forEach(field => { activity[field] = raw[field] ?? ""; });
  activity["Entry ID"] = String(activity["Entry ID"] ?? "").trim();
  activity["Activity Date*"] = excelDate(activity["Activity Date*"]);
  activity["Submission Date*"] = excelDate(activity["Submission Date*"]);
  activity["People Reached (Total)"] = Number(activity["People Reached (Total)"]) || 0;
  activity["Children Reached"] = Number(activity["Children Reached"]) || 0;
  activity["Funding Used (USD)"] = Number(activity["Funding Used (USD)"]) || 0;
  if (!activity["Reporting Month"] && activity["Activity Date*"]) {
    const monthDate = new Date(activity["Activity Date*"]);
    if (!Number.isNaN(monthDate.getTime())) activity["Reporting Month"] = monthDate.toLocaleDateString("en-US", { month: "short", year: "numeric" });
  }
  const coords = COUNTRY_COORDS[activity["Country*"]] || [raw.latitude, raw.longitude];
  activity.latitude = Number.isFinite(Number(raw.latitude)) ? Number(raw.latitude) : Number(coords?.[0]);
  activity.longitude = Number.isFinite(Number(raw.longitude)) ? Number(raw.longitude) : Number(coords?.[1]);
  return activity;
}

function parseExcelData(source, sourceDetails) {
  if (!window.XLSX) throw new Error("The Excel reader did not load. Check your internet connection and try again.");
  const workbook = XLSX.read(source, { type: typeof source === "string" ? "string" : "array", cellDates: true });
  const preferredSheet = workbook.Sheets["Activity Diary"] ? "Activity Diary" : null;
  const candidates = preferredSheet ? [preferredSheet] : workbook.SheetNames;
  let selected;

  for (const sheetName of candidates) {
    const rows = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { header: 1, defval: "", raw: true });
    const headerIndex = rows.findIndex(row => {
      const labels = row.map(canonicalHeader);
      return labels.includes("entry id") && labels.includes("activity date") && labels.includes("activity title");
    });
    if (headerIndex >= 0) { selected = { sheetName, rows, headerIndex }; break; }
  }

  if (!selected) throw new Error("No Activity Diary header row was found in this Excel file.");
  const headers = selected.rows[selected.headerIndex].map(value => String(value ?? "").trim());
  const sourceHeaderIndex = field => {
    const accepted = [field,...(FIELD_ALIASES[field] || [])].map(canonicalHeader);
    return headers.findIndex(header => accepted.includes(canonicalHeader(header)));
  };
  const missing = FIELD_ORDER.filter(field => sourceHeaderIndex(field) < 0);
  if (missing.length) throw new Error(`The live sheet is missing required field${missing.length === 1 ? "" : "s"}: ${missing.join(", ")}`);

  const activities = selected.rows.slice(selected.headerIndex + 1).map(row => Object.fromEntries(FIELD_ORDER.map(field => [field,row[sourceHeaderIndex(field)]])))
    .filter(row => row["Activity Date*"] !== "" && row["Activity Date*"] !== null && row["Activity Title*"] !== "")
    .map(normalizeActivity);
  if (!activities.length) throw new Error(`The ${selected.sheetName} sheet contains no completed activity rows.`);

  metadata = { isDemo: false, ...sourceDetails, sheet: selected.sheetName, generatedAt: new Date().toISOString() };
  return activities;
}

function parseCsvRows(source, requiredHeaders) {
  if (!window.XLSX) throw new Error("The Google Sheet reader did not load.");
  const workbook = XLSX.read(source, { type: "string", cellDates: true });
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "", raw: false });
  const required = requiredHeaders.map(canonicalHeader);
  const headerIndex = rows.findIndex(row => {
    const labels = row.map(canonicalHeader);
    return required.every(header => labels.includes(header));
  });
  if (headerIndex < 0) throw new Error(`Required columns were not found: ${requiredHeaders.join(", ")}`);
  const headers = rows[headerIndex].map(value => String(value ?? "").trim());
  return rows.slice(headerIndex + 1).map(row => Object.fromEntries(headers.map((header,index) => [header,row[index] ?? ""]))).filter(row => Object.values(row).some(value => String(value).trim()));
}

function rowValue(row, names) {
  const wanted = names.map(canonicalHeader);
  const key = Object.keys(row).find(header => wanted.includes(canonicalHeader(header)));
  return key ? String(row[key] ?? "").trim() : "";
}

function directUrl(value) {
  const text = String(value || "").trim();
  return /^https?:\/\/[^\s]+$/i.test(text) ? text : "";
}

function workbookRows(workbook, sheetNames, requiredHeaders) {
  const wantedSheets = sheetNames.map(canonicalHeader);
  const sheetName = workbook.SheetNames.find(name => wantedSheets.includes(canonicalHeader(name)));
  if (!sheetName) throw new Error(`The ${sheetNames[0]} tab was not found.`);
  const sheet = workbook.Sheets[sheetName];
  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "", raw: false });
  const required = requiredHeaders.map(canonicalHeader);
  const headerIndex = rows.findIndex(row => {
    const labels = row.map(canonicalHeader);
    return required.every(header => labels.includes(header));
  });
  if (headerIndex < 0) throw new Error(`Required columns were not found in ${sheetName}: ${requiredHeaders.join(", ")}`);
  const headers = rows[headerIndex].map(value => String(value ?? "").trim());
  return rows.slice(headerIndex + 1).map((row,rowOffset) => {
    const values = Object.fromEntries(headers.map((header,columnIndex) => [header,row[columnIndex] ?? ""]));
    values.__links = Object.fromEntries(headers.map((header,columnIndex) => {
      const cell = sheet[XLSX.utils.encode_cell({ r: headerIndex + rowOffset + 1, c: columnIndex })];
      return [canonicalHeader(header),String(cell?.l?.Target || "").trim()];
    }));
    return values;
  }).filter(row => Object.entries(row).some(([key,value]) => key !== "__links" && String(value).trim()));
}

function rowLink(row, names) {
  for (const name of names) {
    const link = directUrl(row.__links?.[canonicalHeader(name)]);
    if (link) return link;
  }
  return "";
}

async function loadSupportingWorkbook() {
  if (supportingWorkbookPromise) return supportingWorkbookPromise;
  supportingWorkbookPromise = (async() => {
    const url = String(config.googleSheets?.workbookUrl || "").trim();
    if (!url) throw new Error("The Google workbook export link is not configured.");
    const separator = url.includes("?") ? "&" : "?";
    const response = await fetch(`${url}${separator}_=${Date.now()}`, { cache: "no-store" });
    if (!response.ok) throw new Error(`Google workbook export returned ${response.status}.`);
    return XLSX.read(await response.arrayBuffer(), { type: "array", cellDates: true });
  })();
  return supportingWorkbookPromise;
}

function inferSector(value) {
  const text = String(value || "").toLowerCase();
  if (/education|school|learning/.test(text)) return "Education";
  if (/nutrition|feeding|malnutrition/.test(text)) return "Health & Nutrition";
  if (/health|medical|clinic/.test(text)) return "Health & Nutrition";
  if (/wash|water|sanitation|hygiene/.test(text)) return "WASH";
  if (/child protection|protection|safeguard/.test(text)) return "Child Protection";
  if (/social.*behaviour|behavior|sbc|communication|cfm|feedback/.test(text)) return "Social & Behaviour Change";
  if (/supply|logistics|procurement/.test(text)) return "Supply";
  return "Cross-sector / General";
}

async function fetchCsvRows(sheetConfig, requiredHeaders) {
  const url = String(sheetConfig?.dataUrl || "").trim();
  if (!url) return [];
  const separator = url.includes("?") ? "&" : "?";
  const response = await fetch(`${url}${separator}_=${Date.now()}`, { cache: "no-store" });
  if (!response.ok) throw new Error(`Google Sheets returned ${response.status}.`);
  return parseCsvRows(await response.text(),requiredHeaders);
}

async function loadKeyMessages() {
  const sheetConfig = config.googleSheets?.keyMessages || {};
  let rows;
  try {
    rows = workbookRows(await loadSupportingWorkbook(),["Key Messages"],["Key Messages","Link"]);
  } catch (error) {
    rows = await fetchCsvRows(sheetConfig,["Key Messages","Link"]);
  }
  return rows.map(row => {
    const title = rowValue(row,["Key Messages","Key Message","Title"]);
    const documentLabel = rowValue(row,["Link","Document","File"]);
    const sector = rowValue(row,["Sector","UNICEF Sector","UNICEF Unit"]) || inferSector(`${title} ${documentLabel}`);
    const documentUrl = rowLink(row,["Link"]) || directUrl(documentLabel);
    return { title, documentLabel, sector, url: documentUrl };
  }).filter(item => item.title);
}

async function loadResponsePlans() {
  const sheetConfig = config.googleSheets?.responsePlans || {};
  let rows;
  try {
    rows = workbookRows(await loadSupportingWorkbook(),["Response Plan","Response Plans"],["Country","Response Plan Link"]);
  } catch (error) {
    rows = await fetchCsvRows(sheetConfig,["Country","Response Plan Link"]);
  }
  return rows.map(row => {
    const country = rowValue(row,["Country"]);
    const plan = rowValue(row,["Response Plan","Plan","Response Plan Link"]);
    const linkLabel = rowValue(row,["Response Plan Link"]);
    const planUrl = rowLink(row,["Response Plan Link"]) || directUrl(linkLabel);
    return { country, plan, url: planUrl };
  }).filter(item => item.country);
}

async function loadEmergencyDeclarations() {
  const sheetConfig = config.googleSheets?.emergencyDeclarations || {};
  const rows = await fetchCsvRows(sheetConfig,["Country","Date","Source"]);
  return rows.map(row => ({
    country: rowValue(row,["Country"]),
    date: rowValue(row,["Date","Declaration Date"]),
    source: directUrl(rowValue(row,["Source","Link"]))
  })).filter(item => item.country);
}

function planNumber(value) {
  const text = String(value ?? "").trim();
  if (!text || /^(xx|tbd|n\/a|na|-)$/i.test(text)) return null;
  const number = Number(text.replace(/[$,%\s,]/g,""));
  return Number.isFinite(number) ? number : null;
}

function normalizePlanCountry(value) {
  const country = String(value || "").trim();
  const key = canonicalHeader(country);
  if (key === "fsm" || key === "micronesia fsm" || key === "federated states of micronesia") return "Federated States of Micronesia";
  return country;
}

async function loadSectorResponsePlan() {
  const sheetConfig = config.googleSheets?.sectorResponsePlan || {};
  const rows = await fetchCsvRows(sheetConfig,["Sector","Country (select from dropdown)","Indicators (select from dropdown)"]);
  return rows.map((row,index) => ({
    id: `plan-${index + 1}`,
    sector: rowValue(row,["Sector"]),
    country: normalizePlanCountry(rowValue(row,["Country (select from dropdown)","Country"])),
    indicator: rowValue(row,["Indicators (select from dropdown)","Indicator","Indicators"]),
    activity: rowValue(row,["Priority Activities","Priority Activity"]),
    locations: rowValue(row,["Priority Locations","Priority Location"]),
    groups: rowValue(row,["Priority Groups","Priority Group"]),
    peopleInNeed: planNumber(rowValue(row,["People in Need (Number as per National Response plan)","People in Need"])),
    nationalTarget: planNumber(rowValue(row,["Target (National) (if applicable) based on National response plan","Target (National)"])),
    peopleTargeted: planNumber(rowValue(row,["People Targeted by UNICEF — Total","People Targeted by UNICEF - Total","People Targeted by UNICEF Total"])),
    unknown: planNumber(rowValue(row,["Sex / age unknown"])),
    girls: planNumber(rowValue(row,["Female (0 - 18) years"])),
    boys: planNumber(rowValue(row,["Male (0 - 18) years"])),
    women: planNumber(rowValue(row,["Female 18+ years and above"])),
    men: planNumber(rowValue(row,["Male 18 + years and above"])),
    targetUnit: rowValue(row,["Target unit (if not people)","Target unit"]),
    nonPersonTarget: planNumber(rowValue(row,["Non-person indicator target","Non person indicator target"])),
    unitCost: planNumber(rowValue(row,["per Unit cost","Unit cost"])),
    fundingRequired: planNumber(rowValue(row,["Total Funding required (USD)","Total Funding Required (USD)"])),
    notes: rowValue(row,["Additional Notes/Remarks","Notes","Remarks"])
  })).filter(item => item.sector && item.country && (item.indicator || item.activity));
}

async function loadLocalData() {
  const response = await fetch("data/activities.json", { cache: "no-store" });
  if (!response.ok) throw new Error("The demonstration data file could not be loaded.");
  const payload = await response.json();
  metadata = payload.metadata || { isDemo: true };
  return payload.activities.map(normalizeActivity);
}

async function loadGoogleSheetsData() {
  const url = String(config.googleSheets?.dataUrl || config.googleSheets?.workbookUrl || "").trim();
  if (!url) throw new Error("The Google Sheet link is not configured.");
  const separator = url.includes("?") ? "&" : "?";
  const response = await fetch(`${url}${separator}_=${Date.now()}`, { cache: "no-store" });
  if (!response.ok) throw new Error(`Google Sheets returned ${response.status}. Confirm that the sheet is shared with anyone who has the link.`);
  return parseExcelData(await response.text(), { source: "Google Sheets", fileName: config.googleSheets?.displayName || "El Niño Activity Diary" });
}

async function refreshData() {
  $("refresh-button").disabled = true;
  $("refresh-button").innerHTML = "<span>↻</span> Loading";
  supportingWorkbookPromise = null;
  try {
    allActivities = config.mode === "google-sheets" ? await loadGoogleSheetsData() : await loadLocalData();
    if (config.mode === "google-sheets") {
      const [messagesResult,plansResult,sectorPlanResult,declarationsResult] = await Promise.allSettled([loadKeyMessages(),loadResponsePlans(),loadSectorResponsePlan(),loadEmergencyDeclarations()]);
      keyMessages = messagesResult.status === "fulfilled" ? messagesResult.value : [];
      responsePlans = plansResult.status === "fulfilled" ? plansResult.value : [];
      sectorPlans = sectorPlanResult.status === "fulfilled" ? sectorPlanResult.value : [];
      emergencyDeclarations = declarationsResult.status === "fulfilled" ? declarationsResult.value : [];
      metadata.auxiliaryWarnings = [messagesResult,plansResult,sectorPlanResult,declarationsResult].filter(result => result.status === "rejected").map(result => result.reason?.message || "A supporting sheet could not be loaded.");
    } else {
      keyMessages = [];
      responsePlans = [];
      sectorPlans = [];
      emergencyDeclarations = [];
      metadata.auxiliaryWarnings = [];
    }
    allActivities.sort((a,b) => new Date(a["Activity Date*"]) - new Date(b["Activity Date*"]));
    $("demo-banner").classList.toggle("hidden", !metadata.isDemo);
    $("sync-state").classList.toggle("live", !metadata.isDemo);
    const sourceLabel = metadata.isDemo ? "Demonstration data" : "Live Google Sheet";
    $("sync-state").querySelector("span").textContent = sourceLabel;
    $("sync-state").title = sourceLabel;
    $("footer-source").textContent = metadata.isDemo ? "workbook example entries" : "live Google Sheet";
    populateFilters(); populatePlanFilters(); applyFilters();
    $("last-refreshed").textContent = new Date().toLocaleString("en-GB", { dateStyle: "medium", timeStyle: "short" });
    showToast(`${allActivities.length} activities and ${sectorPlans.length} UNICEF response-plan rows loaded.`);
    return true;
  } catch (error) {
    showToast(error.message || "Data refresh failed.");
    $("demo-banner").classList.remove("hidden");
    $("demo-banner").innerHTML = `<strong>Live data unavailable:</strong> ${escapeHtml(error.message || "The Google Sheet could not be loaded.")}`;
    $("sync-state").classList.remove("live");
    $("sync-state").querySelector("span").textContent = allActivities.length ? "Refresh failed · last loaded data" : "Live data unavailable";
    if (!allActivities.length) {
      allActivities = [];
      $("footer-source").textContent = "Google Sheet unavailable";
      populateFilters(); applyFilters();
    }
    return false;
  } finally {
    $("refresh-button").disabled = false;
    $("refresh-button").innerHTML = "<span>↻</span> Refresh";
  }
}

function unique(field) { return [...new Set(allActivities.map(d => d[field]).filter(Boolean))].sort(); }
function populateSelect(id, values, label) {
  const current = $(id).value;
  $(id).innerHTML = `<option value="">All ${label}</option>` + values.map(v => `<option>${escapeHtml(v)}</option>`).join("");
  if (values.includes(current)) $(id).value = current;
}
function populateFilters() {
  populateSelect("unit-filter", unique("UNICEF Unit*"), "sectors");
  populateSelect("country-filter", unique("Country*"), "countries");
  populateSelect("phase-filter", unique("El Niño Phase*"), "phases");
  populateSelect("status-filter", unique("Implementation Status*"), "statuses");
  const countries = unique("Country*");
  const currentReportCountry = $("sitrep-country").value;
  $("sitrep-country").innerHTML = `<option value="">All countries</option>` + countries.map(country => `<option value="${escapeHtml(country)}">${escapeHtml(country)}</option>`).join("");
  $("sitrep-country").value = countries.includes(currentReportCountry) ? currentReportCountry : "";
  const activityDates = allActivities.map(row => row["Activity Date*"]).filter(value => /^\d{4}-\d{2}-\d{2}$/.test(value)).sort();
  if (activityDates.length) {
    ["activity-start-filter","activity-end-filter"].forEach(id => {
      $(id).min = activityDates[0];
      $(id).max = activityDates[activityDates.length - 1];
    });
  }
}

function populatePlanFilters() {
  const fill = (id,values,label) => {
    const select = $(id);
    const current = select.value;
    select.innerHTML = `<option value="">All ${label}</option>` + values.map(value => `<option value="${escapeHtml(value)}">${escapeHtml(value)}</option>`).join("");
    if (values.includes(current)) select.value = current;
  };
  fill("plan-sector-filter",[...new Set(sectorPlans.map(item => item.sector).filter(Boolean))].sort(),"sectors");
  fill("plan-country-filter",[...new Set(sectorPlans.map(item => item.country).filter(Boolean))].sort(),"countries");
  fill("plan-indicator-filter",[...new Set(sectorPlans.map(item => item.indicator).filter(Boolean))].sort(),"indicators");
}

function applyFilters() {
  const search = $("search-filter").value.trim().toLowerCase();
  const unit = $("unit-filter").value, country = $("country-filter").value, phase = $("phase-filter").value, status = $("status-filter").value;
  const activityStart = $("activity-start-filter").value, activityEnd = $("activity-end-filter").value;
  if ($("sitrep-country") && $("sitrep-country").value !== country) $("sitrep-country").value = country;
  filteredActivities = allActivities.filter(d => {
    const haystack = FIELD_ORDER.map(field => d[field]).join(" ").toLowerCase();
    const activityDate = d["Activity Date*"] || "";
    return (!search || haystack.includes(search)) && (!unit || d["UNICEF Unit*"] === unit) && (!country || d["Country*"] === country) && (!phase || d["El Niño Phase*"] === phase) && (!status || d["Implementation Status*"] === status) && (!activityStart || activityDate >= activityStart) && (!activityEnd || activityDate <= activityEnd);
  });
  renderAll();
}

function sum(field) { return filteredActivities.reduce((total,d) => total + (Number(d[field]) || 0), 0); }
function reachMode(rows) {
  const normalizeStatus = value => String(value || "").trim().toLowerCase().replace(/[\s_-]+/g," ");
  const selectedStatus = normalizeStatus($("status-filter")?.value);
  if (selectedStatus === "planned" || selectedStatus === "on hold") return "targeted";
  if (selectedStatus === "ongoing" || selectedStatus === "completed") return "reached";
  const statuses = rows.map(row => normalizeStatus(row["Implementation Status*"])).filter(Boolean);
  if (!statuses.length) return "reached";
  const targetOnly = statuses.every(status => status.includes("planned") || status.includes("hold"));
  const deliveredOnly = statuses.every(status => status.includes("ongoing") || status.includes("completed"));
  return targetOnly ? "targeted" : (deliveredOnly ? "reached" : "reached / targeted");
}
function reachFoot(rows) {
  const mode = reachMode(rows);
  if (mode === "targeted") return "Planned reach reported";
  if (mode === "reached") return "Reported reach to date";
  return "Reported and planned reach combined";
}
function fundingMode(rows) {
  const reach = reachMode(rows);
  return reach === "targeted" ? "required" : (reach === "reached" ? "used" : "used / required");
}
function fundingFoot(rows) {
  const mode = fundingMode(rows);
  if (mode === "required") return "Planned funding requirement";
  if (mode === "used") return "Reported expenditure";
  return "Reported and required funding combined";
}
function renderKpis() {
  const countries = new Set(filteredActivities.map(d => d["Country*"])).size;
  const completed = filteredActivities.filter(d => d["Implementation Status*"] === "Completed").length;
  const completion = filteredActivities.length ? Math.round(completed / filteredActivities.length * 100) : 0;
  const mode = reachMode(filteredActivities);
  const cards = [
    ["Activities", fmtNum.format(filteredActivities.length), `${countries} locations represented`, true],
    [`People ${mode}`, fmtNum.format(sum("People Reached (Total)")), reachFoot(filteredActivities)],
    [`Children ${mode}`, fmtNum.format(sum("Children Reached")), reachFoot(filteredActivities)],
    [`Funding ${fundingMode(filteredActivities)}`, fmtUSD.format(sum("Funding Used (USD)")), fundingFoot(filteredActivities)],
    ["Completed", `${completion}%`, `${completed} of ${filteredActivities.length} activities`]
  ];
  $("kpi-grid").innerHTML = cards.map(([label,value,foot,highlight]) => `<article class="kpi-card ${highlight ? "highlight" : ""}"><span class="kpi-label">${label}</span><strong class="kpi-value">${value}</strong><span class="kpi-foot">${foot}</span></article>`).join("");
}

function group(field, measure) {
  const out = {};
  filteredActivities.forEach(d => { const key = d[field] || "Not reported"; out[key] = (out[key] || 0) + (measure ? Number(d[measure]) || 0 : 1); });
  return Object.entries(out).sort((a,b) => b[1] - a[1]);
}
function formatMonthLabel(value) {
  const text = String(value ?? "").trim();
  if (!text) return "Not reported";
  const monthNames = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];

  const googleDate = text.match(/^Date\((\d{4}),\s*(\d{1,2})(?:,\s*\d{1,2})?\)$/i);
  if (googleDate) return `${monthNames[Number(googleDate[2])]} ${googleDate[1]}`;

  const isoDate = text.match(/^(\d{4})[-/]([01]?\d)(?:[-/]\d{1,2})?/);
  if (isoDate && Number(isoDate[2]) >= 1 && Number(isoDate[2]) <= 12) return `${monthNames[Number(isoDate[2]) - 1]} ${isoDate[1]}`;

  const monthYear = text.match(/\b(Jan(?:uary)?|Feb(?:ruary)?|Mar(?:ch)?|Apr(?:il)?|May|Jun(?:e)?|Jul(?:y)?|Aug(?:ust)?|Sep(?:tember)?|Oct(?:ober)?|Nov(?:ember)?|Dec(?:ember)?)\b.*?\b(\d{4})\b/i);
  if (monthYear) return `${monthYear[1].slice(0,3)} ${monthYear[2]}`;

  const serialDate = typeof value === "number" || /^\d{5}(\.\d+)?$/.test(text) ? excelDate(value).match(/^(\d{4})-(\d{2})/) : null;
  if (serialDate) return `${monthNames[Number(serialDate[2]) - 1]} ${serialDate[1]}`;

  return text;
}

function monthSortKey(value, fallbackDate) {
  const label = formatMonthLabel(value || fallbackDate);
  const monthNames = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];
  const match = label.match(/^(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\s+(\d{4})$/i);
  if (match) return Date.UTC(Number(match[2]),monthNames.findIndex(month => month.toLowerCase() === match[1].toLowerCase()),1);
  const fallback = new Date(fallbackDate || value);
  return Number.isNaN(fallback.getTime()) ? Number.MAX_SAFE_INTEGER : Date.UTC(fallback.getUTCFullYear(),fallback.getUTCMonth(),1);
}
function renderMonthly() {
  const monthMap = new Map();
  filteredActivities.forEach(activity => {
    const rawMonth = activity["Reporting Month"] || activity["Activity Date*"] || "Not reported";
    const month = formatMonthLabel(rawMonth);
    const sortKey = monthSortKey(rawMonth,activity["Activity Date*"]);
    const sector = activity["UNICEF Unit*"] || "Sector not reported";
    if (!monthMap.has(month)) monthMap.set(month,{ month, sortKey, total: 0, sectors: {} });
    const entry = monthMap.get(month);
    entry.sortKey = Math.min(entry.sortKey,sortKey);
    entry.total += 1;
    entry.sectors[sector] = (entry.sectors[sector] || 0) + 1;
  });

  const months = [...monthMap.values()].sort((a,b) => a.sortKey - b.sortKey || a.month.localeCompare(b.month));
  const sectors = [...new Set(filteredActivities.map(activity => activity["UNICEF Unit*"] || "Sector not reported"))].sort();
  const max = Math.max(1,...months.map(month => month.total));
  $("monthly-chart").innerHTML = months.length ? months.map(({month,total,sectors:sectorCounts}) => {
    const segments = sectors.filter(sector => sectorCounts[sector]).map(sector => `<span class="month-bar-segment" style="flex-grow:${sectorCounts[sector]};background:${sectorColor(sector)}" title="${escapeHtml(sector)}: ${sectorCounts[sector]} ${sectorCounts[sector] === 1 ? "activity" : "activities"}"></span>`).join("");
    return `<div class="month-column"><span class="month-value">${total}</span><div class="month-bar" style="height:${Math.max(8,total/max*145)}px" aria-label="${escapeHtml(formatMonthLabel(month))}: ${total} activities by sector">${segments}</div><span class="month-label">${escapeHtml(formatMonthLabel(month))}</span></div>`;
  }).join("") : `<div class="empty-state">No activities match these filters.</div>`;
  $("monthly-legend").innerHTML = sectors.map(sector => `<span class="sector-legend-item"><i style="background:${sectorColor(sector)}"></i>${escapeHtml(sector)}</span>`).join("");
  const dates = filteredActivities.map(d => new Date(d["Activity Date*"])).filter(d => !isNaN(d));
  $("date-range-label").textContent = dates.length ? `${fmtDate.format(new Date(Math.min(...dates)))} – ${fmtDate.format(new Date(Math.max(...dates)))}` : "No selected dates";
}
function renderPhase() {
  const phases = Object.keys(PHASE_COLORS).map(p => [p, filteredActivities.filter(d => d["El Niño Phase*"] === p).length]);
  const total = Math.max(1, filteredActivities.length); let cursor = 0;
  const stops = phases.map(([p,n]) => { const start = cursor; cursor += n/total*100; return `${phaseColor(p)} ${start}% ${cursor}%`; }).join(",");
  $("phase-donut").innerHTML = `<div class="donut" style="background:conic-gradient(${stops || "#e8eef0 0 100%"})"><strong>${filteredActivities.length}</strong></div>`;
  $("phase-legend").innerHTML = phases.map(([p,n]) => `<div class="legend-row"><i class="legend-dot" style="background:${phaseColor(p)}"></i><span>${p}</span><strong>${n}</strong></div>`).join("");
  $("phase-ribbon").innerHTML = phases.map(([p,n]) => `<div class="phase-step" style="--phase-color:${phaseColor(p)}"><span>${n} ${n === 1 ? "activity" : "activities"}</span><strong>${p}</strong></div>`).join("");
}
function renderRankBars(id, entries, formatter = fmtNum.format.bind(fmtNum)) {
  const max = Math.max(1,...entries.map(d => d[1]));
  $(id).innerHTML = entries.length ? entries.map(([label,value]) => `<div class="rank-row"><span class="rank-label" title="${escapeHtml(label)}">${escapeHtml(label)}</span><div class="rank-track"><div class="rank-fill" style="width:${value/max*100}%"></div></div><strong class="rank-number">${formatter(value)}</strong></div>`).join("") : `<div class="empty-state">No values available.</div>`;
}
function renderFollowup() {
  const ongoing = filteredActivities.filter(d => d["Implementation Status*"] !== "Completed").length;
  const challenges = filteredActivities.filter(d => d["Challenges"]).length;
  const nextSteps = filteredActivities.filter(d => d["Next Step"]).length;
  const evidence = filteredActivities.filter(d => /^https?:\/\//.test(d["Evidence Link"])).length;
  const items = [["↗",`${ongoing} still active`,"Ongoing, planned or on-hold activities"],["!",`${challenges} challenges logged`,"Constraints needing management attention"],["→",`${nextSteps} next steps`,"Forward actions documented"],["✓",`${evidence}/${filteredActivities.length} evidence links`,"Records with a supporting URL"]];
  $("followup-list").innerHTML = items.map(([icon,title,text]) => `<div class="followup-item"><span class="followup-icon">${icon}</span><div><strong>${title}</strong><p>${text}</p></div></div>`).join("");
}

function statusClass(status) { return `status-${String(status || "").toLowerCase().replace(/\s+/g,"-")}`; }
function renderLatest() {
  const latest = [...filteredActivities].sort((a,b)=>new Date(b["Activity Date*"])-new Date(a["Activity Date*"])).slice(0,3);
  $("latest-activities").innerHTML = latest.length ? latest.map(d => `<button class="activity-card" data-entry="${escapeHtml(d["Entry ID"])}" type="button"><div class="meta"><span>${fmtDate.format(new Date(d["Activity Date*"]))} · ${escapeHtml(d["UNICEF Unit*"])}</span><span class="status-pill ${statusClass(d["Implementation Status*"])}">${escapeHtml(d["Implementation Status*"])}</span></div><h3>${escapeHtml(d["Activity Title*"])}</h3><p>${escapeHtml(d["Country*"])} · ${escapeHtml(d["Result / Output"] || "Result not reported")}</p></button>`).join("") : `<div class="empty-state">No activities match these filters.</div>`;
}
function renderTimeline() {
  const rows = [...filteredActivities].sort((a,b)=>new Date(b["Activity Date*"])-new Date(a["Activity Date*"]));
  $("activity-explorer-count").textContent = `${rows.length} ${rows.length === 1 ? "activity" : "activities"}`;
  $("timeline").innerHTML = rows.length ? rows.map(d => `<article class="timeline-entry" style="--phase-color:${phaseColor(d["El Niño Phase*"])}"><time class="timeline-date">${fmtDate.format(new Date(d["Activity Date*"]))}</time><button class="timeline-card" data-entry="${escapeHtml(d["Entry ID"])}" type="button"><div class="timeline-card-head"><div><span class="phase-pill" style="color:${phaseColor(d["El Niño Phase*"])};background:${phaseColor(d["El Niño Phase*"])}18">${escapeHtml(d["El Niño Phase*"])}</span><h3>${escapeHtml(d["Activity Title*"])}</h3></div><span class="status-pill ${statusClass(d["Implementation Status*"])}">${escapeHtml(d["Implementation Status*"])}</span></div><p>${escapeHtml(d["What Was Done?*"])}</p><div class="timeline-tags"><span class="soft-tag">${escapeHtml(d["UNICEF Unit*"])}</span><span class="soft-tag">${escapeHtml(d["Country*"])}</span><span class="soft-tag">${escapeHtml(d["Activity Type*"])}</span><span class="soft-tag">${fmtNum.format(d["Children Reached"])} children</span></div></button></article>`).join("") : `<div class="empty-state">No activities match these filters.</div>`;
}

function initMap() {
  if (map || !window.L) return;
  map = L.map("map-canvas", { worldCopyJump: true, minZoom: 2 }).setView([-7, 173], 3);
  L.tileLayer("https://tile.openstreetmap.org/{z}/{x}/{y}.png", { maxZoom: 18, attribution: '&copy; <a href="https://www.openstreetmap.org/copyright">OpenStreetMap</a>' }).addTo(map);
  markerLayer = L.layerGroup().addTo(map); L.control.scale({ imperial: false }).addTo(map);
}
function renderMap() {
  if (!map) return;
  markerLayer.clearLayers(); const bounds = [];
  filteredActivities.forEach(d => {
    if (!Number.isFinite(d.latitude) || !Number.isFinite(d.longitude)) return;
    let lng = d.longitude < 100 ? d.longitude + 360 : d.longitude;
    const hasDeclaration = emergencyDeclarations.some(item => sameCountry(item.country,d["Country*"]));
    const marker = L.circleMarker([d.latitude, lng], {
  radius: hasDeclaration ? 11 : 9,
  color: "#fff",
  weight: hasDeclaration ? 3 : 2,

  // Red is used only when a declaration exists.
  fillColor: hasDeclaration ? "#d71920" : "#27aae1",

  fillOpacity: 0.95
});
    marker.bindPopup(`${hasDeclaration ? `<p class="map-popup-alert">Emergency/disaster declaration recorded</p>` : ""}<h3>${escapeHtml(d["Activity Title*"])}</h3><p><strong>${escapeHtml(d["Country*"])}</strong> · ${escapeHtml(d["Location / Admin Area"])}</p><p>${escapeHtml(d["UNICEF Unit*"])} · ${escapeHtml(d["Implementation Status*"])}</p><p>${fmtDate.format(new Date(d["Activity Date*"]))}</p>`);
    marker.on("click",() => selectMapCountry(d["Country*"]));
    marker.addTo(markerLayer); bounds.push([d.latitude,lng]);
  });
  if (bounds.length) map.fitBounds(bounds,{padding:[40,40],maxZoom:5});
  const countries = group("Country*");
  $("map-count").textContent = `${countries.length} mapped ${countries.length === 1 ? "location" : "locations"}`;
  $("map-country-list").innerHTML = countries.map(([name,count]) => {
    const declared = emergencyDeclarations.some(item => sameCountry(item.country,name));
    return `<button class="country-row" data-map-country="${escapeHtml(name)}" type="button"><span>${escapeHtml(name)}${declared ? `<small class="declaration-badge">Declared</small>` : ""}</span><strong>${count}</strong></button>`;
  }).join("");
  renderResponsePlans();
  renderEmergencyDeclarations();
}

function countryKey(value) {
  const key = canonicalHeader(value);
  const aliases = {
    "fsm": "federated states of micronesia",
    "micronesia fsm": "federated states of micronesia",
    "rmi": "marshall islands",
    "republic of the marshall islands": "marshall islands"
  };
  return aliases[key] || key;
}

function sameCountry(left,right) { return countryKey(left) === countryKey(right); }

function declarationDateLabel(value) {
  const date = new Date(String(value || "").trim());
  return Number.isNaN(date.getTime()) ? (value || "Date not reported") : fmtDate.format(date);
}

function selectMapCountry(country) {
  const filter = $("country-filter");
  if (!filter || ![...filter.options].some(option => option.value === country)) return;
  filter.value = country;
  applyFilters();
}

function renderResponsePlans() {
  const container = $("map-response-plans");
  const selectedCountry = $("country-filter").value;
  if (!selectedCountry) {
    container.innerHTML = `<p class="supporting-empty">Select a country on the map or in the country list to see its response plan.</p>`;
    return;
  }
  const matches = responsePlans.filter(item => sameCountry(item.country,selectedCountry));
  if (!matches.length) {
    container.innerHTML = `<p class="supporting-empty">No response-plan row is available for ${escapeHtml(selectedCountry)}.</p>`;
    return;
  }
  const available = matches.filter(item => item.plan);
  if (!available.length) {
    container.innerHTML = `<p class="supporting-empty">${escapeHtml(selectedCountry)} is listed, but its response plan has not yet been added.</p>`;
    return;
  }
  container.innerHTML = available.map(item => `<article class="response-plan-card"><strong>${escapeHtml(item.plan)}</strong>${item.url ? `<a href="${escapeHtml(item.url)}" target="_blank" rel="noopener">Open response plan ↗</a>` : `<span class="missing-resource-link">Response Plan Link is not available</span>`}</article>`).join("") + `<button class="map-clear-country" data-map-country="" type="button">Show all countries</button>`;
}

function renderEmergencyDeclarations() {
  const container = $("map-declarations");
  const selectedCountry = $("country-filter").value;
  if (!selectedCountry) {
    container.innerHTML = `<p class="supporting-empty">Select a country to see whether an El Niño emergency or disaster declaration is recorded.</p>`;
    return;
  }
  const matches = emergencyDeclarations.filter(item => sameCountry(item.country,selectedCountry));
  if (!matches.length) {
    container.innerHTML = `<article class="declaration-card not-recorded"><span class="declaration-state">NOT RECORDED</span><strong>No declaration is listed for ${escapeHtml(selectedCountry)}</strong><p>This reflects the live declaration tab and does not confirm that no declaration exists.</p></article>`;
    return;
  }
  container.innerHTML = matches.map(item => `<article class="declaration-card"><span class="declaration-state">DECLARATION RECORDED</span><strong>${escapeHtml(item.country)}</strong><time>${escapeHtml(declarationDateLabel(item.date))}</time>${item.source ? `<a href="${escapeHtml(item.source)}" target="_blank" rel="noopener">Open official source ↗</a>` : `<span class="missing-resource-link">Source link is not available</span>`}</article>`).join("");
}

function renderKeyMessages() {
  const selectedSector = $("unit-filter").value;
  const rows = keyMessages.filter(item => !selectedSector || canonicalHeader(item.sector) === canonicalHeader(selectedSector));
  const warning = metadata.auxiliaryWarnings?.length ? ` ${metadata.auxiliaryWarnings.join(" ")}` : "";
  $("key-message-status").textContent = `${rows.length} ${rows.length === 1 ? "resource" : "resources"} shown from the live Key Messages tab.${warning}`;
  if (!rows.length) {
    $("key-messages-groups").innerHTML = `<div class="empty-state">No key messages match the selected UNICEF Sector.</div>`;
    return;
  }
  const groups = [...new Set(rows.map(item => item.sector || "Cross-sector / General"))].sort();
  $("key-messages-groups").innerHTML = groups.map(sector => {
    const messages = rows.filter(item => (item.sector || "Cross-sector / General") === sector);
    return `<section class="key-message-sector" style="--sector-color:${sectorColor(sector)}"><header><div><p>UNICEF SECTOR</p><h3>${escapeHtml(sector)}</h3></div><span>${messages.length} ${messages.length === 1 ? "resource" : "resources"}</span></header><div class="key-message-grid">${messages.map(item => `<article class="key-message-card"><div><span class="key-message-type">KEY MESSAGE RESOURCE</span><h4>${escapeHtml(item.title)}</h4>${item.documentLabel ? `<p>${escapeHtml(item.documentLabel)}</p>` : ""}</div>${item.url ? `<a href="${escapeHtml(item.url)}" target="_blank" rel="noopener">Open resource ↗</a>` : `<span class="missing-resource-link">Link is not available</span>`}</article>`).join("")}</div></section>`;
  }).join("");
}

function renderRegister() {
  $("register-body").innerHTML = filteredActivities.length ? [...filteredActivities].sort((a,b)=>new Date(b["Activity Date*"])-new Date(a["Activity Date*"])).map(d => `<tr><td>${fmtDate.format(new Date(d["Activity Date*"]))}</td><td><strong>${escapeHtml(d["UNICEF Unit*"])}</strong></td><td>${escapeHtml(d["Country*"])}<br><small>${escapeHtml(d["Location / Admin Area"])}</small></td><td>${escapeHtml(d["Activity Title*"])}</td><td>${escapeHtml(d["El Niño Phase*"])}</td><td><span class="status-pill ${statusClass(d["Implementation Status*"])}">${escapeHtml(d["Implementation Status*"])}</span></td><td>${fmtNum.format(d["People Reached (Total)"])}</td><td>${fmtUSD.format(d["Funding Used (USD)"])}</td><td><button class="row-button" data-entry="${escapeHtml(d["Entry ID"])}" type="button">View →</button></td></tr>`).join("") : `<tr><td colspan="9" class="empty-state">No activities match these filters.</td></tr>`;
}

function filteredSectorPlans() {
  const sector = $("plan-sector-filter").value;
  const country = $("plan-country-filter").value;
  const indicator = $("plan-indicator-filter").value;
  const funding = $("plan-funding-filter").value;
  return sectorPlans.filter(item =>
    (!sector || item.sector === sector) &&
    (!country || item.country === country) &&
    (!indicator || item.indicator === indicator) &&
    (!funding || (funding === "reported" ? item.fundingRequired !== null : item.fundingRequired === null))
  );
}

function renderPlanBars(id,rows,field) {
  const values = {};
  rows.forEach(item => {
    const label = item[field] || "Not reported";
    if (!values[label]) values[label] = { total: 0, reported: 0 };
    if (item.fundingRequired !== null) {
      values[label].total += item.fundingRequired;
      values[label].reported += 1;
    }
  });
  const entries = Object.entries(values).sort((a,b) => b[1].total - a[1].total || a[0].localeCompare(b[0]));
  const max = Math.max(1,...entries.map(([,value]) => value.total));
  $(id).innerHTML = entries.length ? entries.map(([label,value]) => `<div class="plan-bar-row"><div class="plan-bar-label"><span>${escapeHtml(label)}</span><strong>${value.reported ? fmtUSD.format(value.total) : "Not reported"}</strong></div><div class="plan-bar-track"><i class="${value.reported ? "" : "unreported"}" style="width:${value.reported ? Math.max(2,value.total/max*100) : 0}%;background:${field === "sector" ? sectorColor(label) : "#008dc9"}"></i></div></div>`).join("") : `<div class="empty-state">No funding values match the selected filters.</div>`;
}

function planMetric(value,empty = "Not reported") {
  return value === null ? empty : fmtNum.format(value);
}

function renderPlanPivot(rows) {
  const rowField = $("plan-pivot-row").value;
  const columnField = $("plan-pivot-column").value;
  const measure = $("plan-pivot-value").value;
  const label = { country: "Country", sector: "Sector", indicator: "Indicator" };
  const rowLabels = [...new Set(rows.map(item => item[rowField] || "Not reported"))].sort();
  const columnLabels = [...new Set(rows.map(item => item[columnField] || "Not reported"))].sort();
  const metric = item => {
    if (measure === "records") return { value: 1, valid: true };
    if (measure === "funding") return { value: item.fundingRequired || 0, valid: item.fundingRequired !== null };
    if (measure === "target") return { value: item.peopleTargeted || 0, valid: item.peopleTargeted !== null };
    return { value: item.peopleInNeed || 0, valid: item.peopleInNeed !== null };
  };
  const format = value => measure === "funding" ? fmtUSD.format(value) : fmtNum.format(value);
  const cell = (rowLabel,columnLabel) => {
    const matches = rows.filter(item => (item[rowField] || "Not reported") === rowLabel && (item[columnField] || "Not reported") === columnLabel);
    return matches.reduce((total,item) => { const current = metric(item); total.value += current.value; total.reported += current.valid ? 1 : 0; return total; }, { value: 0, reported: 0 });
  };
  const matrix = rowLabels.map(rowLabel => ({ rowLabel, cells: columnLabels.map(columnLabel => cell(rowLabel,columnLabel)) }));
  const max = Math.max(1,...matrix.flatMap(row => row.cells.map(item => item.value)));
  if (!rowLabels.length || !columnLabels.length) {
    $("plan-pivot").innerHTML = `<div class="empty-state">No response-plan data match the current filters.</div>`;
    return;
  }
  const columnTotals = columnLabels.map((_,index) => matrix.reduce((total,row) => ({ value: total.value + row.cells[index].value, reported: total.reported + row.cells[index].reported }), { value: 0, reported: 0 }));
  $("plan-pivot").innerHTML = `<table class="plan-pivot-table"><thead><tr><th>${label[rowField]} / ${label[columnField]}</th>${columnLabels.map(value => `<th>${escapeHtml(value)}</th>`).join("")}<th>Total</th></tr></thead><tbody>${matrix.map(row => {
    const rowTotal = row.cells.reduce((total,item) => ({ value: total.value + item.value, reported: total.reported + item.reported }), { value: 0, reported: 0 });
    return `<tr><th>${escapeHtml(row.rowLabel)}</th>${row.cells.map(item => `<td style="--heat:${item.reported ? Math.max(.08,item.value/max) : 0}"><strong>${item.reported ? format(item.value) : "—"}</strong></td>`).join("")}<td class="pivot-total"><strong>${rowTotal.reported ? format(rowTotal.value) : "—"}</strong></td></tr>`;
  }).join("")}</tbody><tfoot><tr><th>Total</th>${columnTotals.map(item => `<td><strong>${item.reported ? format(item.value) : "—"}</strong></td>`).join("")}<td><strong>${format(columnTotals.reduce((sum,item) => sum + item.value,0))}</strong></td></tr></tfoot></table>`;
}

function planFundingSummary(rows) {
  const reported = rows.filter(item => item.fundingRequired !== null);
  return reported.length ? fmtUSD.format(reported.reduce((sum,item) => sum + item.fundingRequired,0)) : "Not reported";
}

function planSexAge(item) {
  const values = [["Girls",item.girls],["Boys",item.boys],["Women",item.women],["Men",item.men],["Unknown",item.unknown]].filter(([,value]) => value !== null);
  return values.length ? `<div class="plan-sadd">${values.map(([label,value]) => `<span><b>${fmtNum.format(value)}</b>${label}</span>`).join("")}</div>` : `<span class="plan-not-reported">Not reported</span>`;
}

function renderPlanTable(rows) {
  return `<div class="plan-table-scroll"><table class="plan-record-table"><thead><tr><th>Indicator</th><th>Target</th><th>Sex and age</th><th>Funding required</th><th>Details</th></tr></thead><tbody>${rows.map(item => {
    const targetLabel = item.peopleTargeted !== null ? "People targeted" : (item.targetUnit || "Non-person target");
    const targetValue = item.peopleTargeted !== null ? item.peopleTargeted : item.nonPersonTarget;
    const detailsId = `plan-details-${item.id}`;
    return `<tr class="plan-record-row"><td><strong>${escapeHtml(item.indicator || "Indicator not reported")}</strong></td><td><b>${planMetric(targetValue)}</b><span>${escapeHtml(targetLabel)}</span></td><td>${planSexAge(item)}</td><td><b class="${item.fundingRequired === null ? "missing" : ""}">${item.fundingRequired === null ? "Not reported" : fmtUSD.format(item.fundingRequired)}</b></td><td><button class="plan-view-more" data-plan-toggle="${escapeHtml(detailsId)}" type="button" aria-expanded="false">View more</button></td></tr><tr class="plan-expanded-row hidden" id="${escapeHtml(detailsId)}"><td colspan="5"><div class="plan-expanded-content"><div class="plan-expanded-primary"><span>Priority activity</span><p>${escapeHtml(item.activity || "Not reported")}</p></div><div><span>Priority locations</span><p>${escapeHtml(item.locations || "Not reported")}</p></div><div><span>Priority groups</span><p>${escapeHtml(item.groups || "Not reported")}</p></div><div><span>People in need</span><p>${planMetric(item.peopleInNeed)}</p></div><div><span>National target</span><p>${planMetric(item.nationalTarget)}</p></div><div><span>Unit cost</span><p>${item.unitCost === null ? "Not reported" : fmtUSD.format(item.unitCost)}</p></div>${item.notes ? `<div class="plan-expanded-notes"><span>Notes</span><p>${escapeHtml(item.notes)}</p></div>` : ""}</div></td></tr>`;
  }).join("")}</tbody></table></div>`;
}

function renderPlanHierarchy(rows) {
  const primaryField = $("plan-group-by").value;
  const secondaryField = primaryField === "sector" ? "country" : "sector";
  const primaryLabel = primaryField === "sector" ? "UNICEF sector" : "Country";
  const secondaryLabel = secondaryField === "sector" ? "UNICEF sector" : "Country";
  const primaryGroups = [...new Set(rows.map(item => item[primaryField] || "Not reported"))].sort();
  return primaryGroups.map((primary,index) => {
    const primaryRows = rows.filter(item => (item[primaryField] || "Not reported") === primary);
    const secondaryGroups = [...new Set(primaryRows.map(item => item[secondaryField] || "Not reported"))].sort();
    const color = primaryField === "sector" ? sectorColor(primary) : "#008dc9";
    return `<details class="plan-hierarchy-primary" style="--group-color:${color}" ${index === 0 ? "open" : ""}><summary><div><span>${primaryLabel}</span><strong>${escapeHtml(primary)}</strong></div><div><b>${primaryRows.length} ${primaryRows.length === 1 ? "record" : "records"}</b><b>${planFundingSummary(primaryRows)}</b><i aria-hidden="true"></i></div></summary><div class="plan-hierarchy-children">${secondaryGroups.map(secondary => {
      const secondaryRows = primaryRows.filter(item => (item[secondaryField] || "Not reported") === secondary);
      const secondaryColor = secondaryField === "sector" ? sectorColor(secondary) : color;
      return `<details class="plan-hierarchy-secondary" style="--child-color:${secondaryColor}"><summary><div><span>${secondaryLabel}</span><strong>${escapeHtml(secondary)}</strong></div><div><b>${secondaryRows.length}</b><b>${planFundingSummary(secondaryRows)}</b><i aria-hidden="true"></i></div></summary>${renderPlanTable(secondaryRows)}</details>`;
    }).join("")}</div></details>`;
  }).join("");
}

function renderSectorPlan() {
  const rows = filteredSectorPlans();
  const totalFunding = rows.reduce((sum,item) => sum + (item.fundingRequired || 0),0);
  const peopleTargeted = rows.reduce((sum,item) => sum + (item.peopleTargeted || 0),0);
  const countries = new Set(rows.map(item => item.country).filter(Boolean));
  const fundedRows = rows.filter(item => item.fundingRequired !== null).length;
  $("sector-plan-status").textContent = `${sectorPlans.length} live plan rows · refreshed ${new Date().toLocaleTimeString("en-GB",{hour:"2-digit",minute:"2-digit"})}`;
  $("plan-record-count").textContent = `${rows.length} of ${sectorPlans.length} records`;
  $("plan-kpis").innerHTML = [
    ["Priority interventions",fmtNum.format(rows.length),`${countries.size} ${countries.size === 1 ? "country" : "countries"}`],
    ["Combined indicator targets",fmtNum.format(peopleTargeted),"Sum of UNICEF targets; people may overlap"],
    ["Funding required",fundedRows ? fmtUSD.format(totalFunding) : "Not reported",`${fundedRows} records with funding`],
    ["Funding data coverage",rows.length ? `${Math.round(fundedRows/rows.length*100)}%` : "0%",`${rows.length - fundedRows} records still missing a value`]
  ].map(([label,value,note],index) => `<article class="plan-kpi ${index === 2 ? "accent" : ""}"><span>${label}</span><strong>${value}</strong><small>${note}</small></article>`).join("");
  renderPlanBars("plan-sector-funding",rows,"sector");
  renderPlanBars("plan-country-funding",rows,"country");
  renderPlanPivot(rows);

  if (!rows.length) {
    $("sector-plan-cards").innerHTML = `<div class="empty-state">No response-plan records match these filters.</div>`;
    return;
  }
  $("sector-plan-cards").innerHTML = renderPlanHierarchy(rows);
}

function sitrepFilterRows() {
  return [...filteredActivities];
}

function uniqueNarratives(rows, field, limit = 2) {
  return [...new Set(rows.map(row => String(row[field] || "").trim()).filter(Boolean))].slice(0, limit);
}

function narrativeList(items, emptyText) {
  return items.length ? `<ul>${items.map(item => `<li>${escapeHtml(item)}</li>`).join("")}</ul>` : `<p class="sitrep-not-reported">${escapeHtml(emptyText)}</p>`;
}

function interventionList(rows) {
  const items = rows.map(row => {
    const title = String(row["Activity Title*"] || "Intervention not titled").trim();
    const description = String(row["What Was Done?*"] || "No intervention details reported.").trim();
    const activityDate = new Date(row["Activity Date*"]);
    const reportingDate = new Date(row["Submission Date*"]);
    const activityDateLabel = Number.isNaN(activityDate.getTime()) ? "Not reported" : fmtDate.format(activityDate);
    const reportingDateLabel = Number.isNaN(reportingDate.getTime()) ? "Not reported" : fmtDate.format(reportingDate);
    return `<li><strong>${escapeHtml(title)}</strong><span>${escapeHtml(description)}</span><small class="sitrep-activity-meta"><span><b>Activity date</b>${activityDateLabel}</span><span><b>Reporting date</b>${reportingDateLabel}</span><span><b>Status</b>${escapeHtml(row["Implementation Status*"] || "Not reported")}</span></small></li>`;
  });
  return `<ul class="sitrep-intervention-list">${items.join("")}</ul>`;
}

function renderSitrep() {
  const rows = sitrepFilterRows().sort((a,b) => new Date(b["Activity Date*"]) - new Date(a["Activity Date*"]));
  const selectedCountry = $("sitrep-country").value;
  const reportTitle = selectedCountry || "All countries";
  if (!rows.length) {
    $("sitrep-paper").innerHTML = `<div class="sitrep-empty"><strong>No reportable activities for ${escapeHtml(reportTitle)}</strong><p>Change the country or clear the sector, phase, status and search filters.</p></div>`;
    return;
  }

  const total = field => rows.reduce((sum,row) => sum + (Number(row[field]) || 0), 0);
  const dates = rows.map(row => new Date(row["Activity Date*"])).filter(date => !Number.isNaN(date.getTime()));
  const submissions = rows.map(row => new Date(row["Submission Date*"])).filter(date => !Number.isNaN(date.getTime()));
  const period = dates.length ? `${fmtDate.format(new Date(Math.min(...dates)))} – ${fmtDate.format(new Date(Math.max(...dates)))}` : "Reporting period not recorded";
  const latestSubmission = submissions.length ? fmtDate.format(new Date(Math.max(...submissions))) : "Not recorded";
  const active = rows.filter(row => row["Implementation Status*"] !== "Completed").length;
  const evidence = rows.filter(row => /^https?:\/\//.test(row["Evidence Link"])).length;
  const countryNames = selectedCountry ? [selectedCountry] : [...new Set(rows.map(row => row["Country*"] || "Country not reported"))].sort();

  const renderSector = (sectorRows, sector, countryName) => {
    const challenges = uniqueNarratives(sectorRows,"Challenges",2);
    const nextSteps = uniqueNarratives(sectorRows,"Next Step",2);
    const partners = [...new Set(sectorRows.flatMap(row => String(row["Partners"] || "").split(/[,;]/)).map(value => value.trim()).filter(Boolean))];
    const sectorPeople = sectorRows.reduce((sum,row) => sum + (Number(row["People Reached (Total)"]) || 0),0);
    const sectorChildren = sectorRows.reduce((sum,row) => sum + (Number(row["Children Reached"]) || 0),0);
    const sectorReachMode = reachMode(sectorRows);
    return `<section class="sitrep-sector">
      <header class="sitrep-sector-head"><div><span>${escapeHtml(countryName)} · UNICEF SECTOR</span><h3>${escapeHtml(sector)}</h3></div><p>${sectorRows.length} ${sectorRows.length === 1 ? "activity" : "activities"}</p></header>
      <div class="sitrep-sector-grid">
        <div class="sitrep-column sitrep-interventions"><h4>UNICEF interventions</h4>${interventionList(sectorRows)}</div>
        <div class="sitrep-column sitrep-partners"><h4>Partners</h4>${partners.length ? `<ul>${partners.map(partner => `<li>${escapeHtml(partner)}</li>`).join("")}</ul>` : `<p class="sitrep-not-reported">No partners reported.</p>`}</div>
        <div class="sitrep-column sitrep-reach"><h4>People & children reached / targeted</h4><div class="sitrep-reach-metric"><strong class="sitrep-reach-number">${fmtNum.format(sectorPeople)}</strong><span class="sitrep-reach-label">People ${sectorReachMode}</span></div><div class="sitrep-reach-metric"><strong class="sitrep-reach-number">${fmtNum.format(sectorChildren)}</strong><span class="sitrep-reach-label">Children ${sectorReachMode}</span></div></div>
        <div class="sitrep-column sitrep-gaps"><h4>Gaps & priority actions</h4>${narrativeList(challenges,"No gaps reported.")}<h5>Priority actions</h5>${narrativeList(nextSteps,"No next step reported.")}</div>
      </div>
    </section>`;
  };

  const countrySections = countryNames.map(countryName => {
    const countryRows = rows.filter(row => (row["Country*"] || "Country not reported") === countryName);
    const sectorNames = [...new Set(countryRows.map(row => row["UNICEF Unit*"] || "Sector not reported"))].sort();
    const sectors = sectorNames.map(sector => renderSector(
      countryRows.filter(row => (row["UNICEF Unit*"] || "Sector not reported") === sector),
      sector,
      countryName
    )).join("");
    const countryPeople = countryRows.reduce((sum,row) => sum + (Number(row["People Reached (Total)"]) || 0),0);
    const countryChildren = countryRows.reduce((sum,row) => sum + (Number(row["Children Reached"]) || 0),0);
    const countryFunding = countryRows.reduce((sum,row) => sum + (Number(row["Funding Used (USD)"]) || 0),0);
    const countryReachMode = reachMode(countryRows);
    const countryHeading = selectedCountry ? "" : `<header class="sitrep-country-head"><div><span>COUNTRY</span><h3>${escapeHtml(countryName)}</h3></div><div class="sitrep-country-stats"><span><strong>${fmtNum.format(countryRows.length)}</strong> activities</span><span><strong>${fmtNum.format(countryPeople)}</strong> people ${countryReachMode}</span><span><strong>${fmtNum.format(countryChildren)}</strong> children ${countryReachMode}</span><span><strong>${fmtUSD.format(countryFunding)}</strong> ${fundingMode(countryRows)}</span></div></header>`;
    return `<section class="sitrep-country-group">${countryHeading}<div class="sitrep-country-sectors">${sectors}</div></section>`;
  }).join("");

  const reportScope = selectedCountry ? `Country: ${selectedCountry}` : `Regional overview · ${countryNames.length} ${countryNames.length === 1 ? "country" : "countries"}`;
  $("sitrep-paper").innerHTML = `<header class="sitrep-masthead">
      <div><p class="sitrep-kicker">UNICEF PACIFIC · HUMANITARIAN REPORTING</p><h2>El Niño Situation Report</h2><p class="sitrep-period"><strong>${escapeHtml(reportScope)}</strong> · Reporting period: ${period}</p></div>
      <div class="sitrep-mark"><img src="https://upload.wikimedia.org/wikipedia/commons/e/ed/Logo_of_UNICEF.svg" alt="UNICEF"><span>Generated ${fmtDate.format(new Date())}</span></div>
    </header>
    <section class="sitrep-figures" aria-label="Key figures">
      <div><strong>${fmtNum.format(rows.length)}</strong><span>reported activities</span></div>
      <div><strong>${fmtNum.format(total("People Reached (Total)"))}</strong><span>people ${reachMode(rows)}</span></div>
      <div><strong>${fmtNum.format(total("Children Reached"))}</strong><span>children ${reachMode(rows)}</span></div>
      <div><strong>${fmtUSD.format(total("Funding Used (USD)"))}</strong><span>funding ${fundingMode(rows)}</span></div>
      <div><strong>${fmtNum.format(active)}</strong><span>active / pending</span></div>
    </section>
    <div class="sitrep-section-title response-title"><span>01</span><h3>${selectedCountry ? "UNICEF response by sector" : "UNICEF response by country and sector"}</h3></div>
    <div class="sitrep-country-list">${countrySections}</div>
    <footer class="sitrep-foot"><p><strong>Reporting note:</strong> UNICEF interventions use Activity Title and What Was Done; partners use Partners; reach uses People Reached (Total) and is labelled from Implementation Status; gaps and priority actions use Challenges and Next Step. Validate before external circulation.</p><p><strong>Data assurance:</strong> Latest submission ${latestSubmission} · ${evidence}/${rows.length} records include evidence links · Source: ${escapeHtml(metadata.source || (metadata.isDemo ? "Demonstration data" : "Activity Diary"))}</p></footer>`;
}

function detailItem(label, value, full = false, format) {
  const shown = format ? format(value) : escapeHtml(value || value === 0 ? value : "Not reported");
  return `<div class="detail-item ${full ? "full" : ""}"><span>${label}</span><p>${shown}</p></div>`;
}
function openDetail(entryId) {
  const key = String(entryId ?? "").trim();
  const d = allActivities.find(row => String(row["Entry ID"] ?? "").trim() === key); if (!d) { showToast("This activity record could not be opened."); return; }
  const evidence = metadata.isDemo ? escapeHtml(d["Evidence Link"] || "Not reported") + " (illustrative link)" : (/^https?:\/\//.test(d["Evidence Link"]) ? `<a class="detail-link" href="${escapeHtml(d["Evidence Link"])}" target="_blank" rel="noopener">Open supporting evidence ↗</a>` : "Not reported");
  $("dialog-content").innerHTML = `<header class="dialog-title"><p class="section-kicker">${escapeHtml(d["Entry ID"])}</p><h2>${escapeHtml(d["Activity Title*"])}</h2><p>${formatDateValue(d["Activity Date*"])} · ${escapeHtml(d["UNICEF Unit*"])} · ${escapeHtml(d["Country*"])}</p></header><div class="detail-groups">
    <section class="detail-group"><h3>System & classification</h3><div class="detail-grid">${detailItem("Entry ID",d["Entry ID"])}${detailItem("Activity date",d["Activity Date*"],false,formatDateValue)}${detailItem("Reporting month",d["Reporting Month"])}${detailItem("UNICEF sector",d["UNICEF Unit*"])}${detailItem("Country",d["Country*"])}${detailItem("Location / admin area",d["Location / Admin Area"])}${detailItem("El Niño phase",d["El Niño Phase*"])}${detailItem("Activity type",d["Activity Type*"])}${detailItem("Implementation status",d["Implementation Status*"])}</div></section>
    <section class="detail-group"><h3>Results & delivery</h3><div class="detail-grid">${detailItem(`People ${reachMode([d])}`,d["People Reached (Total)"],false,v=>fmtNum.format(v))}${detailItem(`Children ${reachMode([d])}`,d["Children Reached"],false,v=>fmtNum.format(v))}${detailItem(`Funding ${fundingMode([d])}`,d["Funding Used (USD)"],false,v=>fmtUSD.format(v))}${detailItem("Partners",d["Partners"],true)}${detailItem("Result / output",d["Result / Output"],true)}</div></section>
    <section class="detail-group full"><h3>Activity narrative</h3><div class="detail-grid">${detailItem("Activity title",d["Activity Title*"],true)}${detailItem("What was done?",d["What Was Done?*"],true)}</div></section>
    <section class="detail-group full"><h3>Follow-up & evidence</h3><div class="detail-grid">${detailItem("Challenges",d["Challenges"],true)}${detailItem("Next step",d["Next Step"],true)}${detailItem("Evidence link",evidence,true,v=>v)}${detailItem("Focal point",d["Focal Point*"])}${detailItem("Submission date",d["Submission Date*"],false,formatDateValue)}</div></section>
  </div>`;
  $("activity-dialog").showModal();
}

function renderAll() {
  renderKpis(); renderMonthly(); renderPhase(); renderRankBars("unit-bars",group("UNICEF Unit*")); renderRankBars("reach-bars",group("UNICEF Unit*","People Reached (Total)")); renderFollowup(); renderLatest(); renderTimeline(); renderRegister(); renderMap(); renderKeyMessages(); renderSectorPlan(); renderSitrep();
}
function switchView(name) {
  document.querySelectorAll(".tab").forEach(t => t.classList.toggle("active",t.dataset.view === name));
  document.querySelectorAll(".view").forEach(v => v.classList.toggle("active",v.id === `view-${name}`));
  $("activity-filters").classList.toggle("hidden",name === "sector-plan");
  if (name === "map") { initMap(); setTimeout(() => { map.invalidateSize(); renderMap(); },50); }
  window.scrollTo({ top: 0, behavior: "smooth" });
}
function showToast(message) { const toast = $("toast"); toast.textContent = message; toast.classList.add("show"); clearTimeout(showToast.timer); showToast.timer = setTimeout(()=>toast.classList.remove("show"),3000); }

document.querySelectorAll(".tab").forEach(tab => tab.addEventListener("click",()=>switchView(tab.dataset.view)));
document.querySelectorAll("[data-go-view]").forEach(button => button.addEventListener("click",()=>switchView(button.dataset.goView)));
document.querySelectorAll("[data-activity-layout]").forEach(button => button.addEventListener("click",()=>{
  const layout = button.dataset.activityLayout;
  document.querySelectorAll("[data-activity-layout]").forEach(item => item.classList.toggle("active",item === button));
  $("activity-timeline-layout").classList.toggle("active",layout === "timeline");
  $("activity-register-layout").classList.toggle("active",layout === "register");
}));
["search-filter","unit-filter","country-filter","phase-filter","status-filter","activity-start-filter","activity-end-filter"].forEach(id => $(id).addEventListener(id === "search-filter" ? "input" : "change",applyFilters));
['plan-sector-filter','plan-country-filter','plan-indicator-filter','plan-funding-filter'].forEach(id => $(id).addEventListener('change',renderSectorPlan));
['plan-pivot-row','plan-pivot-column','plan-pivot-value'].forEach(id => $(id).addEventListener('change',()=>{
  if ($("plan-pivot-row").value === $("plan-pivot-column").value) {
    $("plan-pivot-column").value = $("plan-pivot-row").value === "country" ? "sector" : "country";
  }
  renderSectorPlan();
}));
$("plan-group-by").addEventListener("change",renderSectorPlan);
$("plan-clear-filters").addEventListener("click",()=>{ ['plan-sector-filter','plan-country-filter','plan-indicator-filter','plan-funding-filter'].forEach(id => $(id).value = ""); renderSectorPlan(); });
$("clear-filters").addEventListener("click",()=>{ ["search-filter","unit-filter","country-filter","phase-filter","status-filter","activity-start-filter","activity-end-filter"].forEach(id=>$(id).value=""); applyFilters(); });
$("refresh-button").addEventListener("click",async()=>{
  await refreshData();
});
$("sitrep-country").addEventListener("change",()=>{
  $("country-filter").value = $("sitrep-country").value;
  applyFilters();
});
$("sitrep-print").addEventListener("click",()=>window.print());
$("dialog-close").addEventListener("click",()=>$("activity-dialog").close());
$("activity-dialog").addEventListener("click",e=>{ if(e.target === $("activity-dialog")) $("activity-dialog").close(); });
document.body.addEventListener("click",e=>{
  const planToggle = e.target.closest("[data-plan-toggle]");
  if (planToggle) {
    const details = $(planToggle.dataset.planToggle);
    if (details) {
      const expanded = planToggle.getAttribute("aria-expanded") === "true";
      planToggle.setAttribute("aria-expanded",String(!expanded));
      planToggle.textContent = expanded ? "View more" : "Show less";
      details.classList.toggle("hidden",expanded);
    }
    return;
  }
  const countryTarget = e.target.closest("[data-map-country]");
  if (countryTarget) {
    $("country-filter").value = countryTarget.dataset.mapCountry;
    applyFilters();
    return;
  }
  const target=e.target.closest("[data-entry]");
  if(target) openDetail(target.dataset.entry);
});

refreshData();
setInterval(()=>refreshData(),Math.max(1,Number(config.refreshMinutes)||5)*60000);
