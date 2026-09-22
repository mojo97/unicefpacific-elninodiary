const FIELD_ORDER = [
  "Entry ID", "Activity Date*", "Reporting Month", "UNICEF Unit*", "Country*",
  "Location / Admin Area", "El Niño Phase*", "Activity Type*", "Activity Title*",
  "What Was Done?*", "Result / Output", "People Reached (Total)", "Children Reached",
  "Partners", "Implementation Status*", "Funding Used (USD)", "Challenges", "Next Step",
  "Evidence Link", "Focal Point*", "Submission Date*"
];
const FIELD_ALIASES = {
  "UNICEF Unit*": ["UNICEF Sector*", "UNICEF Sectors*", "UNICEF Unit / Sector*"]
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
const config = window.DIARY_CONFIG || { mode: "local", refreshMinutes: 5 };
let allActivities = [], filteredActivities = [], metadata = { isDemo: true }, map, markerLayer;

const $ = (id) => document.getElementById(id);
const fmtNum = new Intl.NumberFormat("en-US", { maximumFractionDigits: 0 });
const fmtUSD = new Intl.NumberFormat("en-US", { style: "currency", currency: "USD", maximumFractionDigits: 0 });
const fmtDate = new Intl.DateTimeFormat("en-GB", { day: "numeric", month: "short", year: "numeric" });
const escapeHtml = (value) => String(value ?? "—").replace(/[&<>'"]/g, c => ({"&":"&amp;","<":"&lt;",">":"&gt;","'":"&#39;",'"':"&quot;"}[c]));
const phaseColor = (phase) => PHASE_COLORS[phase] || "#91a6af";
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

function parseExcelData(bytes, sourceDetails) {
  if (!window.XLSX) throw new Error("The Excel reader did not load. Check your internet connection and try again.");
  const workbook = XLSX.read(bytes, { type: "array", cellDates: true });
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
  return parseExcelData(await response.arrayBuffer(), { source: "Google Sheets", fileName: config.googleSheets?.displayName || "El Niño Activity Diary" });
}

async function refreshData() {
  $("refresh-button").disabled = true;
  $("refresh-button").innerHTML = "<span>↻</span> Loading";
  try {
    allActivities = config.mode === "google-sheets" ? await loadGoogleSheetsData() : await loadLocalData();
    allActivities.sort((a,b) => new Date(a["Activity Date*"]) - new Date(b["Activity Date*"]));
    $("demo-banner").classList.toggle("hidden", !metadata.isDemo);
    $("sync-state").classList.toggle("live", !metadata.isDemo);
    const sourceLabel = metadata.isDemo ? "Demonstration data" : "Live Google Sheet";
    $("sync-state").querySelector("span").textContent = sourceLabel;
    $("sync-state").title = sourceLabel;
    $("footer-source").textContent = metadata.isDemo ? "workbook example entries" : "live Google Sheet";
    populateFilters(); applyFilters();
    $("last-refreshed").textContent = new Date().toLocaleString("en-GB", { dateStyle: "medium", timeStyle: "short" });
    showToast(`${allActivities.length} activities loaded.`);
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
}

function applyFilters() {
  const search = $("search-filter").value.trim().toLowerCase();
  const unit = $("unit-filter").value, country = $("country-filter").value, phase = $("phase-filter").value, status = $("status-filter").value;
  filteredActivities = allActivities.filter(d => {
    const haystack = FIELD_ORDER.map(field => d[field]).join(" ").toLowerCase();
    return (!search || haystack.includes(search)) && (!unit || d["UNICEF Unit*"] === unit) && (!country || d["Country*"] === country) && (!phase || d["El Niño Phase*"] === phase) && (!status || d["Implementation Status*"] === status);
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
function renderMonthly() {
  const months = group("Reporting Month").sort((a,b) => new Date(`1 ${a[0]}`) - new Date(`1 ${b[0]}`));
  const max = Math.max(1,...months.map(d => d[1]));
  $("monthly-chart").innerHTML = months.length ? months.map(([month,value]) => `<div class="month-column"><span class="month-value">${value}</span><div class="month-bar" style="height:${Math.max(8,value/max*145)}px"></div><span class="month-label">${escapeHtml(month)}</span></div>`).join("") : `<div class="empty-state">No activities match these filters.</div>`;
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
  const lags = filteredActivities.map(d => (new Date(d["Submission Date*"]) - new Date(d["Activity Date*"])) / 86400000).filter(Number.isFinite);
  const avgLag = lags.length ? Math.round(lags.reduce((a,b)=>a+b,0)/lags.length) : 0;
  const items = [["↗",`${ongoing} still active`,"Ongoing, planned or on-hold activities"],["!",`${challenges} challenges logged`,"Constraints needing management attention"],["→",`${nextSteps} next steps`,"Forward actions documented"],["✓",`${evidence}/${filteredActivities.length} evidence links`,"Records with a supporting URL"],["◷",`${avgLag} day submission lag`,"Average time from activity to submission"]];
  $("followup-list").innerHTML = items.map(([icon,title,text]) => `<div class="followup-item"><span class="followup-icon">${icon}</span><div><strong>${title}</strong><p>${text}</p></div></div>`).join("");
}

function statusClass(status) { return `status-${String(status || "").toLowerCase().replace(/\s+/g,"-")}`; }
function renderLatest() {
  const latest = [...filteredActivities].sort((a,b)=>new Date(b["Activity Date*"])-new Date(a["Activity Date*"])).slice(0,3);
  $("latest-activities").innerHTML = latest.length ? latest.map(d => `<button class="activity-card" data-entry="${escapeHtml(d["Entry ID"])}" type="button"><div class="meta"><span>${fmtDate.format(new Date(d["Activity Date*"]))} · ${escapeHtml(d["UNICEF Unit*"])}</span><span class="status-pill ${statusClass(d["Implementation Status*"])}">${escapeHtml(d["Implementation Status*"])}</span></div><h3>${escapeHtml(d["Activity Title*"])}</h3><p>${escapeHtml(d["Country*"])} · ${escapeHtml(d["Result / Output"] || "Result not reported")}</p></button>`).join("") : `<div class="empty-state">No activities match these filters.</div>`;
}
function renderTimeline() {
  const rows = [...filteredActivities].sort((a,b)=>new Date(b["Activity Date*"])-new Date(a["Activity Date*"]));
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
    const marker = L.circleMarker([d.latitude,lng], { radius: 9, color: "#fff", weight: 2, fillColor: phaseColor(d["El Niño Phase*"]), fillOpacity: .95 });
    marker.bindPopup(`<h3>${escapeHtml(d["Activity Title*"])}</h3><p><strong>${escapeHtml(d["Country*"])}</strong> · ${escapeHtml(d["Location / Admin Area"])}</p><p>${escapeHtml(d["UNICEF Unit*"])} · ${escapeHtml(d["Implementation Status*"])}</p><p>${fmtDate.format(new Date(d["Activity Date*"]))}</p>`);
    marker.addTo(markerLayer); bounds.push([d.latitude,lng]);
  });
  if (bounds.length) map.fitBounds(bounds,{padding:[40,40],maxZoom:5});
  const countries = group("Country*");
  $("map-count").textContent = `${countries.length} mapped ${countries.length === 1 ? "location" : "locations"}`;
  $("map-country-list").innerHTML = countries.map(([name,count]) => `<div class="country-row"><span>${escapeHtml(name)}</span><strong>${count}</strong></div>`).join("");
}

function renderRegister() {
  $("register-body").innerHTML = filteredActivities.length ? [...filteredActivities].sort((a,b)=>new Date(b["Activity Date*"])-new Date(a["Activity Date*"])).map(d => `<tr><td>${fmtDate.format(new Date(d["Activity Date*"]))}</td><td><strong>${escapeHtml(d["UNICEF Unit*"])}</strong></td><td>${escapeHtml(d["Country*"])}<br><small>${escapeHtml(d["Location / Admin Area"])}</small></td><td>${escapeHtml(d["Activity Title*"])}</td><td>${escapeHtml(d["El Niño Phase*"])}</td><td><span class="status-pill ${statusClass(d["Implementation Status*"])}">${escapeHtml(d["Implementation Status*"])}</span></td><td>${fmtNum.format(d["People Reached (Total)"])}</td><td>${fmtUSD.format(d["Funding Used (USD)"])}</td><td><button class="row-button" data-entry="${escapeHtml(d["Entry ID"])}" type="button">View →</button></td></tr>`).join("") : `<tr><td colspan="9" class="empty-state">No activities match these filters.</td></tr>`;
}

function sitrepFilterRows() {
  const search = $("search-filter").value.trim().toLowerCase();
  const unit = $("unit-filter").value, phase = $("phase-filter").value, status = $("status-filter").value;
  const country = $("sitrep-country").value;
  return allActivities.filter(d => {
    const haystack = FIELD_ORDER.map(field => d[field]).join(" ").toLowerCase();
    return (!country || d["Country*"] === country) && (!search || haystack.includes(search)) && (!unit || d["UNICEF Unit*"] === unit) && (!phase || d["El Niño Phase*"] === phase) && (!status || d["Implementation Status*"] === status);
  });
}

function uniqueNarratives(rows, field, limit = 2) {
  return [...new Set(rows.map(row => String(row[field] || "").trim()).filter(Boolean))].slice(0, limit);
}

function narrativeList(items, emptyText) {
  return items.length ? `<ul>${items.map(item => `<li>${escapeHtml(item)}</li>`).join("")}</ul>` : `<p class="sitrep-not-reported">${escapeHtml(emptyText)}</p>`;
}

function interventionList(rows, limit = 3) {
  const items = rows.slice(0, limit).map(row => {
    const title = String(row["Activity Title*"] || "Intervention not titled").trim();
    const description = String(row["What Was Done?*"] || "No intervention details reported.").trim();
    const date = new Date(row["Activity Date*"]);
    const dateLabel = Number.isNaN(date.getTime()) ? "Date not reported" : fmtDate.format(date);
    return `<li><strong>${escapeHtml(title)}</strong><span>${escapeHtml(description)}</span><small>${dateLabel} · ${escapeHtml(row["Implementation Status*"] || "Status not reported")}</small></li>`;
  });
  const remainder = rows.length - items.length;
  return `<ul class="sitrep-intervention-list">${items.join("")}</ul>${remainder > 0 ? `<p class="sitrep-more">+${fmtNum.format(remainder)} additional ${remainder === 1 ? "activity" : "activities"} included in totals.</p>` : ""}`;
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
  const d = allActivities.find(row => row["Entry ID"] === entryId); if (!d) return;
  const evidence = metadata.isDemo ? escapeHtml(d["Evidence Link"] || "Not reported") + " (illustrative link)" : (/^https?:\/\//.test(d["Evidence Link"]) ? `<a class="detail-link" href="${escapeHtml(d["Evidence Link"])}" target="_blank" rel="noopener">Open supporting evidence ↗</a>` : "Not reported");
  $("dialog-content").innerHTML = `<header class="dialog-title"><p class="section-kicker">${escapeHtml(d["Entry ID"])}</p><h2>${escapeHtml(d["Activity Title*"])}</h2><p>${fmtDate.format(new Date(d["Activity Date*"]))} · ${escapeHtml(d["UNICEF Unit*"])} · ${escapeHtml(d["Country*"])}</p></header><div class="detail-groups">
    <section class="detail-group"><h3>System & classification</h3><div class="detail-grid">${detailItem("Entry ID",d["Entry ID"])}${detailItem("Activity date",d["Activity Date*"],false,v=>fmtDate.format(new Date(v)))}${detailItem("Reporting month",d["Reporting Month"])}${detailItem("UNICEF sector",d["UNICEF Unit*"])}${detailItem("Country",d["Country*"])}${detailItem("Location / admin area",d["Location / Admin Area"])}${detailItem("El Niño phase",d["El Niño Phase*"])}${detailItem("Activity type",d["Activity Type*"])}${detailItem("Implementation status",d["Implementation Status*"])}</div></section>
    <section class="detail-group"><h3>Results & delivery</h3><div class="detail-grid">${detailItem(`People ${reachMode([d])}`,d["People Reached (Total)"],false,v=>fmtNum.format(v))}${detailItem(`Children ${reachMode([d])}`,d["Children Reached"],false,v=>fmtNum.format(v))}${detailItem(`Funding ${fundingMode([d])}`,d["Funding Used (USD)"],false,v=>fmtUSD.format(v))}${detailItem("Partners",d["Partners"],true)}${detailItem("Result / output",d["Result / Output"],true)}</div></section>
    <section class="detail-group full"><h3>Activity narrative</h3><div class="detail-grid">${detailItem("Activity title",d["Activity Title*"],true)}${detailItem("What was done?",d["What Was Done?*"],true)}</div></section>
    <section class="detail-group full"><h3>Follow-up & evidence</h3><div class="detail-grid">${detailItem("Challenges",d["Challenges"],true)}${detailItem("Next step",d["Next Step"],true)}${detailItem("Evidence link",evidence,true,v=>v)}${detailItem("Focal point",d["Focal Point*"])}${detailItem("Submission date",d["Submission Date*"],false,v=>fmtDate.format(new Date(v)))}</div></section>
  </div>`;
  $("activity-dialog").showModal();
}

function renderAll() {
  renderKpis(); renderMonthly(); renderPhase(); renderRankBars("unit-bars",group("UNICEF Unit*")); renderRankBars("reach-bars",group("UNICEF Unit*","People Reached (Total)")); renderFollowup(); renderLatest(); renderTimeline(); renderRegister(); renderMap(); renderSitrep();
}
function switchView(name) {
  document.querySelectorAll(".tab").forEach(t => t.classList.toggle("active",t.dataset.view === name));
  document.querySelectorAll(".view").forEach(v => v.classList.toggle("active",v.id === `view-${name}`));
  if (name === "map") { initMap(); setTimeout(() => { map.invalidateSize(); renderMap(); },50); }
  window.scrollTo({ top: 0, behavior: "smooth" });
}
function showToast(message) { const toast = $("toast"); toast.textContent = message; toast.classList.add("show"); clearTimeout(showToast.timer); showToast.timer = setTimeout(()=>toast.classList.remove("show"),3000); }

document.querySelectorAll(".tab").forEach(tab => tab.addEventListener("click",()=>switchView(tab.dataset.view)));
document.querySelectorAll("[data-go-view]").forEach(button => button.addEventListener("click",()=>switchView(button.dataset.goView)));
["search-filter","unit-filter","country-filter","phase-filter","status-filter"].forEach(id => $(id).addEventListener(id === "search-filter" ? "input" : "change",applyFilters));
$("clear-filters").addEventListener("click",()=>{ ["search-filter","unit-filter","country-filter","phase-filter","status-filter"].forEach(id=>$(id).value=""); applyFilters(); });
$("refresh-button").addEventListener("click",async()=>{
  await refreshData();
});
$("sitrep-country").addEventListener("change",renderSitrep);
$("sitrep-print").addEventListener("click",()=>window.print());
$("dialog-close").addEventListener("click",()=>$("activity-dialog").close());
$("activity-dialog").addEventListener("click",e=>{ if(e.target === $("activity-dialog")) $("activity-dialog").close(); });
document.body.addEventListener("click",e=>{ const target=e.target.closest("[data-entry]"); if(target) openDetail(target.dataset.entry); });

refreshData();
setInterval(()=>refreshData(),Math.max(1,Number(config.refreshMinutes)||5)*60000);
