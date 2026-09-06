const FIELD_ORDER = [
  "Entry ID", "Activity Date*", "Reporting Month", "UNICEF Sector*", "Country*",
  "Location / Admin Area", "El Niño Phase*", "Activity Type*", "Activity Title*",
  "What Was Done?*", "Result / Output", "People Reached (Total)", "Children Reached",
  "Partners", "Implementation Status*", "Funding Used (USD)", "Challenges", "Next Step",
  "Evidence Link", "Focal Point*", "Submission Date*"
];

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
let allActivities = [], filteredActivities = [], metadata = { isDemo: true }, map, markerLayer, msalClient, uploadedWorkbookFile;

const $ = (id) => document.getElementById(id);
const fmtNum = new Intl.NumberFormat("en-US", { maximumFractionDigits: 0 });
const fmtUSD = new Intl.NumberFormat("en-US", { style: "currency", currency: "USD", maximumFractionDigits: 0 });
const fmtDate = new Intl.DateTimeFormat("en-GB", { day: "numeric", month: "short", year: "numeric" });
const escapeHtml = (value) => String(value ?? "—").replace(/[&<>'"]/g, c => ({"&":"&amp;","<":"&lt;",">":"&gt;","'":"&#39;",'"':"&quot;"}[c]));
const phaseColor = (phase) => PHASE_COLORS[phase] || "#91a6af";

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

async function loadExcelData(file) {
  if (!window.XLSX) throw new Error("The Excel reader did not load. Check your internet connection and try again.");
  const bytes = await file.arrayBuffer();
  const workbook = XLSX.read(bytes, { type: "array", cellDates: true });
  const preferredSheet = workbook.Sheets["Activity Diary"] ? "Activity Diary" : null;
  const candidates = preferredSheet ? [preferredSheet] : workbook.SheetNames;
  let selected;

  for (const sheetName of candidates) {
    const rows = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { header: 1, defval: "", raw: true });
    const headerIndex = rows.findIndex(row => {
      const labels = row.map(value => String(value ?? "").trim());
      return labels.includes("Entry ID") && labels.includes("Activity Date*") && labels.includes("Activity Title*");
    });
    if (headerIndex >= 0) { selected = { sheetName, rows, headerIndex }; break; }
  }

  if (!selected) throw new Error("No Activity Diary header row was found in this Excel file.");
  const headers = selected.rows[selected.headerIndex].map(value => String(value ?? "").trim());
  const missing = FIELD_ORDER.filter(field => !headers.includes(field));
  if (missing.length) throw new Error(`The Excel file is missing required column${missing.length === 1 ? "" : "s"}: ${missing.join(", ")}`);

  const activities = selected.rows.slice(selected.headerIndex + 1).map(row => Object.fromEntries(headers.map((header,index) => [header,row[index]])))
    .filter(row => row["Activity Date*"] !== "" && row["Activity Date*"] !== null && row["Activity Title*"] !== "")
    .map(normalizeActivity);
  if (!activities.length) throw new Error(`The ${selected.sheetName} sheet contains no completed activity rows.`);

  metadata = { isDemo: false, source: "Uploaded Excel", fileName: file.name, sheet: selected.sheetName, generatedAt: new Date().toISOString() };
  return activities;
}

async function loadLocalData() {
  const response = await fetch("data/activities.json", { cache: "no-store" });
  if (!response.ok) throw new Error("The demonstration data file could not be loaded.");
  const payload = await response.json();
  metadata = payload.metadata || { isDemo: true };
  return payload.activities.map(normalizeActivity);
}

async function initialiseMicrosoft() {
  if (!window.msal) throw new Error("Microsoft sign-in library did not load.");
  const sp = config.sharePoint || {};
  const missing = ["tenantId", "clientId", "siteId", "driveId", "itemId"].some(k => !sp[k] || String(sp[k]).startsWith("YOUR_"));
  if (missing) throw new Error("SharePoint live mode is not configured yet. Complete config.js first.");
  if (!msalClient) {
    msalClient = new msal.PublicClientApplication({
      auth: { clientId: sp.clientId, authority: `https://login.microsoftonline.com/${sp.tenantId}`, redirectUri: window.location.origin + window.location.pathname },
      cache: { cacheLocation: "sessionStorage", storeAuthStateInCookie: false }
    });
    if (msalClient.initialize) await msalClient.initialize();
  }
}

async function loadGraphData(interactive = false) {
  await initialiseMicrosoft();
  let account = msalClient.getAllAccounts()[0];
  const scopes = ["Files.Read.All", "Sites.Read.All"];
  if (!account && interactive) {
    const login = await msalClient.loginPopup({ scopes, prompt: "select_account" });
    account = login.account;
  }
  if (!account) throw new Error("Select Connect Microsoft 365 to load the private workbook.");
  let token;
  try { token = await msalClient.acquireTokenSilent({ scopes, account }); }
  catch (error) {
    if (!interactive) throw error;
    token = await msalClient.acquireTokenPopup({ scopes, account });
  }
  const sp = config.sharePoint;
  const url = `https://graph.microsoft.com/v1.0/sites/${encodeURIComponent(sp.siteId)}/drives/${encodeURIComponent(sp.driveId)}/items/${encodeURIComponent(sp.itemId)}/workbook/tables/${encodeURIComponent(sp.tableName)}/rows`;
  const response = await fetch(url, { headers: { Authorization: `Bearer ${token.accessToken}` }, cache: "no-store" });
  if (!response.ok) throw new Error(`Microsoft Graph returned ${response.status}. Check file permissions and identifiers.`);
  const payload = await response.json();
  const rows = (payload.value || []).map(row => row.values?.[0] || []).filter(row => row[1] !== null && row[1] !== "");
  metadata = { isDemo: false, source: "SharePoint", table: sp.tableName, generatedAt: new Date().toISOString() };
  return rows.map(row => normalizeActivity(Object.fromEntries(FIELD_ORDER.map((field, index) => [field, row[index]]))));
}

async function refreshData(interactive = false) {
  $("refresh-button").disabled = true;
  $("refresh-button").innerHTML = "<span>↻</span> Loading";
  try {
    allActivities = uploadedWorkbookFile ? await loadExcelData(uploadedWorkbookFile) : (config.mode === "graph" ? await loadGraphData(interactive) : await loadLocalData());
    allActivities.sort((a,b) => new Date(a["Activity Date*"]) - new Date(b["Activity Date*"]));
    $("connect-button").classList.toggle("hidden", config.mode !== "graph");
    $("demo-banner").classList.toggle("hidden", !metadata.isDemo);
    $("sync-state").classList.toggle("live", !metadata.isDemo);
    const sourceLabel = metadata.source === "Uploaded Excel" ? `Excel · ${metadata.fileName}` : (metadata.isDemo ? "Demonstration data" : "Live from SharePoint");
    $("sync-state").querySelector("span").textContent = sourceLabel;
    $("sync-state").title = sourceLabel;
    $("footer-source").textContent = metadata.source === "Uploaded Excel" ? `${metadata.fileName} · ${metadata.sheet}` : (metadata.isDemo ? "workbook example entries" : "live SharePoint workbook");
    populateFilters(); applyFilters();
    $("last-refreshed").textContent = new Date().toLocaleString("en-GB", { dateStyle: "medium", timeStyle: "short" });
    showToast(`${allActivities.length} activities loaded.`);
    return true;
  } catch (error) {
    showToast(error.message || "Data refresh failed.");
    if (!allActivities.length && config.mode === "graph") {
      allActivities = await loadLocalData(); populateFilters(); applyFilters();
      $("connect-button").classList.remove("hidden");
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
  populateSelect("Sector-filter", unique("UNICEF Sector*"), "Sectors");
  populateSelect("country-filter", unique("Country*"), "countries");
  populateSelect("phase-filter", unique("El Niño Phase*"), "phases");
  populateSelect("status-filter", unique("Implementation Status*"), "statuses");
}

function applyFilters() {
  const search = $("search-filter").value.trim().toLowerCase();
  const Sector = $("Sector-filter").value, country = $("country-filter").value, phase = $("phase-filter").value, status = $("status-filter").value;
  filteredActivities = allActivities.filter(d => {
    const haystack = FIELD_ORDER.map(field => d[field]).join(" ").toLowerCase();
    return (!search || haystack.includes(search)) && (!Sector || d["UNICEF Sector*"] === Sector) && (!country || d["Country*"] === country) && (!phase || d["El Niño Phase*"] === phase) && (!status || d["Implementation Status*"] === status);
  });
  renderAll();
}

function sum(field) { return filteredActivities.reduce((total,d) => total + (Number(d[field]) || 0), 0); }
function renderKpis() {
  const countries = new Set(filteredActivities.map(d => d["Country*"])).size;
  const completed = filteredActivities.filter(d => d["Implementation Status*"] === "Completed").length;
  const completion = filteredActivities.length ? Math.round(completed / filteredActivities.length * 100) : 0;
  const cards = [
    ["Activities", fmtNum.format(filteredActivities.length), `${countries} locations represented`, true],
    ["People reached", fmtNum.format(sum("People Reached (Total)")), "Direct reach reported"],
    ["Children reached", fmtNum.format(sum("Children Reached")), "Included in total reach"],
    ["Funding used", fmtUSD.format(sum("Funding Used (USD)")), "Reported expenditure"],
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
  $("latest-activities").innerHTML = latest.length ? latest.map(d => `<button class="activity-card" data-entry="${escapeHtml(d["Entry ID"])}" type="button"><div class="meta"><span>${fmtDate.format(new Date(d["Activity Date*"]))} · ${escapeHtml(d["UNICEF Sector*"])}</span><span class="status-pill ${statusClass(d["Implementation Status*"])}">${escapeHtml(d["Implementation Status*"])}</span></div><h3>${escapeHtml(d["Activity Title*"])}</h3><p>${escapeHtml(d["Country*"])} · ${escapeHtml(d["Result / Output"] || "Result not reported")}</p></button>`).join("") : `<div class="empty-state">No activities match these filters.</div>`;
}
function renderTimeline() {
  const rows = [...filteredActivities].sort((a,b)=>new Date(b["Activity Date*"])-new Date(a["Activity Date*"]));
  $("timeline").innerHTML = rows.length ? rows.map(d => `<article class="timeline-entry" style="--phase-color:${phaseColor(d["El Niño Phase*"])}"><time class="timeline-date">${fmtDate.format(new Date(d["Activity Date*"]))}</time><button class="timeline-card" data-entry="${escapeHtml(d["Entry ID"])}" type="button"><div class="timeline-card-head"><div><span class="phase-pill" style="color:${phaseColor(d["El Niño Phase*"])};background:${phaseColor(d["El Niño Phase*"])}18">${escapeHtml(d["El Niño Phase*"])}</span><h3>${escapeHtml(d["Activity Title*"])}</h3></div><span class="status-pill ${statusClass(d["Implementation Status*"])}">${escapeHtml(d["Implementation Status*"])}</span></div><p>${escapeHtml(d["What Was Done?*"])}</p><div class="timeline-tags"><span class="soft-tag">${escapeHtml(d["UNICEF Sector*"])}</span><span class="soft-tag">${escapeHtml(d["Country*"])}</span><span class="soft-tag">${escapeHtml(d["Activity Type*"])}</span><span class="soft-tag">${fmtNum.format(d["Children Reached"])} children</span></div></button></article>`).join("") : `<div class="empty-state">No activities match these filters.</div>`;
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
    marker.bindPopup(`<h3>${escapeHtml(d["Activity Title*"])}</h3><p><strong>${escapeHtml(d["Country*"])}</strong> · ${escapeHtml(d["Location / Admin Area"])}</p><p>${escapeHtml(d["UNICEF Sector*"])} · ${escapeHtml(d["Implementation Status*"])}</p><p>${fmtDate.format(new Date(d["Activity Date*"]))}</p>`);
    marker.addTo(markerLayer); bounds.push([d.latitude,lng]);
  });
  if (bounds.length) map.fitBounds(bounds,{padding:[40,40],maxZoom:5});
  const countries = group("Country*");
  $("map-count").textContent = `${countries.length} mapped ${countries.length === 1 ? "location" : "locations"}`;
  $("map-country-list").innerHTML = countries.map(([name,count]) => `<div class="country-row"><span>${escapeHtml(name)}</span><strong>${count}</strong></div>`).join("");
}

function renderRegister() {
  $("register-body").innerHTML = filteredActivities.length ? [...filteredActivities].sort((a,b)=>new Date(b["Activity Date*"])-new Date(a["Activity Date*"])).map(d => `<tr><td>${fmtDate.format(new Date(d["Activity Date*"]))}</td><td><strong>${escapeHtml(d["UNICEF Sector*"])}</strong></td><td>${escapeHtml(d["Country*"])}<br><small>${escapeHtml(d["Location / Admin Area"])}</small></td><td>${escapeHtml(d["Activity Title*"])}</td><td>${escapeHtml(d["El Niño Phase*"])}</td><td><span class="status-pill ${statusClass(d["Implementation Status*"])}">${escapeHtml(d["Implementation Status*"])}</span></td><td>${fmtNum.format(d["People Reached (Total)"])}</td><td>${fmtUSD.format(d["Funding Used (USD)"])}</td><td><button class="row-button" data-entry="${escapeHtml(d["Entry ID"])}" type="button">View →</button></td></tr>`).join("") : `<tr><td colspan="9" class="empty-state">No activities match these filters.</td></tr>`;
}

function detailItem(label, value, full = false, format) {
  const shown = format ? format(value) : escapeHtml(value || value === 0 ? value : "Not reported");
  return `<div class="detail-item ${full ? "full" : ""}"><span>${label}</span><p>${shown}</p></div>`;
}
function openDetail(entryId) {
  const d = allActivities.find(row => row["Entry ID"] === entryId); if (!d) return;
  const evidence = metadata.isDemo ? escapeHtml(d["Evidence Link"] || "Not reported") + " (illustrative link)" : (/^https?:\/\//.test(d["Evidence Link"]) ? `<a class="detail-link" href="${escapeHtml(d["Evidence Link"])}" target="_blank" rel="noopener">Open supporting evidence ↗</a>` : "Not reported");
  $("dialog-content").innerHTML = `<header class="dialog-title"><p class="section-kicker">${escapeHtml(d["Entry ID"])}</p><h2>${escapeHtml(d["Activity Title*"])}</h2><p>${fmtDate.format(new Date(d["Activity Date*"]))} · ${escapeHtml(d["UNICEF Sector*"])} · ${escapeHtml(d["Country*"])}</p></header><div class="detail-groups">
    <section class="detail-group"><h3>System & classification</h3><div class="detail-grid">${detailItem("Entry ID",d["Entry ID"])}${detailItem("Activity date",d["Activity Date*"],false,v=>fmtDate.format(new Date(v)))}${detailItem("Reporting month",d["Reporting Month"])}${detailItem("UNICEF Sector",d["UNICEF Sector*"])}${detailItem("Country",d["Country*"])}${detailItem("Location / admin area",d["Location / Admin Area"])}${detailItem("El Niño phase",d["El Niño Phase*"])}${detailItem("Activity type",d["Activity Type*"])}${detailItem("Implementation status",d["Implementation Status*"])}</div></section>
    <section class="detail-group"><h3>Results & delivery</h3><div class="detail-grid">${detailItem("People reached",d["People Reached (Total)"],false,v=>fmtNum.format(v))}${detailItem("Children reached",d["Children Reached"],false,v=>fmtNum.format(v))}${detailItem("Funding used",d["Funding Used (USD)"],false,v=>fmtUSD.format(v))}${detailItem("Partners",d["Partners"],true)}${detailItem("Result / output",d["Result / Output"],true)}</div></section>
    <section class="detail-group full"><h3>Activity narrative</h3><div class="detail-grid">${detailItem("Activity title",d["Activity Title*"],true)}${detailItem("What was done?",d["What Was Done?*"],true)}</div></section>
    <section class="detail-group full"><h3>Follow-up & evidence</h3><div class="detail-grid">${detailItem("Challenges",d["Challenges"],true)}${detailItem("Next step",d["Next Step"],true)}${detailItem("Evidence link",evidence,true,v=>v)}${detailItem("Focal point",d["Focal Point*"])}${detailItem("Submission date",d["Submission Date*"],false,v=>fmtDate.format(new Date(v)))}</div></section>
  </div>`;
  $("activity-dialog").showModal();
}

function renderAll() {
  renderKpis(); renderMonthly(); renderPhase(); renderRankBars("Sector-bars",group("UNICEF Sector*")); renderRankBars("reach-bars",group("UNICEF Sector*","People Reached (Total)")); renderFollowup(); renderLatest(); renderTimeline(); renderRegister(); renderMap();
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
["search-filter","Sector-filter","country-filter","phase-filter","status-filter"].forEach(id => $(id).addEventListener(id === "search-filter" ? "input" : "change",applyFilters));
$("clear-filters").addEventListener("click",()=>{ ["search-filter","Sector-filter","country-filter","phase-filter","status-filter"].forEach(id=>$(id).value=""); applyFilters(); });
$("refresh-button").addEventListener("click",()=>refreshData(false));
$("connect-button").addEventListener("click",()=>refreshData(true));
$("excel-upload-button").addEventListener("click",()=>$("excel-file-input").click());
$("excel-file-input").addEventListener("change",async event=>{
  const file = event.target.files?.[0];
  if (!file) return;
  const previousFile = uploadedWorkbookFile;
  uploadedWorkbookFile = file;
  const loaded = await refreshData(false);
  if (!loaded) uploadedWorkbookFile = previousFile;
  event.target.value = "";
});
$("dialog-close").addEventListener("click",()=>$("activity-dialog").close());
$("activity-dialog").addEventListener("click",e=>{ if(e.target === $("activity-dialog")) $("activity-dialog").close(); });
document.body.addEventListener("click",e=>{ const target=e.target.closest("[data-entry]"); if(target) openDetail(target.dataset.entry); });

refreshData(false);
setInterval(()=>refreshData(false),Math.max(1,Number(config.refreshMinutes)||5)*60000);
