"use strict";

/* ================= Storage ================= */
const STORAGE_KEY = "uitzendschema.schedules.v1";

function loadSchedules() {
  try {
    return JSON.parse(localStorage.getItem(STORAGE_KEY)) || [];
  } catch {
    return [];
  }
}
function saveSchedules() {
  try {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(schedules));
  } catch (e) {
    toast("⚠️ Opslaan mislukt (opslag vol?)");
  }
}

/* ================= State ================= */
let schedules = loadSchedules();
let anchorDate = new Date();
let currentView = localStorage.getItem(STORAGE_KEY + ".view") || "week";
let searchQuery = "";
let typeFilterState = loadTypeFilters();
let groupCollapsed = loadGroupCollapsed();

function loadTypeFilters() {
  try {
    return JSON.parse(localStorage.getItem(STORAGE_KEY + ".types")) || {};
  } catch {
    return {};
  }
}
function saveTypeFilters() {
  localStorage.setItem(STORAGE_KEY + ".types", JSON.stringify(typeFilterState));
}
function loadGroupCollapsed() {
  try {
    return JSON.parse(localStorage.getItem(STORAGE_KEY + ".groups")) || {};
  } catch {
    return {};
  }
}
function saveGroupCollapsed() {
  localStorage.setItem(STORAGE_KEY + ".groups", JSON.stringify(groupCollapsed));
}

const COLORS = [
  "#2563eb", "#db2777", "#059669", "#d97706", "#7c3aed",
  "#dc2626", "#0891b2", "#65a30d", "#c026d3", "#ea580c",
];

const STATIONS = {
  "Qmusic": "#e4003a",
  "JOE": "#0055ff",
};
const PALETTE = [
  "#e4003a", "#0055ff", "#dc2626", "#ea580c", "#d97706",
  "#ca8a04", "#65a30d", "#059669", "#0d9488", "#0891b2",
  "#2563eb", "#7c3aed", "#c026d3", "#db2777", "#475569",
];
function detectStation(text) {
  if (/qmusic|q-music/i.test(text)) return "Qmusic";
  if (/\bjoe\b/i.test(text)) return "JOE";
  return "";
}

/* ================= Date helpers ================= */
function mondayOf(d) {
  const date = new Date(d.getFullYear(), d.getMonth(), d.getDate());
  const day = (date.getDay() + 6) % 7; // ma=0 ... zo=6
  date.setDate(date.getDate() - day);
  return date;
}
function addDays(d, n) {
  const r = new Date(d);
  r.setDate(r.getDate() + n);
  return r;
}
function isoDate(d) {
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(d.getDate()).padStart(2, "0")}`;
}
function weekNumber(d) {
  const date = new Date(Date.UTC(d.getFullYear(), d.getMonth(), d.getDate()));
  const dayNum = date.getUTCDay() || 7;
  date.setUTCDate(date.getUTCDate() + 4 - dayNum);
  const yearStart = new Date(Date.UTC(date.getUTCFullYear(), 0, 1));
  return Math.ceil(((date - yearStart) / 86400000 + 1) / 7);
}
const DAY_NAMES = ["Maandag", "Dinsdag", "Woensdag", "Donderdag", "Vrijdag", "Zaterdag", "Zondag"];
const MONTHS = ["jan", "feb", "mrt", "apr", "mei", "jun", "jul", "aug", "sep", "okt", "nov", "dec"];
const MONTHS_FULL = ["januari", "februari", "maart", "april", "mei", "juni", "juli", "augustus", "september", "oktober", "november", "december"];

/* ================= Spot type classification ================= */
function classifyType(name, sameTimeAsPrev, prevLongerThanThis) {
  const n = name.toLowerCase();
  if (/audioboard|audio board/.test(n)) return "Audioboard";
  if (/audio ?tag/.test(n)) return "Audio tag";
  if (/tag ?-? ?on/.test(n)) return "Tag-on";
  if (/\bpromo\b/.test(n)) return "Promo";
  if (/\bcromo\b/.test(n)) return "Cromo";
  // Heuristic: short spot directly after longer spot on exact same date+time = tag-on
  if (sameTimeAsPrev && prevLongerThanThis) return "Tag-on";
  return "Spot";
}

const TYPE_ICONS = {
  "Spot": "🎬", "Promo": "📣", "Cromo": "📣", "Audioboard": "🔊",
  "Audio tag": "🏷️", "Tag-on": "➕",
};
const TYPE_COLORS = {
  "Spot": "#1d4ed8", "Promo": "#be185d", "Cromo": "#a21caf",
  "Audioboard": "#6d28d9", "Audio tag": "#b45309", "Tag-on": "#047857",
};

/* ================= XLS parsing ================= */
function parseWorkbook(wb, fileName) {
  // Find the sheet containing the schedule table. The header row always has
  // "Dag", "Datum" and "Tijd"; the spot description column varies per format
  // (Commercial, Programma, ...). Scan all sheets.
  let rows = null, headerIdx = -1;
  for (const sheetName of wb.SheetNames) {
    const candidate = XLSX.utils.sheet_to_json(wb.Sheets[sheetName], { header: 1, raw: false, defval: "" });
    for (let i = 0; i < Math.min(candidate.length, 60); i++) {
      const cells = candidate[i].map(v => String(v).toLowerCase().trim());
      if (cells.includes("tijd") && (cells.includes("datum") || cells.includes("commercial") || cells.includes("programma"))) {
        rows = candidate;
        headerIdx = i;
        break;
      }
    }
    if (rows) break;
  }
  // Fallback: no header row found — look for a sheet with data rows directly
  // (day abbrev + dd-mm-yyyy date + hh:mm time in the first columns)
  if (headerIdx === -1) {
    for (const sheetName of wb.SheetNames) {
      const candidate = XLSX.utils.sheet_to_json(wb.Sheets[sheetName], { header: 1, raw: false, defval: "" });
      for (let i = 0; i < Math.min(candidate.length, 80); i++) {
        const r = candidate[i].map(v => String(v).trim());
        if (/^\d{1,2}-\d{1,2}-\d{4}$/.test(r[1] || "") && /^\d{1,2}[:.]\d{2}$/.test(r[2] || "")) {
          rows = candidate;
          headerIdx = i - 1; // parse from row i onward
          break;
        }
      }
      if (rows) break;
    }
  }
  if (headerIdx === -1 && !rows) throw new Error("Geen geldige schema-indeling gevonden (geen kop met 'Tijd'/'Datum' en geen datarijen)");

  const header = (rows[Math.max(headerIdx, 0)] || []).map(v => String(v).toLowerCase().trim());
  const find = (...names) => header.findIndex(h => names.some(n => h === n || h.startsWith(n)));
  let col = {
    day: find("dag"),
    date: find("datum"),
    time: find("tijd"),
    number: find("nummer"),
    name: find("commercial"),
    program: find("programma"),
    edition: find("editie"),
    length: find("spotlengte", "lengte"),
  };
  // Derive missing positional columns from "Tijd"
  if (col.time === -1) col.time = 2;
  if (col.date === -1) col.date = col.time - 1;
  if (col.day === -1) col.day = col.time - 2;

  const items = [];
  let prev = null;
  for (let i = Math.max(headerIdx, 0) + 1; i < rows.length; i++) {
    const r = rows[i];
    const dayVal = String(r[Math.max(col.day, 0)] || "").trim().toLowerCase();
    if (dayVal === "totaal") break;
    const dateRaw = String(r[col.date] || "").trim();
    const m = dateRaw.match(/^(\d{1,2})-(\d{1,2})-(\d{4})$/);
    if (!m) continue;
    const date = `${m[3]}-${m[2].padStart(2, "0")}-${m[1].padStart(2, "0")}`;
    const time = String(r[col.time] || "").trim().replace(".", ":");
    if (!/^\d{1,2}:\d{2}$/.test(time)) continue;

    // Spot description: Commercial column, else Programma (+ Editie), else generic
    let name = col.name >= 0 ? String(r[col.name] || "").trim() : "";
    if (!name && col.program >= 0) {
      const prog = String(r[col.program] || "").trim();
      const ed = col.edition >= 0 ? String(r[col.edition] || "").trim() : "";
      name = [prog, ed].filter(Boolean).join(" · ");
    }
    if (!name) name = "Spot";

    const length = col.length >= 0 ? (parseFloat(String(r[col.length]).replace(",", ".")) || 0) : 0;
    const sameTime = prev && prev.date === date && prev.time === time;
    // For grouped spots at the same time, compare with the longest spot in the group
    const groupMax = sameTime ? Math.max(prev.groupMax ?? prev.length, prev.length) : 0;
    const type = classifyType(name, sameTime, groupMax > length);
    const item = {
      date, time, name, length, type,
      number: col.number >= 0 ? String(r[col.number] || "").trim() : "",
    };
    item.groupMax = sameTime ? Math.max(groupMax, length) : length;
    items.push(item);
    prev = item;
  }
  for (const it of items) delete it.groupMax;
  if (items.length === 0) throw new Error("Geen uitzendingen gevonden in dit bestand");

  // Schedule display name: campaign + station if present, else filename
  let title = fileName.replace(/\.(xls|xlsx)$/i, "");
  let campaign = null, station = null;
  for (let i = 0; i < headerIdx; i++) {
    const r = rows[i];
    for (let c = 0; c < r.length; c++) {
      const cell = String(r[c]).trim();
      if (/^campagne:?$/i.test(cell)) {
        const val = r.slice(c + 1).map(v => String(v).trim()).find(v => v);
        if (val) campaign = val;
      }
      const m = cell.match(/^uitzendschema\s+([^\-].*)/i);
      if (m) station = m[1].trim();
    }
  }
  if (campaign) title = campaign + (station ? " — " + station : "");
  return { name: title, items };
}

async function handleFiles(fileList) {
  for (const file of fileList) {
    try {
      const buf = await file.arrayBuffer();
      const wb = XLSX.read(buf, { type: "array" });
      const parsed = parseWorkbook(wb, file.name);
      // Replace if same name already loaded
      const existing = schedules.findIndex(s => s.name === parsed.name);
      const station = detectStation(parsed.name + " " + parsed.items.slice(0, 5).map(i => i.name).join(" "));
      const schedule = {
        id: Date.now() + "_" + Math.random().toString(36).slice(2, 8),
        name: parsed.name,
        color: STATIONS[station] || COLORS[schedules.length % COLORS.length],
        station,
        enabled: true,
        items: parsed.items,
      };
      if (existing >= 0) {
        schedule.color = schedules[existing].color;
        schedule.station = schedules[existing].station || schedule.station;
        schedules[existing] = schedule;
        toast(`🔄 "${parsed.name}" bijgewerkt (${parsed.items.length} uitzendingen)`);
      } else {
        schedules.push(schedule);
        toast(`✅ "${parsed.name}" geladen (${parsed.items.length} uitzendingen)`);
      }
      // Jump to the first week of this schedule
      anchorDate = parseISO(parsed.items[0].date);
    } catch (e) {
      toast(`❌ ${file.name}: ${e.message}`);
    }
  }
  saveSchedules();
  renderAll();
}

function parseISO(s) {
  const [y, m, d] = s.split("-").map(Number);
  return new Date(y, m - 1, d);
}

/* ================= Rendering ================= */
function activeItems() {
  const q = searchQuery.trim().toLowerCase();
  const out = [];
  for (const s of schedules) {
    if (!s.enabled) continue;
    for (const it of s.items) {
      if (typeFilterState[it.type] === false) continue;
      if (q && !it.name.toLowerCase().includes(q) && !(it.number || "").toLowerCase().includes(q)) continue;
      out.push({ ...it, schedule: s });
    }
  }
  return out;
}

function allTypes() {
  const set = new Set();
  for (const s of schedules) for (const it of s.items) set.add(it.type);
  return [...set].sort();
}

const GROUP_FALLBACK = "Niet gespecificeerd";
function stationGroup(s) {
  return STATIONS[s.station] ? s.station : GROUP_FALLBACK;
}

function renderSidebar() {
  const list = document.getElementById("scheduleList");
  list.innerHTML = "";
  document.getElementById("emptyMsg").style.display = schedules.length ? "none" : "block";

  const groups = [...Object.keys(STATIONS), GROUP_FALLBACK];
  for (const g of groups) {
    const members = schedules.filter(s => stationGroup(s) === g);
    if (!members.length) continue;
    const collapsed = groupCollapsed[g] === true;
    const li = document.createElement("li");
    li.className = "sched-group" + (collapsed ? " collapsed" : "");
    li.innerHTML = `
      <button class="sg-head" data-group="${escapeHtml(g)}">
        <span class="sg-chevron">▸</span>
        <span class="color-dot" style="background:${STATIONS[g] || "#7c8494"}"></span>
        <span class="sg-name">${escapeHtml(g)}</span>
        <span class="count">${members.length}</span>
      </button>
      <ul class="sg-list">
        ${members.map(s => `
          <li class="sched-card${s.enabled ? "" : " off"}">
            <div class="sc-main">
              <label class="sc-toggle" title="${s.enabled ? "Uitschakelen" : "Inschakelen"}">
                <input type="checkbox" ${s.enabled ? "checked" : ""} data-id="${s.id}">
                <span class="sc-box" style="--c:${s.color}"></span>
              </label>
              <span class="sched-name" data-stats="${s.id}" title="Klik voor overzicht en beheer">${escapeHtml(s.name)}</span>
              <span class="count">${s.items.length}</span>
            </div>
          </li>`).join("")}
      </ul>`;
    list.appendChild(li);
  }

  list.querySelectorAll(".sg-head").forEach(btn => {
    btn.addEventListener("click", () => {
      const g = btn.dataset.group;
      groupCollapsed[g] = !groupCollapsed[g];
      saveGroupCollapsed();
      btn.closest(".sched-group").classList.toggle("collapsed", groupCollapsed[g]);
    });
  });
  list.querySelectorAll("input[type=checkbox]").forEach(cb => {
    cb.addEventListener("change", () => {
      const s = schedules.find(x => x.id === cb.dataset.id);
      s.enabled = cb.checked;
      // Update card state live (dim + tooltip)
      cb.closest(".sched-card").classList.toggle("off", !cb.checked);
      cb.closest(".sc-toggle").title = cb.checked ? "Uitschakelen" : "Inschakelen";
      saveSchedules();
      renderWeekSelect();
      renderView();
    });
  });
  list.querySelectorAll("[data-stats]").forEach(el => {
    el.addEventListener("click", () => {
      const s = schedules.find(x => x.id === el.dataset.stats);
      openStatsModal(s);
    });
  });

  // Type filters
  const tf = document.getElementById("typeFilters");
  tf.innerHTML = "";
  for (const t of allTypes()) {
    const li = document.createElement("li");
    const checked = typeFilterState[t] !== false;
    li.innerHTML = `
      <label class="sched-item">
        <input type="checkbox" ${checked ? "checked" : ""} data-type="${escapeHtml(t)}">
        <span>${TYPE_ICONS[t] || "\ud83c\udfb5"} ${escapeHtml(t)}</span>
      </label>`;
    tf.appendChild(li);
  }
  tf.querySelectorAll("input").forEach(cb => {
    cb.addEventListener("change", () => {
      typeFilterState[cb.dataset.type] = cb.checked;
      saveTypeFilters();
      renderView();
    });
  });
}

function renderWeekSelect() {
  const sel = document.getElementById("weekSelect");
  const weeks = new Map();
  for (const s of schedules) {
    if (!s.enabled) continue;
    for (const it of s.items) {
      const mon = mondayOf(parseISO(it.date));
      weeks.set(isoDate(mon), mon);
    }
  }
  const sorted = [...weeks.values()].sort((a, b) => a - b);
  sel.innerHTML = `<option value="">— spring naar week —</option>`;
  for (const mon of sorted) {
    const opt = document.createElement("option");
    opt.value = isoDate(mon);
    opt.textContent = `Week ${weekNumber(mon)} · ${mon.getDate()} ${MONTHS[mon.getMonth()]} ${mon.getFullYear()}`;
    if (isoDate(mon) === isoDate(mondayOf(anchorDate))) opt.selected = true;
    sel.appendChild(opt);
  }
}

const HOUR_H = 64; // pixel height of one hour in the time grid

function timeToMin(t) {
  const [h, m] = t.split(":").map(Number);
  return h * 60 + (m || 0);
}

function groupByDate(items) {
  const byDate = {};
  for (const it of items) (byDate[it.date] ||= []).push(it);
  for (const k in byDate) byDate[k].sort((a, b) => a.time.localeCompare(b.time));
  return byDate;
}

function fmtDMY(iso) {
  const d = parseISO(iso);
  return `${d.getDate()} ${MONTHS[d.getMonth()]} ${d.getFullYear()}`;
}

/* ================= View engine ================= */
function renderView() {
  document.querySelectorAll("#viewSwitch button").forEach(b =>
    b.classList.toggle("active", b.dataset.view === currentView));
  const cal = document.getElementById("calendar");
  cal.className = "calendar view-" + currentView;
  cal.innerHTML = "";
  if (currentView === "day") renderTimeGrid(cal, 1);
  else if (currentView === "week") renderTimeGrid(cal, 7);
  else if (currentView === "month") renderMonthView(cal);
  else if (currentView === "list") renderListView(cal);
  else renderDashboard(cal);
}

function setView(v) {
  currentView = v;
  localStorage.setItem(STORAGE_KEY + ".view", v);
  renderView();
}

function navStep(dir) {
  if (currentView === "day") anchorDate = addDays(anchorDate, dir);
  else if (currentView === "month") anchorDate = new Date(anchorDate.getFullYear(), anchorDate.getMonth() + dir, 1);
  else anchorDate = addDays(anchorDate, dir * 7);
  renderWeekSelect();
  renderView();
}

/* ---------- Day & Week view: time grid ---------- */
function renderTimeGrid(cal, nDays) {
  cal.style.setProperty("--hour-h", HOUR_H + "px");

  const byDate = groupByDate(activeItems());
  const start = nDays === 7
    ? mondayOf(anchorDate)
    : new Date(anchorDate.getFullYear(), anchorDate.getMonth(), anchorDate.getDate());

  if (nDays === 7) {
    const sunday = addDays(start, 6);
    document.getElementById("weekLabel").textContent =
      `Week ${weekNumber(start)} · ${start.getDate()} ${MONTHS[start.getMonth()]} – ${sunday.getDate()} ${MONTHS[sunday.getMonth()]} ${sunday.getFullYear()}`;
  } else {
    document.getElementById("weekLabel").textContent =
      `${DAY_NAMES[(start.getDay() + 6) % 7]} ${start.getDate()} ${MONTHS_FULL[start.getMonth()]} ${start.getFullYear()} · Week ${weekNumber(start)}`;
  }

  const todayIso = isoDate(new Date());
  const days = [];
  for (let i = 0; i < nDays; i++) {
    const day = addDays(start, i);
    days.push({ day, iso: isoDate(day), items: byDate[isoDate(day)] || [] });
  }

  // Hour range: from earliest to latest spot in view (with padding), default 06-19
  let minH = 6, maxH = 18;
  const viewItems = days.flatMap(d => d.items);
  if (viewItems.length) {
    const hours = viewItems.map(it => Math.floor(timeToMin(it.time) / 60));
    minH = Math.max(0, Math.min(...hours));
    maxH = Math.min(23, Math.max(...hours));
  }
  const endH = maxH + 1;
  const totalPx = (endH - minH) * HOUR_H;
  const gridCols = `56px repeat(${nDays}, 1fr)`;

  // ----- Sticky header row -----
  const head = document.createElement("div");
  head.className = "grid-head";
  head.style.gridTemplateColumns = gridCols;
  head.innerHTML = `<div class="gh-gutter"></div>` + days.map(d => {
    const di = (d.day.getDay() + 6) % 7;
    return `
    <div class="gh-day ${d.iso === todayIso ? "today" : ""} ${di >= 5 ? "weekend" : ""}">
      <span class="gh-num">${d.day.getDate()}</span>
      <div class="gh-meta">
        <span class="day-name">${DAY_NAMES[di]}</span>
        <span class="day-date">${MONTHS[d.day.getMonth()]} ${d.day.getFullYear()}</span>
      </div>
      ${d.items.length ? `<span class="day-count">${d.items.length}</span>` : ""}
    </div>`;
  }).join("");
  cal.appendChild(head);

  // ----- Body: time gutter + day tracks -----
  const body = document.createElement("div");
  body.className = "grid-body";
  body.style.gridTemplateColumns = gridCols;

  const gutter = document.createElement("div");
  gutter.className = "time-gutter";
  gutter.style.height = totalPx + "px";
  for (let h = minH; h < endH; h++) {
    const lbl = document.createElement("div");
    lbl.className = "time-label";
    lbl.style.top = (h - minH) * HOUR_H + "px";
    lbl.textContent = String(h).padStart(2, "0") + ":00";
    gutter.appendChild(lbl);
  }
  body.appendChild(gutter);

  for (const d of days) {
    const track = document.createElement("div");
    track.className = "day-track" + (d.iso === todayIso ? " today" : "");
    track.style.height = totalPx + "px";

    // Place chips full-width, stacked below each other; push down when they
    // would overlap the previous chip (e.g. spots on the exact same time).
    const CHIP_H = 24, GAP = 2;
    let prevBottom = -Infinity;
    for (const it of d.items) {
      const idealTop = ((timeToMin(it.time) - minH * 60) / 60) * HOUR_H;
      const top = Math.max(idealTop, prevBottom + GAP);
      prevBottom = top + CHIP_H;
      const chip = document.createElement("div");
      chip.className = "chip";
      chip.style.top = top + "px";
      chip.style.left = "3px";
      chip.style.width = "calc(100% - 6px)";
      chip.style.borderLeftColor = it.schedule.color;
      chip.title = `${it.time} · ${it.type} · ${it.length}"\n${it.name}\n${it.schedule.name}${it.number ? "\nSpotnummer: " + it.number : ""}`;
      chip.innerHTML = `<span class="c-time">${escapeHtml(it.time)}</span><span class="c-ico">${TYPE_ICONS[it.type] || ""}</span>${it.length ? `<span class="c-len">${it.length}"</span>` : ""}<span class="c-name">${escapeHtml(it.name)}</span>`;
      chip.addEventListener("click", () => openDetailModal(it, d));
      track.appendChild(chip);
    }
    // Make sure the track is tall enough for pushed-down chips
    track.style.height = Math.max(totalPx, prevBottom + GAP) + "px";

    // Red "now" line on today
    if (d.iso === todayIso) {
      const now = new Date();
      const nowMin = now.getHours() * 60 + now.getMinutes();
      if (nowMin >= minH * 60 && nowMin <= endH * 60) {
        const line = document.createElement("div");
        line.className = "now-line";
        line.style.top = ((nowMin - minH * 60) / 60) * HOUR_H + "px";
        track.appendChild(line);
      }
    }
    body.appendChild(track);
  }
  cal.appendChild(body);
}

/* ---------- Month view ---------- */
function renderMonthView(cal) {
  const y = anchorDate.getFullYear(), m = anchorDate.getMonth();
  document.getElementById("weekLabel").textContent = `${MONTHS_FULL[m]} ${y}`;

  const byDate = groupByDate(activeItems());
  const start = mondayOf(new Date(y, m, 1));
  const todayIso = isoDate(new Date());

  const wrap = document.createElement("div");
  wrap.className = "month-grid";
  wrap.innerHTML = `<div class="mg-head">${DAY_NAMES.map(n => `<div>${n.slice(0, 2)}</div>`).join("")}</div>`;

  const body = document.createElement("div");
  body.className = "mg-body";
  for (let i = 0; i < 42; i++) {
    const day = addDays(start, i);
    if (i % 7 === 0 && i >= 28 && day.getMonth() !== m) break; // skip trailing empty weeks
    const iso = isoDate(day);
    const items = byDate[iso] || [];
    const cell = document.createElement("div");
    cell.className = "mg-cell"
      + (day.getMonth() !== m ? " other" : "")
      + (iso === todayIso ? " today" : "")
      + (items.length ? " has-items" : "");

    const typeCounts = new Map();
    for (const it of items) typeCounts.set(it.type, (typeCounts.get(it.type) || 0) + 1);
    cell.innerHTML = `
      <div class="mg-top"><span class="mg-num">${day.getDate()}</span>${items.length ? `<span class="mg-count">${items.length}</span>` : ""}</div>
      <div class="mg-dots">${[...typeCounts.entries()].map(([t, n]) =>
        `<span class="mg-dot" style="background:${TYPE_COLORS[t] || "#7c8494"}" title="${escapeHtml(t)}: ${n}×"></span>`).join("")}</div>`;
    if (items.length) {
      cell.title = `${items.length} uitzendingen — klik voor dagweergave`;
      cell.addEventListener("click", () => { anchorDate = day; renderWeekSelect(); setView("day"); });
    }
    body.appendChild(cell);
  }
  wrap.appendChild(body);
  cal.appendChild(wrap);
}

/* ---------- List view ---------- */
function renderListView(cal) {
  const items = activeItems();
  document.getElementById("weekLabel").textContent =
    items.length ? `${items.length} uitzendingen` : "Geen uitzendingen";
  if (!items.length) {
    cal.innerHTML = `<p class="view-empty">Geen uitzendingen gevonden${searchQuery.trim() ? " voor deze zoekopdracht" : ""}.</p>`;
    return;
  }
  const byDate = groupByDate(items);
  const todayIso = isoDate(new Date());
  const wrap = document.createElement("div");
  wrap.className = "list-view";

  for (const iso of Object.keys(byDate).sort()) {
    const d = parseISO(iso);
    const group = document.createElement("section");
    group.className = "lv-group" + (iso === todayIso ? " today" : "");
    group.dataset.iso = iso;
    group.innerHTML = `
      <h3 class="lv-date">
        <span class="lv-num">${d.getDate()}</span>
        <span class="lv-daytext">${DAY_NAMES[(d.getDay() + 6) % 7]} ${d.getDate()} ${MONTHS_FULL[d.getMonth()]} ${d.getFullYear()}</span>
        <span class="lv-week">Week ${weekNumber(d)}</span>
        <span class="lv-count">${byDate[iso].length}</span>
      </h3>`;
    const rows = document.createElement("div");
    rows.className = "lv-rows";
    for (const it of byDate[iso]) {
      const row = document.createElement("div");
      row.className = "lv-row";
      row.innerHTML = `
        <span class="lv-time">${escapeHtml(it.time)}</span>
        <span class="lv-type" style="color:${TYPE_COLORS[it.type] || "inherit"}">${TYPE_ICONS[it.type] || ""} ${escapeHtml(it.type)}</span>
        <span class="lv-len">${it.length ? it.length + '"' : ""}</span>
        <span class="lv-name">${escapeHtml(it.name)}</span>
        <span class="lv-sched"><span class="color-dot" style="background:${it.schedule.color}"></span>${escapeHtml(it.schedule.name)}</span>`;
      row.addEventListener("click", () => openDetailModal(it));
      rows.appendChild(row);
    }
    group.appendChild(rows);
    wrap.appendChild(group);
  }
  cal.appendChild(wrap);

  // Scroll to the group at/after the anchor date
  const target = isoDate(anchorDate);
  const groups = [...wrap.querySelectorAll(".lv-group")];
  const el = groups.find(g => g.dataset.iso >= target) || groups[groups.length - 1];
  if (el) el.scrollIntoView({ block: "start" });
}

/* ---------- Dashboard ---------- */
function kpiCard(ico, value, label, cls = "") {
  return `<div class="kpi"><span class="kpi-ico">${ico}</span><div class="kpi-body"><div class="kpi-val ${cls}">${value}</div><div class="kpi-label">${label}</div></div></div>`;
}

function fmtDuration(sec) {
  const h = Math.floor(sec / 3600), m = Math.round((sec % 3600) / 60);
  if (h) return `${h}u ${m}m`;
  if (m) return `${m} min`;
  return `${sec} sec`;
}

function dashSection(title, content, cls = "") {
  const sec = document.createElement("section");
  sec.className = ("dash-card " + cls).trim();
  sec.innerHTML = `<h3>${title}</h3>` + content;
  return sec;
}

function countBy(items, keyFn) {
  const m = new Map();
  for (const it of items) {
    const k = keyFn(it);
    m.set(k, (m.get(k) || 0) + 1);
  }
  return [...m.entries()].sort((a, b) => b[1] - a[1]);
}

function buildDistribution(entries) {
  if (!entries.length) return `<p class="dash-empty">Geen data</p>`;
  const max = Math.max(...entries.map(e => e[1]));
  const total = entries.reduce((s, e) => s + e[1], 0);
  return `<div class="dist">${entries.map(([label, n, color]) => `
    <div class="dist-row">
      <span class="dist-label">${label}</span>
      <div class="dist-track"><div class="dist-bar" style="width:${(n / max * 100).toFixed(1)}%;background:${color}"></div></div>
      <span class="dist-num">${n}<em>${Math.round(n / total * 100)}%</em></span>
    </div>`).join("")}</div>`;
}

function buildGantt(scheds) {
  const rows = scheds.filter(s => s.items.length).map(s => {
    const ds = s.items.map(it => it.date).sort();
    return { s, from: ds[0], to: ds[ds.length - 1] };
  });
  if (!rows.length) return `<p class="dash-empty">Geen data</p>`;
  const min = rows.reduce((a, r) => r.from < a ? r.from : a, rows[0].from);
  const max = rows.reduce((a, r) => r.to > a ? r.to : a, rows[0].to);
  const minT = parseISO(min).getTime();
  const span = Math.max(parseISO(max).getTime() - minT, 86400000);
  return `<div class="gantt">
    <div class="gantt-range"><span>${fmtDMY(min)}</span><span>${fmtDMY(max)}</span></div>
    ${rows.map(r => {
      const l = (parseISO(r.from).getTime() - minT) / span * 100;
      const w = Math.max((parseISO(r.to).getTime() - parseISO(r.from).getTime()) / span * 100, 1.5);
      return `<div class="gantt-row">
        <span class="gantt-label" title="${escapeHtml(r.s.name)}">${escapeHtml(r.s.name)}</span>
        <div class="gantt-track">
          <div class="gantt-bar" style="left:${l.toFixed(2)}%;width:${w.toFixed(2)}%;background:${r.s.color}"
               title="${escapeHtml(r.s.name)}\n${fmtDMY(r.from)} – ${fmtDMY(r.to)} · ${r.s.items.length} spots">
            <span>${r.s.items.length}</span>
          </div>
        </div>
      </div>`;
    }).join("")}
  </div>`;
}

function buildHeatmap(items) {
  const counts = Array.from({ length: 24 }, () => new Array(7).fill(0));
  let minH = 24, maxH = -1, maxC = 0;
  for (const it of items) {
    const h = Math.floor(timeToMin(it.time) / 60);
    const dow = (parseISO(it.date).getDay() + 6) % 7;
    counts[h][dow]++;
    if (h < minH) minH = h;
    if (h > maxH) maxH = h;
    if (counts[h][dow] > maxC) maxC = counts[h][dow];
  }
  if (maxH < 0) return `<p class="dash-empty">Geen data</p>`;
  let html = `<div class="heatmap"><div class="hm-row hm-head"><span></span>${["Ma", "Di", "Wo", "Do", "Vr", "Za", "Zo"].map(d => `<span>${d}</span>`).join("")}</div>`;
  for (let h = minH; h <= maxH; h++) {
    html += `<div class="hm-row"><span class="hm-hour">${String(h).padStart(2, "0")}u</span>${counts[h].map((c, i) => {
      const pct = maxC ? Math.round(c / maxC * 100) : 0;
      return `<span class="hm-cell${c ? "" : " zero"}${pct > 55 ? " hi" : ""}" style="--p:${pct}" title="${DAY_NAMES[i]} ${String(h).padStart(2, "0")}:00–${String(h + 1).padStart(2, "0")}:00 — ${c} spots">${c || ""}</span>`;
    }).join("")}</div>`;
  }
  return html + `</div>`;
}

function buildWeekChart(items) {
  const weeks = new Map();
  for (const it of items) {
    const mon = isoDate(mondayOf(parseISO(it.date)));
    weeks.set(mon, (weeks.get(mon) || 0) + 1);
  }
  const sorted = [...weeks.entries()].sort((a, b) => a[0].localeCompare(b[0]));
  if (!sorted.length) return `<p class="dash-empty">Geen data</p>`;
  const max = Math.max(...sorted.map(([, n]) => n));
  return `<div class="barchart">${sorted.map(([iso, n]) => {
    const d = parseISO(iso);
    return `<div class="bc-col" data-iso="${iso}" title="Week ${weekNumber(d)} (v.a. ${fmtDMY(iso)}) — ${n} spots — klik om te openen">
      <span class="bc-val">${n}</span>
      <div class="bc-bar" style="height:${Math.max(n / max * 100, 3).toFixed(1)}%"></div>
      <span class="bc-label">W${weekNumber(d)}</span>
    </div>`;
  }).join("")}</div>`;
}

function buildDayparts(items) {
  const parts = [
    ["🌙 Nacht", 0, 6], ["🌅 Ochtend", 6, 10], ["☀️ Middag", 10, 14],
    ["🌇 Namiddag", 14, 18], ["🌃 Avond", 18, 24],
  ];
  const entries = parts.map(([label, from, to]) => {
    const n = items.filter(it => {
      const h = Math.floor(timeToMin(it.time) / 60);
      return h >= from && h < to;
    }).length;
    return [`${label} <small>${String(from).padStart(2, "0")}–${String(to).padStart(2, "0")}u</small>`, n, "#6366f1"];
  }).filter(e => e[1] > 0);
  return buildDistribution(entries);
}

function buildOverlaps(items) {
  const map = new Map();
  for (const it of items) {
    const key = it.date + "|" + it.time;
    if (!map.has(key)) map.set(key, []);
    map.get(key).push(it);
  }
  const conflicts = [];
  for (const group of map.values()) {
    if (new Set(group.map(g => g.schedule.id)).size > 1) conflicts.push(group);
  }
  conflicts.sort((a, b) => (a[0].date + a[0].time).localeCompare(b[0].date + b[0].time));
  if (!conflicts.length) return `<p class="dash-ok">✅ Geen overlappende uitzendtijden tussen campagnes.</p>`;
  const shown = conflicts.slice(0, 15);
  return `<div class="overlaps">
    ${shown.map(g => `<div class="ov-row">
      <span class="ov-when">${fmtDMY(g[0].date)} · ${g[0].time}</span>
      <span class="ov-what">${g.map(it =>
        `<span class="ov-item" title="${escapeHtml(it.schedule.name)}"><span class="color-dot" style="background:${it.schedule.color}"></span>${escapeHtml(it.name)}</span>`).join(`<span class="ov-vs">↔</span>`)}</span>
    </div>`).join("")}
    ${conflicts.length > shown.length ? `<p class="dash-more">+ ${conflicts.length - shown.length} meer overlappende tijdstippen</p>` : ""}
  </div>`;
}

function renderDashboard(cal) {
  const items = activeItems();
  document.getElementById("weekLabel").textContent = "Dashboard-overzicht";
  if (!items.length) {
    cal.innerHTML = `<p class="view-empty">Geen data — laad een schema of pas de filters aan.</p>`;
    return;
  }

  const dates = items.map(it => it.date).sort();
  const totalSec = items.reduce((s, it) => s + (it.length || 0), 0);
  const uniqueNames = new Set(items.map(it => it.name)).size;
  const activeScheds = schedules.filter(s => s.enabled);

  const wrap = document.createElement("div");
  wrap.className = "dashboard";
  wrap.insertAdjacentHTML("beforeend", `<div class="kpi-row">
    ${kpiCard("📻", items.length, "Uitzendingen")}
    ${kpiCard("⏱️", fmtDuration(totalSec), "Totale zendtijd")}
    ${kpiCard("📁", activeScheds.length, "Campagnes")}
    ${kpiCard("🎬", uniqueNames, "Unieke items")}
    ${kpiCard("📅", `${fmtDMY(dates[0])} – ${fmtDMY(dates[dates.length - 1])}`, "Periode", "small")}
  </div>`);

  wrap.appendChild(dashSection("📆 Campagne-tijdlijn", buildGantt(activeScheds), "full"));

  const grid = document.createElement("div");
  grid.className = "dash-grid";
  grid.appendChild(dashSection("🔥 Spreiding: uur × dag", buildHeatmap(items)));
  grid.appendChild(dashSection("📊 Spots per week", buildWeekChart(items)));

  const typeCounts = countBy(items, it => it.type);
  grid.appendChild(dashSection("🏷️ Verdeling per type", buildDistribution(
    typeCounts.map(([t, n]) => [`${TYPE_ICONS[t] || ""} ${escapeHtml(t)}`, n, TYPE_COLORS[t] || "#7c8494"]))));

  grid.appendChild(dashSection("🌅 Dagdelen", buildDayparts(items)));

  const stCounts = countBy(items, it => it.schedule.station || "Overig");
  if (stCounts.length > 1) {
    grid.appendChild(dashSection("📡 Per station", buildDistribution(
      stCounts.map(([st, n]) => [escapeHtml(st), n, STATIONS[st] || "#7c8494"]))));
  }

  grid.appendChild(dashSection("⚠️ Overlappende tijden", buildOverlaps(items)));
  wrap.appendChild(grid);
  cal.appendChild(wrap);

  // Week chart columns jump to that week
  wrap.querySelectorAll(".bc-col").forEach(col => {
    col.addEventListener("click", () => {
      anchorDate = parseISO(col.dataset.iso);
      renderWeekSelect();
      setView("week");
    });
  });
}

function renderAll() {
  renderSidebar();
  renderWeekSelect();
  renderView();
}

/* ================= Edit modal ================= */
let editTarget = null;
let editState = { station: "", color: "" };

function openEditModal(s) {
  editTarget = s;
  editState = { station: s.station || "", color: s.color };
  document.getElementById("editName").value = s.name;
  renderEditControls();
  document.getElementById("editModal").hidden = false;
  document.getElementById("editName").focus();
}

function renderEditControls() {
  const stEl = document.getElementById("editStations");
  stEl.innerHTML = Object.entries(STATIONS).map(([st, c]) => `
    <button class="pill ${editState.station === st ? "active" : ""}" data-est="${st}" style="--c:${c}">${st}</button>`).join("");
  stEl.querySelectorAll("[data-est]").forEach(btn => {
    btn.addEventListener("click", () => {
      const st = btn.dataset.est;
      editState.station = editState.station === st ? "" : st;
      if (STATIONS[editState.station]) editState.color = STATIONS[editState.station];
      renderEditControls();
    });
  });

  const colEl = document.getElementById("editColors");
  colEl.innerHTML = PALETTE.map(c => `
    <button class="swatch ${editState.color.toLowerCase() === c ? "active" : ""}" data-col="${c}" style="--c:${c}" title="${c}"></button>`).join("");
  colEl.querySelectorAll("[data-col]").forEach(btn => {
    btn.addEventListener("click", () => {
      editState.color = btn.dataset.col;
      renderEditControls();
    });
  });
}

function closeEditModal() {
  document.getElementById("editModal").hidden = true;
  editTarget = null;
}

/* ================= Stats modal ================= */
let statsTarget = null;

function openStatsModal(s) {
  statsTarget = s;
  document.getElementById("statsName").textContent = s.name;

  const dates = s.items.map(it => it.date).sort();
  const fmt = iso => { const d = parseISO(iso); return `${d.getDate()} ${MONTHS[d.getMonth()]} ${d.getFullYear()}`; };
  document.getElementById("statsSummary").textContent =
    `${s.items.length} uitzendingen · ${fmt(dates[0])} t/m ${fmt(dates[dates.length - 1])}`;

  // Counts per type
  const typeCounts = new Map();
  for (const it of s.items) typeCounts.set(it.type, (typeCounts.get(it.type) || 0) + 1);
  document.getElementById("statsTypes").innerHTML = [...typeCounts.entries()]
    .sort((a, b) => b[1] - a[1])
    .map(([t, n]) => `<span class="stats-type-chip">${TYPE_ICONS[t] || "🎵"} ${escapeHtml(t)} <b>${n}×</b></span>`)
    .join("");

  // Counts per unique item (name + type + length)
  const itemMap = new Map();
  for (const it of s.items) {
    const key = `${it.name}|${it.type}|${it.length}`;
    const e = itemMap.get(key) || { name: it.name, type: it.type, length: it.length, count: 0 };
    e.count++;
    itemMap.set(key, e);
  }
  document.getElementById("statsRows").innerHTML = [...itemMap.values()]
    .sort((a, b) => b.count - a.count || a.name.localeCompare(b.name))
    .map(e => `<tr>
      <td class="st-name">${escapeHtml(e.name)}</td>
      <td>${TYPE_ICONS[e.type] || ""} ${escapeHtml(e.type)}</td>
      <td>${e.length ? e.length + '"' : "—"}</td>
      <td class="num"><b>${e.count}×</b></td>
    </tr>`)
    .join("");

  document.querySelector("#statsModal .stats-modal").style.borderTop = `4px solid ${s.color}`;
  document.getElementById("statsModal").hidden = false;
}
function closeStatsModal() {
  document.getElementById("statsModal").hidden = true;
  statsTarget = null;
}

/* ================= Detail modal ================= */
function openDetailModal(it, dayInfo) {
  const date = parseISO(it.date);
  const dayName = DAY_NAMES[(date.getDay() + 6) % 7];
  document.getElementById("detType").textContent = `${TYPE_ICONS[it.type] || ""} ${it.type}`;
  document.getElementById("detName").textContent = it.name;
  const rows = [
    ["Datum", `${dayName} ${date.getDate()} ${MONTHS[date.getMonth()]} ${date.getFullYear()}`],
    ["Tijd", it.time],
    ["Spotlengte", it.length ? `${it.length} seconden` : "—"],
    ["Spotnummer", it.number || "—"],
    ["Week", `Week ${weekNumber(date)}`],
    ["Schema", it.schedule.name],
  ];
  document.getElementById("detGrid").innerHTML = rows
    .map(([k, v]) => `<dt>${escapeHtml(k)}</dt><dd>${escapeHtml(v)}</dd>`)
    .join("");
  document.querySelector("#detailModal .detail-modal").style.borderTop = `4px solid ${it.schedule.color}`;
  document.getElementById("detailModal").hidden = false;
}
function closeDetailModal() {
  document.getElementById("detailModal").hidden = true;
}

/* ================= iCal export dialog ================= */
let exportTarget = null;

function openExportModal(s) {
  exportTarget = s;
  const dates = s.items.map(it => it.date).sort();
  const todayIso = isoDate(new Date());
  const from = document.getElementById("expFrom");
  const to = document.getElementById("expTo");
  from.min = dates[0]; from.max = dates[dates.length - 1];
  to.min = dates[0]; to.max = dates[dates.length - 1];
  from.value = todayIso > dates[dates.length - 1] ? dates[0] : (todayIso < dates[0] ? dates[0] : todayIso);
  to.value = dates[dates.length - 1];
  document.getElementById("exportSchedName").textContent = s.name;

  // Type checkboxes for the types present in this schedule
  const typesEl = document.getElementById("expTypes");
  typesEl.innerHTML = "";
  for (const t of [...new Set(s.items.map(it => it.type))].sort()) {
    const label = document.createElement("label");
    label.className = "exp-type";
    label.innerHTML = `<input type="checkbox" value="${escapeHtml(t)}" checked> ${TYPE_ICONS[t] || ""} ${escapeHtml(t)}`;
    typesEl.appendChild(label);
  }

  document.getElementById("exportModal").hidden = false;
  updateExportCount();
}

function getExportOptions() {
  const from = document.getElementById("expFrom").value;
  const to = document.getElementById("expTo").value;
  const durVal = document.getElementById("expDur").value;
  const types = new Set([...document.querySelectorAll("#expTypes input:checked")].map(cb => cb.value));
  return { from, to, durVal, types };
}

function filterExportItems(s, opt) {
  return s.items.filter(it =>
    (!opt.from || it.date >= opt.from) &&
    (!opt.to || it.date <= opt.to) &&
    opt.types.has(it.type)
  );
}

function updateExportCount() {
  if (!exportTarget) return;
  const n = filterExportItems(exportTarget, getExportOptions()).length;
  document.getElementById("expCount").textContent = `${n} afspraken in deze selectie`;
  document.getElementById("expConfirm").disabled = n === 0;
}

function closeExportModal() {
  document.getElementById("exportModal").hidden = true;
  exportTarget = null;
}

/* ================= iCal export ================= */
function icsEscape(s) {
  return String(s).replace(/\\/g, "\\\\").replace(/;/g, "\\;").replace(/,/g, "\\,").replace(/\r?\n/g, "\\n");
}
// Fold lines longer than 75 octets (RFC 5545)
function icsFold(line) {
  const out = [];
  while (line.length > 74) {
    out.push(line.slice(0, 74));
    line = " " + line.slice(74);
  }
  out.push(line);
  return out.join("\r\n");
}
function downloadIcal(s, opt) {
  const dtstamp = new Date().toISOString().replace(/[-:]/g, "").replace(/\.\d{3}/, "");
  const items = filterExportItems(s, opt);
  if (items.length === 0) {
    toast("⚠\ufe0f Geen uitzendingen in deze selectie");
    return;
  }
  const lines = [
    "BEGIN:VCALENDAR",
    "VERSION:2.0",
    "PRODID:-//Uitzendschema Viewer//NL",
    "CALSCALE:GREGORIAN",
    "METHOD:PUBLISH",
    icsFold("X-WR-CALNAME:" + icsEscape(s.name)),
    "X-WR-TIMEZONE:Europe/Amsterdam",
  ];
  s.items.forEach((it, idx) => {
    if (!items.includes(it)) return;
    const d = it.date.replace(/-/g, "");
    const [h, m] = it.time.split(":").map(Number);
    const startSec = h * 3600 + m * 60;
    // Duration from dialog: fixed minutes, or real spot length (min. 1 min for visibility)
    const durSec = opt.durVal === "real" ? Math.max(60, it.length || 60) : parseInt(opt.durVal, 10) * 60;
    const endSec = startSec + durSec;
    const fmt = sec => String(Math.floor(sec / 3600)).padStart(2, "0") + String(Math.floor((sec % 3600) / 60)).padStart(2, "0") + String(sec % 60).padStart(2, "0");
    const summary = `${TYPE_ICONS[it.type] || ""} ${it.type} ${it.length ? it.length + '"' : ""} · ${it.name}`.replace(/\s+/g, " ").trim();
    const desc = [`Type: ${it.type}`, `Spotlengte: ${it.length}\"`, it.number ? `Spotnummer: ${it.number}` : "", `Schema: ${s.name}`].filter(Boolean).join("\n");
    lines.push(
      "BEGIN:VEVENT",
      `UID:${s.id}-${idx}@uitzendschema-viewer`,
      `DTSTAMP:${dtstamp}`,
      `DTSTART;TZID=Europe/Amsterdam:${d}T${fmt(startSec)}`,
      `DTEND;TZID=Europe/Amsterdam:${d}T${fmt(endSec)}`,
      icsFold("SUMMARY:" + icsEscape(summary)),
      icsFold("DESCRIPTION:" + icsEscape(desc)),
      "END:VEVENT",
    );
  });
  lines.push("END:VCALENDAR");
  // Minimal VTIMEZONE for Europe/Amsterdam (required by some importers)
  const vtz = [
    "BEGIN:VTIMEZONE", "TZID:Europe/Amsterdam",
    "BEGIN:DAYLIGHT", "TZOFFSETFROM:+0100", "TZOFFSETTO:+0200", "TZNAME:CEST", "DTSTART:19700329T020000", "RRULE:FREQ=YEARLY;BYMONTH=3;BYDAY=-1SU", "END:DAYLIGHT",
    "BEGIN:STANDARD", "TZOFFSETFROM:+0200", "TZOFFSETTO:+0100", "TZNAME:CET", "DTSTART:19701025T030000", "RRULE:FREQ=YEARLY;BYMONTH=10;BYDAY=-1SU", "END:STANDARD",
    "END:VTIMEZONE",
  ];
  lines.splice(7, 0, ...vtz);

  const blob = new Blob([lines.join("\r\n") + "\r\n"], { type: "text/calendar;charset=utf-8" });
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob);
  a.download = s.name.replace(/[\\/:*?"<>|]/g, "_") + ".ics";
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(a.href);
  toast(`⤓ "${s.name}" gedownload als .ics (${items.length} events)`);
}

/* ================= Google Drive abonnement-URL helper ================= */
function openSubModal() {
  document.getElementById("subLink").value = "";
  document.getElementById("subResultRow").hidden = true;
  document.getElementById("subModal").hidden = false;
  document.getElementById("subLink").focus();
}
function closeSubModal() {
  document.getElementById("subModal").hidden = true;
}
function makeSubscribeUrl() {
  const input = document.getElementById("subLink").value.trim();
  const m = input.match(/\/d\/([\w-]{20,})/) || input.match(/[?&]id=([\w-]{20,})/);
  if (!m) {
    toast("⚠️ Geen geldige Google Drive-link herkend");
    return;
  }
  const url = `https://drive.google.com/uc?export=download&id=${m[1]}`;
  const out = document.getElementById("subResult");
  out.value = url;
  document.getElementById("subResultRow").hidden = false;
  out.focus();
  out.select();
  if (navigator.clipboard) {
    navigator.clipboard.writeText(url).then(
      () => toast("🔗 Abonnement-URL gekopieerd!"),
      () => toast("Kopieer de URL hierboven handmatig (Ctrl+C)")
    );
  } else {
    toast("Kopieer de URL hierboven handmatig (Ctrl+C)");
  }
}

/* ================= Print / PDF ================= */
function weeksWithItems() {
  const weeks = new Map();
  for (const it of activeItems()) {
    const mon = isoDate(mondayOf(parseISO(it.date)));
    if (!weeks.has(mon)) weeks.set(mon, 0);
    weeks.set(mon, weeks.get(mon) + 1);
  }
  return [...weeks.entries()].sort((a, b) => a[0].localeCompare(b[0]));
}

function openPrintModal() {
  const weeks = weeksWithItems();
  if (!weeks.length) {
    toast("⚠️ Geen uitzendingen om te printen — laad een schema of pas de filters aan");
    return;
  }
  const curMon = isoDate(mondayOf(anchorDate));
  document.getElementById("printWeeks").innerHTML = weeks.map(([iso, n]) => {
    const mon = parseISO(iso);
    const sun = addDays(mon, 6);
    return `<label class="pw-week">
      <input type="checkbox" value="${iso}" checked>
      <span class="pw-label">Week ${weekNumber(mon)} · ${mon.getDate()} ${MONTHS[mon.getMonth()]} – ${sun.getDate()} ${MONTHS[sun.getMonth()]} ${sun.getFullYear()}${iso === curMon ? " <em>(huidige)</em>" : ""}</span>
      <span class="count">${n}</span>
    </label>`;
  }).join("");
  document.getElementById("printModal").hidden = false;
  updatePrintCount();
}

function closePrintModal() {
  document.getElementById("printModal").hidden = true;
}

function selectedPrintWeeks() {
  return [...document.querySelectorAll("#printWeeks input:checked")].map(cb => cb.value);
}

function updatePrintCount() {
  const boxes = [...document.querySelectorAll("#printWeeks input")];
  const sel = boxes.filter(cb => cb.checked);
  document.getElementById("printCount").textContent =
    sel.length ? `${sel.length} van ${boxes.length} weken geselecteerd` : "Geen weken geselecteerd";
  document.getElementById("printConfirm").disabled = sel.length === 0;
}

function buildPrintArea(mondays) {
  const byDate = groupByDate(activeItems());
  const activeScheds = schedules.filter(s => s.enabled);
  const legend = activeScheds.map(s =>
    `<span class="pw-leg"><span class="pw-dot" style="background:${s.color}"></span>${escapeHtml(s.name)}</span>`).join("");

  const CHIP_MM = 4, GAP_MM = 0.4, GRID_MM = 163; // print sizes in mm (A4 landscape)

  document.getElementById("printArea").innerHTML = mondays.map(iso => {
    const mon = parseISO(iso);
    const sun = addDays(mon, 6);
    const days = [];
    for (let i = 0; i < 7; i++) {
      const d = addDays(mon, i);
      days.push({ d, items: byDate[isoDate(d)] || [] });
    }
    const total = days.reduce((s, d) => s + d.items.length, 0);

    // Hour range for this week (same logic as the on-screen week grid)
    let minH = 6, maxH = 18;
    const weekItems = days.flatMap(d => d.items);
    if (weekItems.length) {
      const hrs = weekItems.map(it => Math.floor(timeToMin(it.time) / 60));
      minH = Math.max(0, Math.min(...hrs));
      maxH = Math.min(23, Math.max(...hrs));
    }
    const endH = maxH + 1;

    // Required height (mm) for a given hour height, incl. chips pushed down on overlap
    const required = (hh) => {
      let worst = (endH - minH) * hh;
      for (const d of days) {
        let prev = -Infinity;
        for (const it of d.items) {
          const ideal = ((timeToMin(it.time) - minH * 60) / 60) * hh;
          prev = Math.max(ideal, prev + GAP_MM) + CHIP_MM;
        }
        if (prev > worst) worst = prev;
      }
      return worst;
    };
    // Largest hour height where everything still fits on one page
    let hourH = GRID_MM / (endH - minH);
    if (required(hourH) > GRID_MM) {
      let lo = 1, hi = hourH;
      for (let i = 0; i < 24; i++) {
        const mid = (lo + hi) / 2;
        if (required(mid) <= GRID_MM) lo = mid; else hi = mid;
      }
      hourH = lo;
    }
    const gridH = Math.min(GRID_MM, required(hourH) + 0.5);

    const gutter = Array.from({ length: endH - minH }, (_, i) =>
      `<span class="pw-hlabel" style="top:${(i * hourH).toFixed(2)}mm">${String(minH + i).padStart(2, "0")}:00</span>`).join("");

    return `<section class="print-week">
      <header class="pw-head">
        <div class="pw-title">
          <h2>Week ${weekNumber(mon)}</h2>
          <span class="pw-range">${mon.getDate()} ${MONTHS[mon.getMonth()]} – ${sun.getDate()} ${MONTHS[sun.getMonth()]} ${sun.getFullYear()} · ${total} uitzendingen</span>
        </div>
        <div class="pw-legend">${legend}</div>
      </header>
      <div class="pw-dayrow">
        <span></span>
        ${days.map(day => {
          const di = (day.d.getDay() + 6) % 7;
          return `<span class="pw-dayhead">${DAY_NAMES[di]} ${day.d.getDate()} ${MONTHS[day.d.getMonth()]}${day.items.length ? `<b>${day.items.length}</b>` : ""}</span>`;
        }).join("")}
      </div>
      <div class="pw-tgrid" style="height:${gridH.toFixed(2)}mm">
        <div class="pw-gutter">${gutter}</div>
        ${days.map(day => {
          let prev = -Infinity;
          const chips = day.items.map(it => {
            const ideal = ((timeToMin(it.time) - minH * 60) / 60) * hourH;
            const top = Math.max(ideal, prev + GAP_MM);
            prev = top + CHIP_MM;
            return `<div class="pw-chip" style="top:${top.toFixed(2)}mm;border-left-color:${it.schedule.color}">
              <span class="pw-time">${escapeHtml(it.time)}</span>${it.length ? `<span class="pw-len">${it.length}"</span>` : ""}<span class="pw-name">${escapeHtml(it.name)}</span>
            </div>`;
          }).join("");
          return `<div class="pw-track" style="background-size:100% ${hourH.toFixed(2)}mm">${chips}</div>`;
        }).join("")}
      </div>
    </section>`;
  }).join("");
}

/* ================= Utils ================= */
function escapeHtml(s) {
  return String(s).replace(/[&<>"']/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
}
let toastTimer;
function toast(msg) {
  const el = document.getElementById("toast");
  el.textContent = msg;
  el.classList.add("show");
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => el.classList.remove("show"), 3500);
}

/* ================= Events ================= */
document.getElementById("uploadBtn").addEventListener("click", () => document.getElementById("fileInput").click());
document.getElementById("fileInput").addEventListener("change", (e) => {
  handleFiles([...e.target.files]);
  e.target.value = "";
});
document.getElementById("prevWeek").addEventListener("click", () => navStep(-1));
document.getElementById("nextWeek").addEventListener("click", () => navStep(1));
document.getElementById("todayBtn").addEventListener("click", () => { anchorDate = new Date(); renderWeekSelect(); renderView(); });
document.getElementById("weekSelect").addEventListener("change", (e) => {
  if (e.target.value) {
    anchorDate = parseISO(e.target.value);
    if (currentView === "month" || currentView === "dashboard") setView("week");
    else renderView();
  }
});
document.querySelectorAll("#viewSwitch button").forEach(btn =>
  btn.addEventListener("click", () => setView(btn.dataset.view)));
document.getElementById("searchInput").addEventListener("input", (e) => {
  searchQuery = e.target.value;
  renderView();
});
document.getElementById("clearAll").addEventListener("click", () => {
  if (schedules.length && confirm("Alle schema's en instellingen wissen?")) {
    schedules = [];
    typeFilterState = {};
    saveSchedules();
    saveTypeFilters();
    renderAll();
  }
});

// Export modal
document.getElementById("expCancel").addEventListener("click", closeExportModal);
document.getElementById("exportModal").addEventListener("click", (e) => {
  if (e.target.id === "exportModal") closeExportModal();
});
// Detail modal
document.getElementById("detClose").addEventListener("click", closeDetailModal);
document.getElementById("detailModal").addEventListener("click", (e) => {
  if (e.target.id === "detailModal") closeDetailModal();
});
// Stats modal
document.getElementById("statsClose").addEventListener("click", closeStatsModal);
document.getElementById("statsModal").addEventListener("click", (e) => {
  if (e.target.id === "statsModal") closeStatsModal();
});
document.getElementById("stEdit").addEventListener("click", () => {
  const s = statsTarget;
  closeStatsModal();
  openEditModal(s);
});
document.getElementById("stIcal").addEventListener("click", () => {
  const s = statsTarget;
  closeStatsModal();
  openExportModal(s);
});
document.getElementById("stSub").addEventListener("click", () => openSubModal());
document.getElementById("stRemove").addEventListener("click", () => {
  const s = statsTarget;
  if (s && confirm(`Schema "${s.name}" verwijderen?`)) {
    schedules = schedules.filter(x => x.id !== s.id);
    saveSchedules();
    closeStatsModal();
    renderAll();
  }
});
document.getElementById("stReplace").addEventListener("click", () => {
  document.getElementById("replaceInput").click();
});
document.getElementById("replaceInput").addEventListener("change", async (e) => {
  const file = e.target.files[0];
  e.target.value = "";
  if (!file || !statsTarget) return;
  try {
    const buf = await file.arrayBuffer();
    const wb = XLSX.read(buf, { type: "array" });
    const parsed = parseWorkbook(wb, file.name);
    statsTarget.name = parsed.name;
    statsTarget.items = parsed.items;
    saveSchedules();
    anchorDate = parseISO(parsed.items[0].date);
    closeStatsModal();
    renderAll();
    toast(`🔄 Schema vervangen door "${parsed.name}" (${parsed.items.length} uitzendingen)`);
  } catch (err) {
    toast(`❌ ${file.name}: ${err.message}`);
  }
});
// Subscribe modal
document.getElementById("subCancel").addEventListener("click", closeSubModal);
document.getElementById("subModal").addEventListener("click", (e) => {
  if (e.target.id === "subModal") closeSubModal();
});
document.getElementById("subMake").addEventListener("click", makeSubscribeUrl);
document.getElementById("subLink").addEventListener("keydown", (e) => {
  if (e.key === "Enter") makeSubscribeUrl();
});
// Print modal
document.getElementById("printBtn").addEventListener("click", openPrintModal);
document.getElementById("printCancel").addEventListener("click", closePrintModal);
document.getElementById("printModal").addEventListener("click", (e) => {
  if (e.target.id === "printModal") closePrintModal();
});
document.getElementById("printWeeks").addEventListener("change", updatePrintCount);
document.getElementById("pwAll").addEventListener("click", () => {
  document.querySelectorAll("#printWeeks input").forEach(cb => cb.checked = true);
  updatePrintCount();
});
document.getElementById("pwNone").addEventListener("click", () => {
  document.querySelectorAll("#printWeeks input").forEach(cb => cb.checked = false);
  updatePrintCount();
});
document.getElementById("printConfirm").addEventListener("click", () => {
  const mondays = selectedPrintWeeks();
  if (!mondays.length) return;
  buildPrintArea(mondays);
  closePrintModal();
  window.print();
});

// Edit modal
document.getElementById("editCancel").addEventListener("click", closeEditModal);
document.getElementById("editModal").addEventListener("click", (e) => {
  if (e.target.id === "editModal") closeEditModal();
});
document.getElementById("editSave").addEventListener("click", () => {
  if (editTarget) {
    const name = document.getElementById("editName").value.trim();
    if (name) editTarget.name = name;
    editTarget.station = editState.station;
    editTarget.color = editState.color;
    saveSchedules();
    renderAll();
  }
  closeEditModal();
});
document.getElementById("editName").addEventListener("keydown", (e) => {
  if (e.key === "Enter") document.getElementById("editSave").click();
});
document.getElementById("expConfirm").addEventListener("click", () => {
  if (exportTarget) downloadIcal(exportTarget, getExportOptions());
  closeExportModal();
});
for (const id of ["expFrom", "expTo", "expDur"]) {
  document.getElementById(id).addEventListener("change", updateExportCount);
}
document.getElementById("expTypes").addEventListener("change", updateExportCount);
document.addEventListener("keydown", (e) => {
  if (e.key === "Escape") { closeExportModal(); closeDetailModal(); closeEditModal(); closeStatsModal(); closePrintModal(); closeSubModal(); return; }
  const tag = (e.target.tagName || "").toLowerCase();
  if (tag === "input" || tag === "select" || tag === "textarea") return;
  if (e.key === "ArrowLeft") navStep(-1);
  else if (e.key === "ArrowRight") navStep(1);
  else if (e.key === "t" || e.key === "T") { anchorDate = new Date(); renderWeekSelect(); renderView(); }
  else if (["1", "2", "3", "4", "5"].includes(e.key)) setView(["day", "week", "month", "list", "dashboard"][+e.key - 1]);
});

// Drag & drop
const zone = document.body;
zone.addEventListener("dragover", (e) => { e.preventDefault(); document.getElementById("uploadZone").classList.add("dragover"); });
zone.addEventListener("dragleave", () => document.getElementById("uploadZone").classList.remove("dragover"));
zone.addEventListener("drop", (e) => {
  e.preventDefault();
  document.getElementById("uploadZone").classList.remove("dragover");
  const files = [...e.dataTransfer.files].filter(f => /\.(xls|xlsx)$/i.test(f.name));
  if (files.length) handleFiles(files);
  else toast("⚠️ Sleep een .xls of .xlsx bestand");
});

/* ================= Init ================= */
renderAll();
