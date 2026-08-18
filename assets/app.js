/* Employment Center — browser-only sign-in sheet.
 *
 * Nothing is stored on any server. All data lives in the Excel workbook, which
 * is read and written entirely in this browser and kept on the device (via
 * IndexedDB) so sign-ins continue through the day.
 *
 * Two sheets are managed:
 *   "Intakes"  — one row per client, using the workbook's demographic columns.
 *                A client is "signed up" once they have a row here.
 *   "Sign Ins" — one row per visit (date, time, name, DOB, visitor type).
 */
"use strict";

// ------------------------------------------------------------------ config
const INTAKE_SHEET = "Intakes";
const SIGNIN_SHEET = "Sign Ins";
const MONTHS = ["January","February","March","April","May","June","July",
                "August","September","October","November","December"];

// Exact intake headers (spellings preserved from the existing workbook).
const INTAKE_HEADERS = [
  "Last Name","First Name","Birthdate","UID (If applicable)",
  "Veteren","Hispanic",
  "White","Black","Asian","American India","Other/Multi",
  "House Size",
  "Single Non Elderly","Elderly (62+)","Single Parent","Two Parent ","Other",
  "Adult","TAY","Fam. With Minor","Senior",
  "Female Head of","Disabled",
];
// Columns (0-based) that get summed in the TOTALS row: "Veteren" .. "Disabled".
const SUM_FROM = 4, SUM_TO = INTAKE_HEADERS.length - 1;

const SIGNIN_HEADERS = ["Date","Time","Last Name","First Name","DOB","Visitor Type"];

const RACE_COLS = {white:"White", black:"Black", asian:"Asian",
                   american_indian:"American India", other_multi:"Other/Multi"};
const HH_COLS = {single_non_elderly:"Single Non Elderly", elderly:"Elderly (62+)",
                 single_parent:"Single Parent", two_parent:"Two Parent ", other:"Other"};
const POP_COLS = {adult:"Adult", tay:"TAY", family_with_minor:"Fam. With Minor", senior:"Senior"};
const VISITOR_LABELS = {resident:"Resident (Lot O / Program Participant)",
                        walk_in:"Walk-In Visitor / Community Guest"};

// ------------------------------------------------------------------ state
let WB = null;            // SheetJS workbook (in memory)
let WB_NAME = "workbook.xlsx";
let CURRENT = null;       // person being processed
let busy = false;

// ------------------------------------------------------------------ tiny helpers
const $ = (sel) => document.querySelector(sel);
const pad = (n) => String(n).padStart(2, "0");
const norm = (s) => (s == null ? "" : String(s).trim().toLowerCase().replace(/\s+/g, " "));
const keyOf = (last, first, dob) => `${norm(last)}|${norm(first)}|${dob || ""}`;

function todayISO() {
  const d = new Date();
  return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;
}
function nowTime() {
  const d = new Date();
  return `${pad(d.getHours())}:${pad(d.getMinutes())}`;
}
function dobISO(v) {
  if (v == null || v === "") return "";
  if (v instanceof Date) return `${v.getFullYear()}-${pad(v.getMonth() + 1)}-${pad(v.getDate())}`;
  const s = String(v).trim();
  const m = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (m) return `${m[1]}-${m[2]}-${m[3]}`;
  const d = new Date(s);
  return isNaN(d) ? "" : `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;
}
function isoToDate(iso) {                 // noon local avoids timezone day-rollover
  if (!iso) return null;
  const [y, m, d] = iso.split("-").map(Number);
  return new Date(y, m - 1, d, 12, 0, 0);
}

function toast(msg, isErr) {
  const t = $("#toast");
  t.textContent = msg;
  t.classList.toggle("err", !!isErr);
  t.hidden = false;
  clearTimeout(toast._t);
  toast._t = setTimeout(() => (t.hidden = true), 3200);
}

// ------------------------------------------------------------------ IndexedDB (persist working copy on this device)
function idb(store) {
  return new Promise((res, rej) => {
    const r = indexedDB.open("employment-center", 1);
    r.onupgradeneeded = () => r.result.createObjectStore("kv");
    r.onsuccess = () => res(r.result);
    r.onerror = () => rej(r.error);
  });
}
async function idbSet(k, v) {
  const db = await idb();
  return new Promise((res, rej) => {
    const tx = db.transaction("kv", "readwrite");
    tx.objectStore("kv").put(v, k);
    tx.oncomplete = res; tx.onerror = () => rej(tx.error);
  });
}
async function idbGet(k) {
  const db = await idb();
  return new Promise((res, rej) => {
    const tx = db.transaction("kv", "readonly");
    const rq = tx.objectStore("kv").get(k);
    rq.onsuccess = () => res(rq.result); rq.onerror = () => rej(rq.error);
  });
}
async function idbDel(k) {
  const db = await idb();
  return new Promise((res, rej) => {
    const tx = db.transaction("kv", "readwrite");
    tx.objectStore("kv").delete(k);
    tx.oncomplete = res; tx.onerror = () => rej(tx.error);
  });
}

// ------------------------------------------------------------------ workbook I/O
function serialize() {
  return XLSX.write(WB, { type: "array", bookType: "xlsx" });
}
async function persist() {
  await idbSet("workbook", { name: WB_NAME, data: new Uint8Array(serialize()) });
}
function getSheet(name) { return WB && WB.Sheets[name] ? WB.Sheets[name] : null; }
function putSheet(name, ws) {
  WB.Sheets[name] = ws;
  if (!WB.SheetNames.includes(name)) WB.SheetNames.push(name);
}

// --- read helpers ---------------------------------------------------------
function sheetRows(name) {
  const ws = getSheet(name);
  if (!ws) return { headers: [], rows: [] };
  const aoa = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null });
  if (!aoa.length) return { headers: [], rows: [] };
  return { headers: aoa[0].map((h) => (h == null ? "" : String(h))), rows: aoa.slice(1) };
}

function readIntakes() {
  const { headers, rows } = sheetRows(INTAKE_SHEET);
  if (!headers.length) return [];
  const idx = {}; headers.forEach((h, i) => (idx[h] = i));
  const g = (row, h) => (idx[h] == null ? null : row[idx[h]]);
  const truthy = (v) => v === 1 || v === "1" || v === true;
  const pick = (row, map) => {
    for (const k in map) if (truthy(g(row, map[k]))) return k;
    return null;
  };
  const out = [];
  for (const row of rows) {
    if (!row) continue;
    const last = g(row, "Last Name");
    if (last != null && String(last).toUpperCase() === "TOTALS") continue;
    const first = g(row, "First Name");
    if (last == null && first == null) continue;
    out.push({
      last_name: last || "", first_name: first || "",
      dob: dobISO(g(row, "Birthdate")), uid: g(row, "UID (If applicable)") || "",
      veteran: truthy(g(row, "Veteren")) ? 1 : 0,
      hispanic: truthy(g(row, "Hispanic")) ? 1 : 0,
      race: pick(row, RACE_COLS),
      house_size: g(row, "House Size") || null,
      household_type: pick(row, HH_COLS),
      population_category: pick(row, POP_COLS),
      female_head: truthy(g(row, "Female Head of")) ? 1 : 0,
      disabled: truthy(g(row, "Disabled")) ? 1 : 0,
    });
  }
  return out;
}

function readSignins() {
  const { headers, rows } = sheetRows(SIGNIN_SHEET);
  if (!headers.length) return [];
  const idx = {}; headers.forEach((h, i) => (idx[h] = i));
  const g = (row, h) => (idx[h] == null ? null : row[idx[h]]);
  const out = [];
  for (const row of rows) {
    if (!row) continue;
    if (g(row, "Last Name") == null && g(row, "First Name") == null) continue;
    out.push({
      dateISO: dobISO(g(row, "Date")), time: g(row, "Time") || "",
      last_name: g(row, "Last Name") || "", first_name: g(row, "First Name") || "",
      dob: dobISO(g(row, "DOB")), visitor_type: g(row, "Visitor Type") || "",
    });
  }
  return out;
}

// --- write helpers --------------------------------------------------------
function intakeRowArray(rec) {
  const flag = (v) => (v ? 1 : null);
  const one = (val, target) => (val === target ? 1 : null);
  return [
    rec.last_name || "", rec.first_name || "",
    isoToDate(rec.dob), rec.uid || "",
    flag(rec.veteran), flag(rec.hispanic),
    one(rec.race, "white"), one(rec.race, "black"), one(rec.race, "asian"),
    one(rec.race, "american_indian"), one(rec.race, "other_multi"),
    rec.house_size ? Number(rec.house_size) : null,
    one(rec.household_type, "single_non_elderly"), one(rec.household_type, "elderly"),
    one(rec.household_type, "single_parent"), one(rec.household_type, "two_parent"),
    one(rec.household_type, "other"),
    one(rec.population_category, "adult"), one(rec.population_category, "tay"),
    one(rec.population_category, "family_with_minor"), one(rec.population_category, "senior"),
    flag(rec.female_head), flag(rec.disabled),
  ];
}

function writeIntakeSheet(records) {
  const rowArrays = records.map(intakeRowArray);
  const aoa = [INTAKE_HEADERS.slice(), ...rowArrays];
  const n = records.length;
  aoa.push([n ? "TOTALS" : null]);              // totals placeholder row
  const ws = XLSX.utils.aoa_to_sheet(aoa, { cellDates: true });

  // date format on the Birthdate column (index 2)
  for (let r = 1; r <= n; r++) {
    const ref = XLSX.utils.encode_cell({ r, c: 2 });
    if (ws[ref]) ws[ref].z = "m/d/yyyy";
  }
  // TOTALS row: SUM formula + cached value (SheetJS drops formula cells with
  // no cached value, and the cache lets Excel show the total before recalc).
  if (n) {
    const totalR = n + 1; // 0-based row index of the totals row
    for (let c = SUM_FROM; c <= SUM_TO; c++) {
      const col = XLSX.utils.encode_col(c);
      let sum = 0;
      for (const ra of rowArrays) if (typeof ra[c] === "number") sum += ra[c];
      ws[XLSX.utils.encode_cell({ r: totalR, c })] =
        { t: "n", f: `SUM(${col}2:${col}${n + 1})`, v: sum };
    }
  }
  putSheet(INTAKE_SHEET, ws);
}

function writeSigninSheet(records) {
  const aoa = [SIGNIN_HEADERS.slice()];
  records.forEach((v) => aoa.push([
    isoToDate(v.dateISO), v.time || "", v.last_name || "", v.first_name || "",
    isoToDate(v.dob), v.visitor_type || "",
  ]));
  const ws = XLSX.utils.aoa_to_sheet(aoa, { cellDates: true });
  for (let r = 1; r <= records.length; r++) {
    for (const c of [0, 4]) {                    // Date + DOB columns
      const ref = XLSX.utils.encode_cell({ r, c });
      if (ws[ref]) ws[ref].z = "m/d/yyyy";
    }
  }
  putSheet(SIGNIN_SHEET, ws);
}

function appendVisit(person, visitorType) {
  const visits = readSignins();
  visits.push({
    dateISO: todayISO(), time: nowTime(),
    last_name: person.last_name, first_name: person.first_name, dob: person.dob,
    visitor_type: VISITOR_LABELS[visitorType] || visitorType || "",
  });
  writeSigninSheet(visits);
}

// ------------------------------------------------------------------ lookups
function findClient(last, first, dob) {
  const k = keyOf(last, first, dob);
  return readIntakes().find((r) => keyOf(r.last_name, r.first_name, r.dob) === k) || null;
}

// ------------------------------------------------------------------ navigation
function show(id) {
  document.querySelectorAll(".screen").forEach((s) => (s.hidden = true));
  $("#" + id).hidden = false;
  window.scrollTo(0, 0);
}
function goHome() {
  if (!WB) return show("screen-notready");
  $("#form-identify").reset();
  show("screen-home");
}

// ------------------------------------------------------------------ intake form <-> object
function fillIntakeForm(rec, prefillName) {
  const f = $("#form-intake");
  f.reset();
  const set = (name, val) => { if (f.elements[name] != null) f.elements[name].value = val || ""; };
  const check = (name, val) => {
    const el = f.querySelector(`[name="${name}"][value="${val}"]`);
    if (el) el.checked = true;
  };
  set("first_name", (rec && rec.first_name) || prefillName.first);
  set("last_name", (rec && rec.last_name) || prefillName.last);
  set("date_of_birth", (rec && rec.dob) || prefillName.dob);
  set("uid", rec && rec.uid);
  set("house_size", rec && rec.house_size);
  check("veteran", rec ? rec.veteran : 0);
  check("hispanic", rec ? rec.hispanic : 0);
  check("female_head", rec ? rec.female_head : 0);
  check("disabled", rec ? rec.disabled : 0);
  if (rec && rec.race) check("race", rec.race);
  if (rec && rec.household_type) check("household_type", rec.household_type);
  if (rec && rec.population_category) check("population_category", rec.population_category);
  check("visitor_type", "walk_in");
}

function readIntakeForm() {
  const f = $("#form-intake");
  const val = (n) => (f.elements[n] ? f.elements[n].value : "");
  const radio = (n) => { const el = f.querySelector(`[name="${n}"]:checked`); return el ? el.value : ""; };
  return {
    first_name: val("first_name").trim(), last_name: val("last_name").trim(),
    dob: dobISO(val("date_of_birth")), uid: val("uid").trim(),
    veteran: radio("veteran") === "1" ? 1 : 0,
    hispanic: radio("hispanic") === "1" ? 1 : 0,
    race: radio("race") || null,
    house_size: val("house_size") ? Number(val("house_size")) : null,
    household_type: radio("household_type") || null,
    population_category: radio("population_category") || null,
    female_head: radio("female_head") === "1" ? 1 : 0,
    disabled: radio("disabled") === "1" ? 1 : 0,
    visitor_type: radio("visitor_type") || "walk_in",
  };
}

// ------------------------------------------------------------------ actions
function identify(first, last, dobRaw) {
  const dob = dobISO(dobRaw);
  if (!first || !last || !dob) { toast("Please enter first name, last name, and date of birth.", true); return; }
  const existing = findClient(last, first, dob);
  if (existing) {
    CURRENT = existing;
    $("#welcome-name").textContent = existing.first_name;
    $("#confirm-here").checked = false;
    $("#btn-checkin").disabled = true;
    show("screen-welcome");
  } else {
    CURRENT = { first_name: first, last_name: last, dob };
    fillIntakeForm(null, { first, last, dob });
    show("screen-intake");
  }
}

async function submitIntake() {
  if (busy) return; busy = true;
  try {
    const rec = readIntakeForm();
    if (!rec.first_name || !rec.last_name || !rec.dob) {
      toast("First name, last name, and date of birth are required.", true); busy = false; return;
    }
    const intakes = readIntakes();
    const k = keyOf(rec.last_name, rec.first_name, rec.dob);
    const i = intakes.findIndex((r) => keyOf(r.last_name, r.first_name, r.dob) === k);
    if (i >= 0) intakes[i] = rec; else intakes.push(rec);   // one row per client (no dup)
    writeIntakeSheet(intakes);
    appendVisit(rec, rec.visitor_type);
    await persist();
    finish(rec.first_name);
  } catch (e) {
    console.error(e); toast("Could not save: " + e.message, true);
  } finally { busy = false; }
}

async function confirmCheckin() {
  if (busy || !CURRENT) return; busy = true;
  try {
    appendVisit(CURRENT, CURRENT.visitor_type || "walk_in");
    await persist();
    finish(CURRENT.first_name);
  } catch (e) {
    console.error(e); toast("Could not sign in: " + e.message, true);
  } finally { busy = false; }
}

function finish(name) {
  $("#done-name").textContent = name ? `Thank you, ${name}.` : "Thank you.";
  show("screen-done");
  clearTimeout(finish._t);
  finish._t = setTimeout(goHome, 4500);
}

// ------------------------------------------------------------------ staff / workbook
function refreshStaff() {
  $("#today-label").textContent = new Date().toLocaleDateString(undefined,
    { weekday: "long", year: "numeric", month: "long", day: "numeric" });
  const status = $("#wb-status");
  if (!WB) { status.textContent = "No workbook loaded."; $("#stat-grid").innerHTML = ""; return; }
  status.innerHTML = `Loaded: <strong>${WB_NAME}</strong>`;

  const visits = readSignins();
  const intakes = readIntakes();
  const t = todayISO();
  const todays = visits.filter((v) => v.dateISO === t);
  const uniqToday = new Set(todays.map((v) => keyOf(v.last_name, v.first_name, v.dob)));
  const stat = (num, lbl) => `<div class="stat"><div class="num">${num}</div><div class="lbl">${lbl}</div></div>`;
  $("#stat-grid").innerHTML =
    stat(todays.length, "Sign-ins today") +
    stat(uniqToday.size, "People today") +
    stat(intakes.length, "Clients signed up") +
    stat(visits.length, "Sign-ins total");
}

async function loadWorkbookFromFile(file) {
  const buf = await file.arrayBuffer();
  WB = XLSX.read(new Uint8Array(buf), { type: "array", cellDates: true });
  WB_NAME = file.name;
  await persist();
  $("#wb-filename").textContent = "Loaded: " + file.name;
  toast("Workbook loaded.");
  refreshStaff();
}

function downloadWorkbook() {
  if (!WB) { toast("Load a workbook first.", true); return; }
  const blob = new Blob([serialize()],
    { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob);
  a.download = WB_NAME || "workbook.xlsx";
  document.body.appendChild(a); a.click(); a.remove();
  setTimeout(() => URL.revokeObjectURL(a.href), 1000);
}

// ------------------------------------------------------------------ wiring
function wire() {
  document.body.addEventListener("click", (e) => {
    const go = e.target.closest("[data-go]");
    if (!go) return;
    const dest = go.dataset.go;
    if (dest === "staff") { refreshStaff(); show("screen-staff"); }
    else if (dest === "home") goHome();
  });

  $("#form-identify").addEventListener("submit", (e) => {
    e.preventDefault();
    const f = e.target;
    identify(f.first_name.value.trim(), f.last_name.value.trim(), f.date_of_birth.value);
  });

  $("#form-intake").addEventListener("submit", (e) => { e.preventDefault(); submitIntake(); });

  $("#confirm-here").addEventListener("change", (e) => {
    $("#btn-checkin").disabled = !e.target.checked;
  });
  $("#btn-checkin").addEventListener("click", confirmCheckin);

  // Dropzone
  const zone = $("#dropzone"), input = $("#wb-input");
  zone.addEventListener("click", () => input.click());
  input.addEventListener("change", () => input.files[0] && loadWorkbookFromFile(input.files[0]));
  ["dragenter", "dragover"].forEach((ev) =>
    zone.addEventListener(ev, (e) => { e.preventDefault(); zone.classList.add("drag"); }));
  ["dragleave", "drop"].forEach((ev) =>
    zone.addEventListener(ev, (e) => { e.preventDefault(); zone.classList.remove("drag"); }));
  zone.addEventListener("drop", (e) => {
    const f = e.dataTransfer.files[0];
    if (f) loadWorkbookFromFile(f);
  });

  $("#btn-download").addEventListener("click", downloadWorkbook);
  $("#btn-clear").addEventListener("click", async () => {
    if (!confirm("Remove the working copy from this device? Download it first if you want to keep it.")) return;
    await idbDel("workbook"); WB = null; WB_NAME = "workbook.xlsx";
    refreshStaff(); toast("Cleared from this device.");
  });
}

// ------------------------------------------------------------------ boot
(async function boot() {
  wire();
  try {
    const saved = await idbGet("workbook");
    if (saved && saved.data) {
      WB = XLSX.read(saved.data, { type: "array", cellDates: true });
      WB_NAME = saved.name || "workbook.xlsx";
    }
  } catch (e) { console.error("restore failed", e); }
  goHome();
})();
