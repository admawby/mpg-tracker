"use strict";
(() => {
  const $ = (s, el = document) => el.querySelector(s);
  const $$ = (s, el = document) => [...el.querySelectorAll(s)];
  const reduced = matchMedia("(prefers-reduced-motion: reduce)").matches;
  const OUTBOX = "mpg-outbox";
  const S = window.MPGStore, LOCAL = S.LOCAL;
  const BASE_TAGS = ["Towing", "Road trip", "Cold weather", "Missed trip reset"];
  // Maintenance categories and default reminder intervals (miles; null = no reminder)
  const SERVICE_DEFAULTS = {
    "Oil change": 7500, "Tire rotation": 7500, "Brakes": null, "Tires": null, "Battery": null,
    "Air filter": 30000, "Cabin filter": 20000, "Alignment": null, "Coolant": 100000,
    "Transmission": 150000, "Spark plugs": 100000, "Wipers": null, "Repair": null,
  };
  const BASE_MODES = ["Normal", "Eco", "Sport"];
  const BASE_OCTANE = [87, 89, 93];

  let data = [];           // fill-ups, oldest first
  let services = [];       // maintenance records, oldest first
  let appSettings = {};    // {intervals: {category: miles|null}}
  let editing = null;      // row being edited, or null for a new fill-up
  const st = { octane: null, mode: null, tags: new Set(), user: {} };
  let range = "all";
  let dirty = false;       // new-entry form has unsaved typing
  const DRAFT = "mpg-draft";

  // ------------------------------------------------------------ utils
  function h(tag, attrs = {}, ...kids) {
    const el = document.createElement(tag);
    for (const [k, v] of Object.entries(attrs)) {
      if (v == null || v === false) continue;
      if (k === "class") el.className = v;
      else if (k === "style") el.style.cssText = v;
      else if (k.startsWith("on")) el.addEventListener(k.slice(2), v);
      else el.setAttribute(k, v === true ? "" : v);
    }
    for (const kid of kids.flat()) if (kid != null && kid !== false) el.append(kid);
    return el;
  }
  const num = (v) => { if (v == null || v === "") return null; const n = parseFloat(String(v).replace(/[,$\s]/g, "")); return Number.isFinite(n) ? n : null; };
  const f = (n, d = 1) => (n == null || !Number.isFinite(n) ? "--" : n.toLocaleString("en-US", { minimumFractionDigits: d, maximumFractionDigits: d }));
  const todayISO = () => { const d = new Date(); return new Date(d - d.getTimezoneOffset() * 6e4).toISOString().slice(0, 10); };
  const toDate = (iso) => new Date(iso + "T12:00:00");
  const days = (a, b) => Math.round((toDate(b) - toDate(a)) / 864e5);
  const fmtDate = (iso, o = { month: "short", day: "numeric" }) => toDate(iso).toLocaleDateString("en-US", o);
  const hm = (mins) => (mins == null ? "" : `${Math.floor(mins / 60)}:${String(mins % 60).padStart(2, "0")}`);
  function parseHM(s) {
    s = (s || "").trim();
    if (!s) return null;
    let m = s.match(/^(\d{1,3})[:.h ](\d{1,2})$/);
    if (m) return +m[1] * 60 + Math.min(59, +m[2]);
    m = s.match(/^\d+$/);
    if (m) return s.length <= 2 ? +s * 60 : +s.slice(0, -2) * 60 + Math.min(59, +s.slice(-2));
    return NaN;
  }
  const calcMpg = (r) => (r.trip_miles && r.gallons ? r.trip_miles / r.gallons : null);
  const tagsOf = (r) => (r.tags ? r.tags.split(",").map((t) => t.trim()).filter(Boolean) : []);
  function avgMpg(rows) { // mile-weighted dash average (equals total miles / total gallons the dash believes)
    let mi = 0, gal = 0, n = 0, sum = 0;
    for (const r of rows) {
      if (!r.dash_mpg) continue;
      n++; sum += r.dash_mpg;
      if (r.trip_miles) { mi += r.trip_miles; gal += r.trip_miles / r.dash_mpg; }
    }
    return gal ? mi / gal : n ? sum / n : null;
  }
  function freq(list) {
    const m = new Map();
    for (const x of list) if (x != null && x !== "") m.set(x, (m.get(x) || 0) + 1);
    return [...m.entries()].sort((a, b) => b[1] - a[1]).map(([k, c]) => ({ k, c }));
  }
  function toast(msg) {
    const t = $("#toast");
    t.textContent = msg;
    t.classList.add("on");
    clearTimeout(toast.t);
    toast.t = setTimeout(() => t.classList.remove("on"), 2600);
  }

  // ------------------------------------------------------------ api
  const outbox = () => JSON.parse(localStorage.getItem(OUTBOX) || "[]");
  const setOutbox = (q) => { localStorage.setItem(OUTBOX, JSON.stringify(q)); renderSync(); };
  function renderSync() {
    const n = outbox().length, el = $("#sync");
    el.textContent = n ? `${n} waiting to sync` : "";
    el.classList.toggle("pending", !!n);
  }
  async function flush() {
    let q = outbox();
    if (!q.length) return;
    while (q.length) {
      try { await S.create(q[0]); } catch (e) { if (!e.http) break; } // keep on network error, drop on 4xx
      q = q.slice(1); setOutbox(q);
    }
    if (!outbox().length) { toast("Offline fill-ups synced"); await load(); }
  }
  async function load() {
    try {
      const j = await S.list();
      data = j.fillups;
      services = j.services || [];
      appSettings = j.settings || {};
      setVehicle(j.vehicle);
      if (!LOCAL) localStorage.setItem("mpg-cache", JSON.stringify(j));
    } catch {
      const c = JSON.parse(localStorage.getItem("mpg-cache") || "null");
      if (c && !LOCAL) { data = c.fillups; services = c.services || []; appSettings = c.settings || {}; setVehicle(c.vehicle); toast("Offline: showing last copy"); }
      else if (LOCAL) toast("Couldn't open this phone's storage");
    }
    renderAll();
  }

  // ------------------------------------------------------------ vehicle, export, backup
  function setVehicle(v) {
    v = v || "My Truck";
    $("#vehicle").textContent = v;
    document.title = `${v} MPG`;
  }
  const fileName = () => `${($("#vehicle").textContent || "MPG").replace(/[^\w-]+/g, "_")}_MPG_${todayISO()}.xlsx`;
  let exportCache = null; // prebuilt so the iOS share sheet opens straight from the tap
  async function prepareExport() {
    exportCache = null;
    if (!LOCAL) return;
    try { exportCache = new File([await S.exportBlob()], fileName(), { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }); } catch {}
  }
  async function exportXlsx() {
    let file = exportCache;
    try {
      if (!file) file = new File([await S.exportBlob()], fileName(), { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
    } catch (err) { toast(`Export failed: ${err.message}`); return false; }
    if (navigator.canShare && navigator.canShare({ files: [file] })) {
      try { await navigator.share({ files: [file], title: file.name }); markBackup(); return true; }
      catch (err) { if (err.name === "AbortError") return false; } // otherwise fall through to download
    }
    const a = h("a", { href: URL.createObjectURL(file), download: file.name });
    document.body.append(a); a.click(); a.remove();
    setTimeout(() => URL.revokeObjectURL(a.href), 30000);
    markBackup();
    return true;
  }
  function markBackup() { if (LOCAL) { S.settings.lastBackup = new Date().toISOString(); renderSettings(); } }
  function offerBackup() {
    const bar = $("#backup-bar");
    bar.hidden = false;
    bar.classList.remove("on"); void bar.offsetWidth; bar.classList.add("on");
  }
  $("#backup-now").addEventListener("click", async () => { if (await exportXlsx()) $("#backup-bar").hidden = true; });
  $("#backup-later").addEventListener("click", () => { $("#backup-bar").hidden = true; });
  $$(".export").forEach((b) => b.addEventListener("click", exportXlsx));

  function ago(iso) {
    if (!iso) return "never";
    const d = Math.floor((Date.now() - new Date(iso)) / 864e5);
    return d <= 0 ? "today" : d === 1 ? "yesterday" : `${d} days ago`;
  }
  // ------------------------------------------------------------ palette
  const PALETTES = [
    { id: "amber", label: "Amber", bg: "#11110f", accent: "#f5a524" },
    { id: "ocean", label: "Ocean", bg: "#0d1117", accent: "#4aa8ff" },
    { id: "forest", label: "Forest", bg: "#0e1411", accent: "#7ddc8f" },
    { id: "ruby", label: "Ruby", bg: "#140f10", accent: "#ff6b81" },
    { id: "daylight", label: "Daylight", bg: "#f6f4ef", accent: "#9a5b00" },
    { id: "glacier", label: "Glacier", bg: "#f3f6fa", accent: "#1d5fbf" }
  ];
  function applyTheme(id, save) {
    const p = PALETTES.find((x) => x.id === id) || PALETTES[0];
    if (p.id === "amber") delete document.documentElement.dataset.theme;
    else document.documentElement.dataset.theme = p.id;
    $('meta[name="theme-color"]').setAttribute("content", p.bg);
    if (save) localStorage.setItem("mpg-theme", p.id);
    $$("#palettes .pal").forEach((b) => b.setAttribute("aria-checked", String(b.dataset.id === p.id)));
    if (save && !$("#view-stats").hidden) renderChart(inRange().filter((r) => r.dash_mpg)); // gradient takes the new accent
  }
  $("#palettes").replaceChildren(...PALETTES.map((p) => h("button", {
    type: "button", class: "pal", role: "radio", "data-id": p.id, "aria-checked": "false",
    onclick: () => applyTheme(p.id, true),
  }, h("span", { class: "sw", "aria-hidden": "true", style: `background:${p.bg}` }, h("i", { style: `background:${p.accent}` })), p.label)));
  applyTheme(localStorage.getItem("mpg-theme") || "amber", false);

  function renderSettings() {
    if (!LOCAL) return;
    const v = $("#set-vehicle");
    if (document.activeElement !== v) v.value = S.settings.vehicle;
    $("#set-remind").checked = S.settings.remind;
    const since = data.filter((r) => !S.settings.lastBackup || r.updated_at > S.settings.lastBackup.slice(0, 19)).length;
    $("#last-backup").textContent = `Last backup: ${ago(S.settings.lastBackup)}` + (since && S.settings.lastBackup ? ` · ${since} change${since > 1 ? "s" : ""} since` : "");
  }
  if (LOCAL) {
    $("#set-vehicle").addEventListener("change", (e) => {
      S.settings.vehicle = e.target.value.trim() || "My Truck";
      setVehicle(S.settings.vehicle); prepareExport();
    });
    $("#set-remind").addEventListener("change", (e) => { S.settings.remind = e.target.checked; });
    $("#wipe").addEventListener("click", async () => {
      if (!data.length) return;
      if (!confirm(`Delete all ${data.length} fill-ups from this phone? Back up first if you want to keep them.`)) return;
      if (!confirm("This can't be undone. Delete everything?")) return;
      await S.clear(); toast("All data deleted"); await load();
    });
    // first-run tip: Safari tabs can lose site data; a home-screen app keeps it
    const standalone = navigator.standalone || matchMedia("(display-mode: standalone)").matches;
    if (!standalone && !localStorage.getItem("mpg-tip-dismissed")) $("#local-tip").hidden = false;
    $("#tip-close").addEventListener("click", () => { $("#local-tip").hidden = true; localStorage.setItem("mpg-tip-dismissed", "1"); });
    if (navigator.storage && navigator.storage.persist) navigator.storage.persist().catch(() => {});
  }
  if (!S.cfg.scan) $("#scan").hidden = true;
  $$(".local-only").forEach((el) => (el.hidden = !LOCAL));
  $$(".server-only").forEach((el) => (el.hidden = LOCAL));
  setVehicle(LOCAL ? S.settings.vehicle : S.cfg.vehicle);

  // ------------------------------------------------------------ hero
  function renderHero() {
    const withMpg = data.filter((r) => r.dash_mpg);
    const last5 = withMpg.slice(-5), life = avgMpg(withMpg), cur = avgMpg(last5);
    $("#hero-mpg").textContent = f(cur);
    const d = $("#hero-delta");
    d.replaceChildren();
    if (cur != null && life != null) {
      const diff = cur - life;
      d.append(h("span", { class: diff >= 0 ? "up" : "down" }, `${diff >= 0 ? "+" : "−"}${f(Math.abs(diff))}`), ` vs ${f(life)} lifetime`);
    }
    const vals = withMpg.map((r) => r.dash_mpg).sort((a, b) => a - b);
    const lo = vals.length ? Math.floor(vals[Math.floor(vals.length * 0.05)]) - 1 : 14;
    const hi = vals.length ? Math.ceil(vals[Math.ceil(vals.length * 0.95) - 1]) + 1 : 22;
    $("#g-min").textContent = lo; $("#g-max").textContent = hi;
    const p = cur == null ? 0 : Math.max(0, Math.min(1, (cur - lo) / (hi - lo)));
    const ticks = $("#gauge-ticks");
    if (!ticks.childElementCount) {
      for (let i = 0; i <= 8; i++) {
        const a = Math.PI * (1 - i / 8), x = 60 + Math.cos(a) * 42, y = 64 - Math.sin(a) * 42, x2 = 60 + Math.cos(a) * 37, y2 = 64 - Math.sin(a) * 37;
        const l = document.createElementNS("http://www.w3.org/2000/svg", "line");
        l.setAttribute("x1", x); l.setAttribute("y1", y); l.setAttribute("x2", x2); l.setAttribute("y2", y2);
        ticks.append(l);
      }
    }
    requestAnimationFrame(() => {
      $("#gauge-fill").style.strokeDasharray = `${p * 100} 100`;
      $("#needle").style.transform = `rotate(${-90 + p * 180}deg)`;
    });
  }

  // ------------------------------------------------------------ form
  const form = $("#form");
  const F = (n) => form.elements[n];

  function prevFor(row) {
    if (!row) return data[data.length - 1] || null;
    const i = data.findIndex((r) => r.id === row.id);
    return i > 0 ? data[i - 1] : null;
  }
  function lastWith(key) { for (let i = data.length - 1; i >= 0; i--) if (data[i][key] != null && data[i][key] !== "") return data[i][key]; return null; }
  function lastOdometer(before) {
    const rows = before ? data.slice(0, data.findIndex((r) => r.id === before.id)) : data;
    for (let i = rows.length - 1; i >= 0; i--) if (rows[i].odometer) return rows[i];
    return null;
  }

  function seg(container, options, current, onPick, fmt = String) {
    container.replaceChildren(...options.map((o) => h("button", {
      type: "button", role: "radio", "aria-checked": String(o === current),
      onclick: () => { onPick(o); },
    }, fmt(o))));
  }
  function renderChoices() {
    const octs = [...new Set([...BASE_OCTANE, ...data.map((r) => r.octane).filter(Boolean)])].sort();
    seg($("#octane"), octs, st.octane, (o) => { st.octane = o; renderChoices(); });

    const modes = freq(data.map((r) => r.drive_mode));
    const names = [...modes.map((m) => m.k), ...BASE_MODES.filter((b) => !modes.some((m) => m.k === b))];
    const otherOn = st.mode === "__other";
    $("#mode").replaceChildren(
      ...names.map((n) => h("button", {
        type: "button", class: "chip", role: "radio", "aria-checked": String(st.mode === n),
        onclick: () => { st.mode = n; F("drive_mode_other").hidden = true; renderChoices(); },
      }, n)),
      h("button", { type: "button", class: "chip", role: "radio", "aria-checked": String(otherOn),
        onclick: () => { st.mode = "__other"; renderChoices(); F("drive_mode_other").hidden = false; F("drive_mode_other").focus(); } }, "Other…"),
    );

    const used = freq(data.flatMap(tagsOf));
    const tags = [...new Set([...BASE_TAGS, ...used.map((u) => u.k), ...st.tags])];
    const count = Object.fromEntries(used.map((u) => [u.k, u.c]));
    $("#tags").replaceChildren(...tags.map((t) => h("button", {
      type: "button", class: "chip", "aria-pressed": String(st.tags.has(t)),
      onclick: () => { st.tags.has(t) ? st.tags.delete(t) : st.tags.add(t); renderChoices(); },
    }, t, count[t] ? h("span", { class: "ct" }, String(count[t])) : null)));

    $("#station-list").replaceChildren(...freq(data.map((r) => r.station)).map((s) => h("option", { value: s.k })));
  }

  function setAuto(name, value) {
    const el = F(name), badge = $(`[data-auto="${name}"]`);
    if (st.user[name]) return;
    el.value = value == null ? "" : value;
    el.classList.toggle("is-auto", value != null && value !== "");
    badge.hidden = value == null || value === "";
  }
  function userEdit(name) {
    st.user[name] = F(name).value.trim() !== "";
    F(name).classList.remove("is-auto", "is-photo");
    const b = $(`[data-auto="${name}"]`);
    if (b) b.hidden = true;
  }

  function fillForm(row) {
    editing = row;
    form.reset();
    st.user = {};
    $$("[data-auto]").forEach((b) => (b.hidden = true));
    $$("input.is-auto").forEach((i) => i.classList.remove("is-auto"));
    clearPhotoMarks();
    $("#scan-thumb").hidden = true;
    $("#scan-title").textContent = "Scan a photo";
    $("#scan-sub").textContent = "Dash trip screen, pump or receipt";
    const prev = prevFor(row), lo = lastOdometer(row);
    if (row) {
      for (const k of ["date", "odometer", "trip_miles", "dash_mpg", "gallons", "price_per_gal", "total_cost", "station", "notes"]) F(k).value = row[k] ?? "";
      F("drive_time").value = hm(row.drive_minutes);
      for (const k of ["odometer", "trip_miles", "total_cost"]) st.user[k] = row[k] != null;
      st.octane = row.octane; st.mode = row.drive_mode; st.tags = new Set(tagsOf(row));
      $("#log-title").textContent = `Edit fill-up · ${fmtDate(row.date, { month: "short", day: "numeric", year: "numeric" })}`;
      $("#save span").textContent = "Save changes";
    } else {
      F("date").value = todayISO();
      st.octane = lastWith("octane") ?? 87;
      st.mode = lastWith("drive_mode") ?? "Normal";
      st.tags = new Set();
      F("station").value = lastWith("station") ?? "";
      $("#log-title").textContent = `New fill-up`;
      $("#save span").textContent = "Save fill-up";
      dirty = false;
      if (restoreDraft()) $("#log-title").textContent = "New fill-up (draft)";
    }
    const known = new Set(freq(data.map((r) => r.drive_mode)).map((m) => m.k).concat(BASE_MODES));
    if (st.mode && !known.has(st.mode)) { F("drive_mode_other").value = st.mode; st.mode = "__other"; }
    F("drive_mode_other").hidden = st.mode !== "__other";

    // placeholders from history
    const trips = data.map((r) => r.trip_miles).filter(Boolean).sort((a, b) => a - b);
    const medTrip = trips.length ? trips[Math.floor(trips.length / 2)] : null;
    F("odometer").placeholder = lo && medTrip ? `≈ ${f(lo.odometer + medTrip, 0)}` : "";
    F("trip_miles").placeholder = medTrip ? f(medTrip, 0) : "";
    const gals = data.map((r) => r.gallons).filter(Boolean);
    F("gallons").placeholder = gals.length ? f(gals[gals.length - 1], 1) : "";
    F("dash_mpg").placeholder = f(avgMpg(data.slice(-5)));
    const lp = lastWith("price_per_gal");
    F("price_per_gal").placeholder = lp ? f(lp, 2) : "";

    $("#last-line").textContent = prev
      ? `Last: ${fmtDate(prev.date)} · ${prev.odometer ? f(prev.odometer, 1) + " mi" : f(prev.trip_miles, 1) + " trip mi"} · ${f(prev.dash_mpg)} mpg`
      : "First fill-up";
    $("#cancel-edit").hidden = $("#delete").hidden = !row;
    $("#cancel-edit").textContent = row ? "Cancel" : "Clear draft";
    if (!row && dirty) $("#cancel-edit").hidden = false;
    $("#form-err").textContent = "";
    renderChoices();
    preview();
  }

  function preview() {
    const trip = num(F("trip_miles").value), gal = num(F("gallons").value), dash = num(F("dash_mpg").value);
    const calc = trip && gal ? trip / gal : null;
    $("#pv-calc").textContent = calc ? `${f(calc)} mpg` : "--";
    const avg = avgMpg(editing ? data.filter((r) => r.id !== editing.id) : data);
    const mine = dash ?? calc;
    const vs = $("#pv-vs");
    vs.replaceChildren();
    if (mine && avg) { const d = mine - avg; vs.append(h("span", { class: d >= 0 ? "up" : "down" }, `${d >= 0 ? "+" : "−"}${f(Math.abs(d))}`)); }
    else vs.textContent = "--";
    const prev = prevFor(editing), date = F("date").value;
    const dd = prev && date ? days(prev.date, date) : null;
    $("#pv-days").textContent = trip && dd > 0 ? f(trip / dd, 0) : dd === 0 ? "same day" : dd != null ? `${dd} d` : "--";
    const price = num(F("price_per_gal").value);
    if (price && gal) setAuto("total_cost", (price * gal).toFixed(2));
    const tot = num(F("total_cost").value);
    $("#cost-sum").textContent = tot ? `$${f(tot, 2)}${F("station").value ? " · " + F("station").value : ""}` : F("station").value || "optional";
  }

  function saveDraft() {
    if (editing) return;
    dirty = true;
    const vals = Object.fromEntries([...form.elements].filter((el) => el.name).map((el) => [el.name, el.value]));
    localStorage.setItem(DRAFT, JSON.stringify({ at: Date.now(), vals, user: st.user, octane: st.octane, mode: st.mode, tags: [...st.tags] }));
  }
  function restoreDraft() {
    const d = JSON.parse(localStorage.getItem(DRAFT) || "null");
    if (!d || Date.now() - d.at > 12 * 36e5) return false;
    for (const [k, v] of Object.entries(d.vals)) if (form.elements[k]) form.elements[k].value = v;
    st.user = d.user || {}; st.octane = d.octane; st.mode = d.mode; st.tags = new Set(d.tags || []);
    F("drive_mode_other").hidden = st.mode !== "__other";
    dirty = true;
    return true;
  }
  const clearDraft = () => { localStorage.removeItem(DRAFT); dirty = false; };
  form.addEventListener("click", (e) => { if (e.target.closest("button[type=button]") && !e.target.closest(".actions")) setTimeout(() => { saveDraft(); if (!editing) $("#cancel-edit").hidden = false; }); });

  form.addEventListener("input", (e) => {
    setTimeout(() => { saveDraft(); if (!editing) { $("#cancel-edit").hidden = false; $("#cancel-edit").textContent = "Clear draft"; } });
    const n = e.target.name;
    const lo = lastOdometer(editing);
    if (n === "odometer") {
      userEdit("odometer");
      const odo = num(F("odometer").value);
      if (lo && odo && odo > lo.odometer) setAuto("trip_miles", (odo - lo.odometer).toFixed(1));
      else if (!odo) setAuto("trip_miles", null);
    } else if (n === "trip_miles") {
      userEdit("trip_miles");
      const trip = num(F("trip_miles").value);
      if (lo && trip) setAuto("odometer", (lo.odometer + trip).toFixed(1));
      else if (!trip) setAuto("odometer", null);
    } else if (n === "total_cost") userEdit("total_cost");
    preview();
  });
  F("drive_time").addEventListener("blur", () => {
    const m = parseHM(F("drive_time").value);
    if (Number.isFinite(m)) F("drive_time").value = hm(m);
  });

  function collect() {
    const mins = parseHM(F("drive_time").value);
    if (Number.isNaN(mins)) throw new Error("Drive time should look like 12:34");
    const mode = st.mode === "__other" ? F("drive_mode_other").value.trim() : st.mode;
    const out = {
      date: F("date").value,
      odometer: num(F("odometer").value), trip_miles: num(F("trip_miles").value),
      dash_mpg: num(F("dash_mpg").value), gallons: num(F("gallons").value),
      drive_minutes: mins, octane: st.octane, drive_mode: mode || null,
      price_per_gal: num(F("price_per_gal").value), total_cost: num(F("total_cost").value),
      station: F("station").value.trim() || null,
      tags: [...st.tags].join(", ") || null, notes: F("notes").value.trim() || null,
    };
    for (const k of ["odometer", "trip_miles", "dash_mpg", "gallons", "price_per_gal", "total_cost"]) {
      const raw = F(k).value.trim();
      if (raw && out[k] == null) throw new Error(`${F(k).closest(".field").querySelector(".lbl").firstChild.textContent.trim()} isn't a number`);
    }
    if (!out.date) throw new Error("Pick a date");
    if (!out.dash_mpg && !(out.trip_miles && out.gallons)) throw new Error("Enter dash MPG, or trip miles and gallons");
    if (out.dash_mpg && (out.dash_mpg < 3 || out.dash_mpg > 60)) throw new Error("Dash MPG looks off");
    const lo = lastOdometer(editing);
    if (!editing && out.odometer && lo && out.odometer < lo.odometer &&
        !confirm(`Odometer ${f(out.odometer)} is lower than last time (${f(lo.odometer)}). Save anyway?`)) throw new Error("");
    return out;
  }

  form.addEventListener("submit", async (e) => {
    e.preventDefault();
    let body;
    try { body = collect(); } catch (err) { $("#form-err").textContent = err.message; return; }
    $("#form-err").textContent = "";
    const btn = $("#save");
    btn.disabled = true; btn.classList.add("saving");
    try {
      if (editing) { await S.update(editing.id, body); toast("Changes saved"); }
      else {
        const before = avgMpg(data);
        await S.create(body);
        const mpg = body.dash_mpg ?? body.trip_miles / body.gallons;
        toast(before ? `Saved · ${f(mpg)} mpg, ${mpg >= before ? "+" : "−"}${f(Math.abs(mpg - before))} vs avg` : "Saved");
      }
      clearDraft();
      const wasNew = !editing;
      await load();
      fillForm(null);
      if (LOCAL && wasNew && S.settings.remind) offerBackup();
      if (editing === null && body) window.scrollTo({ top: 0, behavior: reduced ? "auto" : "smooth" });
    } catch (err) {
      if (!LOCAL && !err.http && !editing) { setOutbox([...outbox(), body]); clearDraft(); toast("No connection: saved on phone, will sync"); fillForm(null); }
      else $("#form-err").textContent = err.message;
    } finally {
      setTimeout(() => { btn.disabled = false; btn.classList.remove("saving"); }, 350);
    }
  });
  $("#cancel-edit").addEventListener("click", () => {
    if (!editing) { clearDraft(); fillForm(null); return; }
    fillForm(null); show("history");
  });
  $("#delete").addEventListener("click", async () => {
    if (!editing || !confirm(`Delete the ${fmtDate(editing.date)} fill-up?`)) return;
    try { await S.remove(editing.id); toast("Deleted"); await load(); fillForm(null); show("history"); }
    catch (err) { $("#form-err").textContent = err.message; }
  });

  // ------------------------------------------------------------ photo scan
  function shrink(file, max = 1600) {
    return new Promise((resolve, reject) => {
      const img = new Image(), url = URL.createObjectURL(file);
      img.onload = () => {
        const k = Math.min(1, max / Math.max(img.naturalWidth, img.naturalHeight));
        const c = document.createElement("canvas");
        c.width = Math.round(img.naturalWidth * k); c.height = Math.round(img.naturalHeight * k);
        c.getContext("2d").drawImage(img, 0, 0, c.width, c.height);
        URL.revokeObjectURL(url);
        c.toBlob((b) => (b ? resolve(b) : reject(new Error("Could not read photo"))), "image/jpeg", 0.86);
      };
      img.onerror = () => { URL.revokeObjectURL(url); reject(new Error("Could not read photo")); };
      img.src = url;
    });
  }
  function markPhoto(name) {
    const el = F(name);
    st.user[name] = true;
    el.classList.remove("is-auto", "is-photo"); void el.offsetWidth; el.classList.add("is-photo");
    const lbl = el.closest(".field").querySelector(".lbl");
    let b = lbl.querySelector(".auto");
    if (!b) { b = h("span", { class: "auto", "data-photo": "" }); lbl.append(b); }
    b.textContent = "photo"; b.classList.add("photo"); b.hidden = false;
  }
  function clearPhotoMarks() {
    $$("input.is-photo").forEach((i) => i.classList.remove("is-photo"));
    $$(".auto.photo").forEach((b) => { b.classList.remove("photo"); b.textContent = "auto"; b.hidden = true; });
  }
  $("#scan-file").addEventListener("change", async (e) => {
    const file = e.target.files[0];
    e.target.value = "";
    if (!file) return;
    const scan = $("#scan"), thumb = $("#scan-thumb");
    thumb.src = URL.createObjectURL(file); thumb.hidden = false; thumb.alt = "Scanned photo";
    scan.classList.add("busy");
    $("#scan-title").textContent = "Reading photo…";
    $("#scan-sub").textContent = "First scan can take ~40 s";
    try {
      const r = await S.scan(await shrink(file).catch(() => file));
      const got = [];
      const put = (name, v, label) => { if (v == null) return; F(name).value = v; markPhoto(name); got.push(label); };
      put("odometer", r.odometer != null ? r.odometer.toFixed(1) : null, "odometer");
      put("trip_miles", r.trip_miles != null ? r.trip_miles.toFixed(1) : null, "trip");
      put("dash_mpg", r.dash_mpg != null ? r.dash_mpg.toFixed(1) : null, "MPG");
      put("drive_time", r.drive_minutes != null ? hm(r.drive_minutes) : null, "time");
      put("gallons", r.gallons != null ? r.gallons.toFixed(3).replace(/0+$/, "") : null, "gallons");
      put("price_per_gal", r.price_per_gal != null ? r.price_per_gal.toFixed(3).replace(/0$/, "") : null, "$/gal");
      put("total_cost", r.total_cost != null ? r.total_cost.toFixed(2) : null, "total");
      if (r.price_per_gal || r.total_cost) $("#more").open = true;
      if (r.octane) { st.octane = r.octane; renderChoices(); got.push("octane"); }
      // dash odometer only: derive trip from last fill-up, and vice versa
      const lo = lastOdometer(editing);
      if (r.odometer && r.trip_miles == null && lo && r.odometer > lo.odometer) { st.user.trip_miles = false; setAuto("trip_miles", (r.odometer - lo.odometer).toFixed(1)); }
      const warn = lo && r.odometer && r.trip_miles && Math.abs(r.odometer - lo.odometer - r.trip_miles) > 25
        ? ` Trip (${f(r.trip_miles)}) doesn't match odometer change (${f(r.odometer - lo.odometer)}).` : "";
      $("#scan-title").textContent = got.length ? `Filled ${got.join(", ")}` : "Couldn't read any values";
      $("#scan-sub").textContent = got.length ? "Check them before saving" + (warn ? " ·" + warn : "") : "Try again closer, less glare";
      preview(); saveDraft();
      $("#cancel-edit").hidden = false;
    } catch (err) {
      $("#scan-title").textContent = "Scan failed";
      $("#scan-sub").textContent = err.message;
    } finally {
      scan.classList.remove("busy");
    }
  });

  // ------------------------------------------------------------ history
  function renderHistory() {
    const q = $("#search").value.trim().toLowerCase();
    const list = $("#hist-list");
    const rows = data.slice().reverse().filter((r) => !q ||
      [r.notes, r.drive_mode, r.tags, r.station, r.date, String(r.octane)].some((x) => x && String(x).toLowerCase().includes(q)));
    if (!rows.length) { list.replaceChildren(h("p", { class: "empty" }, data.length ? "Nothing matches." : "No fill-ups yet.")); return; }
    const mpgs = data.map((r) => r.dash_mpg).filter(Boolean);
    const lo = Math.min(...mpgs) - 1, hi = Math.max(...mpgs), avg = avgMpg(data);
    const out = []; let month = null, i = 0, group = [];
    const flushMonth = () => {
      if (!group.length) return;
      const a = avgMpg(group), mi = group.reduce((s, r) => s + (r.trip_miles || 0), 0);
      out.push(h("div", { class: "month" }, h("h2", {}, fmtDate(group[0].date, { month: "long", year: "numeric" })),
        h("span", {}, `${group.length} fill${group.length > 1 ? "s" : ""} · ${f(mi, 0)} mi · ${f(a)} avg`)));
      out.push(...group.map((r) => entry(r, i++, lo, hi, avg)));
      group = [];
    };
    for (const r of rows) {
      const m = r.date.slice(0, 7);
      if (m !== month) { flushMonth(); month = m; }
      group.push(r);
    }
    flushMonth();
    list.replaceChildren(...out);
  }
  function entry(r, i, lo, hi, avg) {
    const d = r.dash_mpg && avg ? r.dash_mpg - avg : null;
    const meta = [r.trip_miles && `${f(r.trip_miles)} mi`, r.gallons && `${f(r.gallons, 2)} gal`, r.drive_minutes != null && hm(r.drive_minutes), r.octane, r.drive_mode].filter(Boolean).join(" · ");
    const note = [r.notes, r.station].filter(Boolean).join(" · ");
    return h("button", {
      type: "button", class: "entry", style: `--i:${Math.min(i, 12)}`,
      "aria-label": `${fmtDate(r.date, { month: "long", day: "numeric", year: "numeric" })}, ${f(r.dash_mpg)} mpg. Edit`,
      onclick: () => { fillForm(r); show("log"); },
    },
    h("span", { class: "e-date" }, h("b", {}, String(toDate(r.date).getDate())), toDate(r.date).toLocaleDateString("en-US", { weekday: "short" })),
    h("span", { class: "e-mid" },
      h("span", { class: "e-meta" }, meta, ...tagsOf(r).map((t) => h("span", { class: "tag" }, t))),
      note ? h("span", { class: "e-note", style: "display:block" }, note) : null),
    h("span", { class: "e-mpg" }, f(r.dash_mpg ?? calcMpg(r)),
      d != null ? h("small", { class: d >= 0 ? "up" : "down" }, `${d >= 0 ? "+" : "−"}${f(Math.abs(d))}`) : h("small", {}, "mpg")),
    h("span", { class: "e-bar", "aria-hidden": "true" }, h("i", { style: `width:${r.dash_mpg ? Math.max(4, ((r.dash_mpg - lo) / (hi - lo)) * 100) : 0}%;--i:${Math.min(i, 12)}` })));
  }
  $("#search").addEventListener("input", () => { cancelAnimationFrame(renderHistory.raf); renderHistory.raf = requestAnimationFrame(renderHistory); });

  // ------------------------------------------------------------ stats
  function inRange() {
    if (range === "all") return data;
    const cut = new Date(); cut.setMonth(cut.getMonth() - (range === "1y" ? 12 : 6));
    const iso = cut.toISOString().slice(0, 10);
    return data.filter((r) => r.date >= iso);
  }
  function renderStats() {
    seg($("#range"), ["6m", "1y", "all"], range, (r) => { range = r; renderStats(); }, (r) => r.toUpperCase());
    const rows = inRange(), withMpg = rows.filter((r) => r.dash_mpg);
    const pump = rows.filter((r) => r.trip_miles && r.gallons);
    const pmi = pump.reduce((s, r) => s + r.trip_miles, 0), pgal = pump.reduce((s, r) => s + r.gallons, 0);
    const pumpMpg = pgal ? pmi / pgal : null;
    const dashOnPump = avgMpg(pump);
    const best = withMpg.reduce((a, r) => (!a || r.dash_mpg > a.dash_mpg ? r : a), null);
    const worst = withMpg.reduce((a, r) => (!a || r.dash_mpg < a.dash_mpg ? r : a), null);
    const odo = rows.filter((r) => r.odometer);
    const miles = odo.length > 1 ? odo[odo.length - 1].odometer - odo[0].odometer : rows.reduce((s, r) => s + (r.trip_miles || 0), 0);
    const gallons = rows.reduce((s, r) => s + (r.gallons || 0), 0);
    const trips = rows.map((r) => r.trip_miles).filter(Boolean);
    const gaps = rows.slice(1).map((r, i) => days(rows[i].date, r.date)).filter((d) => d > 0);
    const spent = rows.reduce((s, r) => s + (r.total_cost || 0), 0);
    const tile = (label, v, unit, sub, wide) => h("div", { class: "tile" + (wide ? " wide" : "") },
      h("p", { class: "eyebrow" }, label), h("p", { class: "v" }, v, unit ? h("small", {}, unit) : null), sub ? h("p", { class: "s" }, sub) : null);
    const tiles = [
      tile("Average", f(avgMpg(withMpg)), "mpg", `${withMpg.length} tanks, dash`),
      tile("Pump math", f(pumpMpg), "mpg", pumpMpg && dashOnPump ? `dash reads ${f(((dashOnPump - pumpMpg) / pumpMpg) * 100, 0)}% high` : "miles ÷ gallons"),
      tile("Best tank", f(best?.dash_mpg), "mpg", best ? fmtDate(best.date, { month: "short", day: "numeric", year: "numeric" }) : ""),
      tile("Worst tank", f(worst?.dash_mpg), "mpg", worst ? (worst.notes ? worst.notes.split(".")[0] : fmtDate(worst.date, { month: "short", day: "numeric", year: "numeric" })).slice(0, 48) : ""),
      tile("Miles", f(miles, 0), "mi", odo.length > 1 ? "by odometer" : "sum of trips"),
      tile("Fuel", f(gallons, 0), "gal", `${rows.filter((r) => r.gallons).length} tanks logged`),
      tile("Avg tank", f(trips.reduce((a, b) => a + b, 0) / (trips.length || 1), 0), "mi", trips.length ? `range ${f(Math.min(...trips), 0)}–${f(Math.max(...trips), 0)}` : ""),
      tile("Between fills", f(gaps.reduce((a, b) => a + b, 0) / (gaps.length || 1), 0), "days", gaps.length ? `${f(Math.min(...gaps), 0)}–${f(Math.max(...gaps), 0)} days` : ""),
    ];
    if (spent) tiles.push(tile("Fuel spent", `$${f(spent, 0)}`, "", `${rows.filter((r) => r.total_cost).length} receipts`));
    const svcIn = rows.length ? services.filter((r) => r.date >= rows[0].date) : [];
    const svcCost = svcIn.reduce((a, r) => a + (r.cost || 0), 0);
    if (svcIn.length) tiles.push(tile("Maintenance", `$${f(svcCost, 0)}`, "", miles && svcCost ? `${f((svcCost / miles) * 100, 1)}¢/mi · ${svcIn.length} services` : `${svcIn.length} services`, !spent));
    $("#tiles").replaceChildren(...tiles);
    $("#tiles").classList.remove("stagger"); void $("#tiles").offsetWidth; $("#tiles").classList.add("stagger");
    tiles.forEach((t, i) => t.style.setProperty("--i", i));

    const groups = [
      ...freq(rows.map((r) => r.drive_mode)).map(({ k }) => ({ k, rows: rows.filter((r) => r.drive_mode === k) })),
      ...freq(rows.flatMap(tagsOf)).map(({ k }) => ({ k: `#${k}`, rows: rows.filter((r) => tagsOf(r).includes(k)) })),
    ];
    $("#modes").replaceChildren(
      h("thead", {}, h("tr", {}, h("th", { scope: "col" }, "Mode / tag"), h("th", { scope: "col" }, "Tanks"), h("th", { scope: "col" }, "Avg mpg"))),
      h("tbody", {}, ...groups.map((g) => h("tr", {}, h("td", {}, g.k), h("td", {}, String(g.rows.length)), h("td", {}, f(avgMpg(g.rows)))))));
    renderChart(withMpg);
  }

  function renderChart(rows) {
    const box = $("#chart");
    box.replaceChildren();
    if (rows.length < 2) { box.append(h("p", { class: "empty" }, "Need two fill-ups for a chart.")); return; }
    const NS = "http://www.w3.org/2000/svg";
    const s = (tag, attrs = {}) => { const el = document.createElementNS(NS, tag); for (const [k, v] of Object.entries(attrs)) el.setAttribute(k, v); return el; };
    const W = Math.max(280, box.clientWidth || 320), H = window.innerWidth >= 768 ? 280 : 220, pl = 28, pr = 8, pt = 8, pb = 22;
    const t = rows.map((r) => toDate(r.date).getTime()), v = rows.map((r) => r.dash_mpg);
    const avg = v.map((_, i) => { const w = v.slice(Math.max(0, i - 4), i + 1); return w.reduce((a, b) => a + b, 0) / w.length; });
    const t0 = t[0], t1 = t[t.length - 1];
    const y0 = Math.floor(Math.min(...v)) - 1, y1 = Math.ceil(Math.max(...v)) + 1;
    const X = (x) => pl + ((x - t0) / (t1 - t0 || 1)) * (W - pl - pr), Y = (y) => pt + (1 - (y - y0) / (y1 - y0)) * (H - pt - pb);
    const svg = s("svg", { viewBox: `0 0 ${W} ${H}`, role: "img", "aria-label": `MPG per tank, ${rows.length} tanks from ${f(Math.min(...v))} to ${f(Math.max(...v))}, latest ${f(v[v.length - 1])}` });
    const defs = s("defs"), lg = s("linearGradient", { id: "accentfade", x1: 0, y1: 0, x2: 0, y2: 1 });
    const accent = getComputedStyle(document.documentElement).getPropertyValue("--accent").trim() || "#f5a524";
    lg.append(s("stop", { offset: "0", "stop-color": accent, "stop-opacity": ".22" }), s("stop", { offset: "1", "stop-color": accent, "stop-opacity": "0" }));
    defs.append(lg); svg.append(defs);
    const grid = s("g", { class: "grid" }), axis = s("g", { class: "axis" });
    const step = y1 - y0 > 8 ? 2 : 1;
    for (let y = Math.ceil(y0 / step) * step; y <= y1; y += step) {
      grid.append(s("line", { x1: pl, x2: W - pr, y1: Y(y), y2: Y(y) }));
      const tx = s("text", { x: pl - 6, y: Y(y) + 3, "text-anchor": "end" }); tx.textContent = y; axis.append(tx);
    }
    // x ticks: quarter starts (or months for short ranges)
    const span = (t1 - t0) / 864e5, mStep = span > 500 ? 6 : span > 200 ? 3 : 1;
    const d = new Date(t0); d.setDate(1); d.setMonth(d.getMonth() + 1);
    while (d.getTime() <= t1) {
      if (d.getMonth() % mStep === 0) {
        const tx = s("text", { x: X(d.getTime()), y: H - 6, "text-anchor": "middle" });
        tx.textContent = d.getMonth() === 0 ? String(d.getFullYear()) : d.toLocaleDateString("en-US", { month: "short" });
        axis.append(tx);
      }
      d.setMonth(d.getMonth() + 1);
    }
    svg.append(grid, axis);
    const pts = t.map((x, i) => [X(x), Y(v[i])]);
    const path = "M" + pts.map((p) => p.map((n) => n.toFixed(1)).join(",")).join("L");
    svg.append(s("path", { class: "area", d: `${path}L${pts[pts.length - 1][0]},${Y(y0)}L${pts[0][0]},${Y(y0)}Z` }));
    const avgPath = s("path", { class: "avg", d: "M" + t.map((x, i) => `${X(x).toFixed(1)},${Y(avg[i]).toFixed(1)}`).join("L") });
    const line = s("path", { class: "ln", d: path });
    svg.append(avgPath, line);
    const hair = s("line", { class: "hair", y1: pt, y2: H - pb, opacity: 0 });
    const dot = s("circle", { class: "dot", r: 5, cx: pts[pts.length - 1][0], cy: pts[pts.length - 1][1] });
    const ring = s("circle", { class: "dot", r: 4, cx: pts[pts.length - 1][0], cy: pts[pts.length - 1][1] });
    // service events: small ticks on the baseline
    for (const sv of services) {
      const x = toDate(sv.date).getTime();
      if (x < t0 || x > t1) continue;
      const m = s("path", { class: "svc-mark", d: `M${X(x).toFixed(1)},${(H - pb - 1).toFixed(1)} l-4,-7 h8 z` });
      svg.append(m);
    }
    svg.append(hair, ring, dot);
    const hit = s("rect", { x: 0, y: 0, width: W, height: H, fill: "transparent", tabindex: 0, "aria-label": "Chart: use arrow keys to read each tank" });
    svg.append(hit);
    box.append(svg);
    if (!reduced) { line.style.setProperty("--len", Math.ceil(line.getTotalLength())); line.classList.add("draw"); }

    const tip = h("div", { class: "tip", "aria-hidden": "true" });
    box.append(tip);
    let cur = rows.length - 1, raf = 0, px = null;
    const showAt = (i) => {
      cur = Math.max(0, Math.min(rows.length - 1, i));
      const [x, y] = pts[cur], r = rows[cur];
      hair.setAttribute("x1", x); hair.setAttribute("x2", x); hair.setAttribute("opacity", 1);
      dot.setAttribute("cx", x); dot.setAttribute("cy", y);
      tip.replaceChildren(...[
        h("div", {}, fmtDate(r.date, { month: "short", day: "numeric", year: "numeric" })),
        h("div", { class: "row" }, h("i"), h("b", {}, f(r.dash_mpg)), " dash"),
        h("div", { class: "row" }, h("i", { class: "d" }), f(avg[cur]), " 5-tank avg"),
        r.notes ? h("div", { style: "max-width:200px;white-space:normal;color:var(--muted)" }, r.notes) : null,
        ...services.filter((sv) => sv.date <= r.date && (cur === 0 || sv.date > rows[cur - 1].date))
          .map((sv) => h("div", { style: "color:var(--text)" }, `▲ ${sv.category} ${fmtDate(sv.date)}`)),
      ].filter(Boolean));
      tip.classList.add("on");
      const scale = box.clientWidth / W, tw = tip.offsetWidth;
      let left = x * scale + 12; if (left + tw > box.clientWidth) left = x * scale - tw - 12;
      tip.style.left = `${Math.max(0, left)}px`;
    };
    const hide = () => { tip.classList.remove("on"); hair.setAttribute("opacity", 0); const [x, y] = pts[pts.length - 1]; dot.setAttribute("cx", x); dot.setAttribute("cy", y); };
    const nearest = (cx) => { const bx = box.getBoundingClientRect(), x = ((cx - bx.left) / bx.width) * W; let bi = 0; pts.forEach((p, i) => { if (Math.abs(p[0] - x) < Math.abs(pts[bi][0] - x)) bi = i; }); return bi; };
    hit.addEventListener("pointermove", (e) => { px = e.clientX; if (!raf) raf = requestAnimationFrame(() => { raf = 0; showAt(nearest(px)); }); });
    hit.addEventListener("pointerdown", (e) => showAt(nearest(e.clientX)));
    hit.addEventListener("pointerleave", hide);
    hit.addEventListener("blur", hide);
    hit.addEventListener("focus", () => showAt(cur));
    hit.addEventListener("keydown", (e) => {
      if (e.key === "ArrowLeft") { showAt(cur - 1); e.preventDefault(); }
      if (e.key === "ArrowRight") { showAt(cur + 1); e.preventDefault(); }
    });
  }

  // ------------------------------------------------------------ service
  const intervals = () => ({ ...SERVICE_DEFAULTS, ...(appSettings.intervals || {}) });
  // Odometer readings from fill-ups and services, oldest first
  function readings() {
    return [...data, ...services].filter((r) => r.odometer).map((r) => ({ date: r.date, odo: r.odometer }))
      .sort((a, b) => (a.date < b.date ? -1 : a.date > b.date ? 1 : a.odo - b.odo));
  }
  // Recent miles per day: odometer change over (up to) the last 180 days of readings
  function milesPerDay() {
    const rs = readings();
    if (rs.length < 2) return null;
    const last = rs[rs.length - 1];
    const cut = new Date(toDate(last.date) - 180 * 864e5).toISOString().slice(0, 10);
    const first = rs.find((r) => r.date >= cut && r !== last) || rs[0];
    const d = days(first.date, last.date);
    return d > 0 && last.odo > first.odo ? (last.odo - first.odo) / d : null;
  }
  function estOdometer(onDate = todayISO()) {
    const rs = readings();
    if (!rs.length) return null;
    const before = rs.filter((r) => r.date <= onDate), after = rs.filter((r) => r.date > onDate);
    if (before.length && after.length) { // interpolate between surrounding readings
      const a = before[before.length - 1], b = after[0];
      return a.odo + ((b.odo - a.odo) * days(a.date, onDate)) / Math.max(1, days(a.date, b.date));
    }
    const last = before.length ? before[before.length - 1] : rs[0], rate = milesPerDay() || 0;
    return last.odo + rate * Math.max(0, days(last.date, onDate));
  }
  function dueList() {
    const est = estOdometer(), rate = milesPerDay(), iv = intervals(), out = [], missing = [];
    for (const [cat, miles] of Object.entries(iv)) {
      if (!miles) continue;
      const done = services.filter((r) => r.category === cat);
      if (!done.length) { missing.push(cat); continue; }
      const last = done[done.length - 1], lastOdo = last.odometer ?? estOdometer(last.date);
      if (lastOdo == null || est == null) continue;
      const left = lastOdo + miles - est, used = (est - lastOdo) / miles;
      const status = left < 0 ? "over" : left < Math.max(500, miles * 0.1) ? "soon" : "ok";
      const eta = rate && left > 0 ? new Date(Date.now() + (left / rate) * 864e5) : null;
      out.push({ cat, miles, last, lastOdo, left, used, status, eta });
    }
    out.sort((a, b) => a.left - b.left);
    return { items: out, missing, est, rate };
  }
  const wrench = () => {
    const svg = document.createElementNS("http://www.w3.org/2000/svg", "svg");
    svg.setAttribute("viewBox", "0 0 24 24"); svg.setAttribute("aria-hidden", "true");
    const p = document.createElementNS("http://www.w3.org/2000/svg", "path");
    p.setAttribute("d", "M14.7 6.3a4 4 0 0 0-5.4 5.1L3.6 17.1a1.9 1.9 0 0 0 2.7 2.7l5.7-5.7a4 4 0 0 0 5.1-5.4l-2.5 2.5-2.3-.4-.4-2.3z");
    svg.append(p); return svg;
  };
  const plural = (n, w) => `${n} ${w}${n === 1 ? "" : "s"}`;
  const dueText = (d) => d.left < 0 ? `Overdue by ~${f(-d.left, 0)} mi` : `Due in ~${f(d.left, 0)} mi`;

  function renderDueBanner() {
    const b = $("#due-banner"), { items } = dueList();
    const urgent = items.filter((d) => d.status !== "ok");
    b.hidden = !urgent.length;
    if (!urgent.length) return;
    const top = urgent[0];
    b.className = "due-banner" + (top.status === "over" ? " over" : "");
    b.replaceChildren(wrench(), h("span", {}, `${top.cat}: ${dueText(top).toLowerCase()}`,
      urgent.length > 1 ? ` · +${urgent.length - 1} more` : ""));
  }
  $("#due-banner").addEventListener("click", () => { show("service"); window.scrollTo({ top: 0 }); });

  const svcForm = $("#svc-form");
  const G = (n) => svcForm.elements[n];
  let svcEditing = null, svcCat = null;

  function svcCategories() {
    const used = freq(services.map((r) => r.category));
    return [...new Set([...used.map((u) => u.k), ...Object.keys(SERVICE_DEFAULTS)])];
  }
  function renderSvcCats() {
    const cats = svcCategories(), other = svcCat === "__other";
    $("#svc-cats").replaceChildren(
      ...cats.map((c) => h("button", { type: "button", class: "chip", role: "radio", "aria-checked": String(svcCat === c),
        onclick: () => { svcCat = c; G("category_other").hidden = true; renderSvcCats(); svcAutofill(); } }, c)),
      h("button", { type: "button", class: "chip", role: "radio", "aria-checked": String(other),
        onclick: () => { svcCat = "__other"; renderSvcCats(); G("category_other").hidden = false; G("category_other").focus(); } }, "Other…"));
    $("#shop-list").replaceChildren(...freq(services.map((r) => r.shop)).map((x) => h("option", { value: x.k })));
  }
  function svcAutofill() { // last shop and cost for this category
    if (svcEditing || !svcCat || svcCat === "__other") return;
    const prev = services.filter((r) => r.category === svcCat).pop();
    if (prev?.shop && (!G("shop").value || G("shop").dataset.auto)) { G("shop").value = prev.shop; G("shop").dataset.auto = "1"; }
    G("cost").placeholder = prev?.cost ? f(prev.cost, 2) : "";
  }
  function openSvcForm(row, cat) {
    svcEditing = row || null;
    svcForm.reset();
    delete G("shop").dataset.auto;
    $("#svc-err").textContent = "";
    $("#svc-scan").hidden = !S.cfg.scan;
    $("#svc-scan-thumb").hidden = true;
    $("#svc-scan-title").textContent = "Scan a receipt";
    $("#svc-scan-sub").textContent = "Fills in date, shop, cost and mileage";
    const known = new Set(svcCategories());
    if (row) {
      for (const k of ["date", "odometer", "cost", "shop", "notes"]) G(k).value = row[k] ?? "";
      svcCat = known.has(row.category) ? row.category : row.category ? "__other" : null;
      G("category_other").value = svcCat === "__other" ? row.category : "";
      $("#svc-form-title").textContent = `Edit service · ${fmtDate(row.date, { month: "short", day: "numeric", year: "numeric" })}`;
      $("#svc-save span").textContent = "Save changes";
    } else {
      G("date").value = todayISO();
      svcCat = cat || null;
      $("#svc-form-title").textContent = "New service";
      $("#svc-save span").textContent = "Save service";
    }
    G("category_other").hidden = svcCat !== "__other";
    $("#svc-delete").hidden = !row;
    svcOdoHint();
    renderSvcCats(); svcAutofill();
    svcForm.hidden = false;
    svcForm.classList.remove("stagger"); void svcForm.offsetWidth;
    svcForm.style.animation = "none"; void svcForm.offsetWidth; svcForm.style.animation = "rise .45s var(--ease-out) both";
    svcForm.scrollIntoView({ behavior: reduced ? "auto" : "smooth", block: "start" });
  }
  function svcOdoHint() {
    const est = estOdometer(G("date").value || todayISO());
    G("odometer").placeholder = est ? `≈ ${f(est, 0)}` : "";
    $("#svc-odo-hint").textContent = est && !svcEditing ? "blank = estimate" : "";
  }
  G("date").addEventListener("change", svcOdoHint);
  G("shop").addEventListener("input", () => delete G("shop").dataset.auto);
  $("#svc-add").addEventListener("click", () => openSvcForm(null));
  $("#svc-cancel").addEventListener("click", () => { svcForm.hidden = true; svcEditing = null; });

  svcForm.addEventListener("submit", async (e) => {
    e.preventDefault();
    const err = $("#svc-err");
    const cat = svcCat === "__other" ? G("category_other").value.trim() : svcCat;
    if (!G("date").value) { err.textContent = "Pick a date"; return; }
    if (!cat) { err.textContent = "Pick what was done"; return; }
    for (const k of ["odometer", "cost"]) if (G(k).value.trim() && num(G(k).value) == null) { err.textContent = `${k === "cost" ? "Cost" : "Odometer"} isn't a number`; return; }
    let odo = num(G("odometer").value), estimated = false;
    if (odo == null && !svcEditing) { const est = estOdometer(G("date").value); if (est) { odo = Math.round(est); estimated = true; } }
    const body = { date: G("date").value, odometer: odo, category: cat, cost: num(G("cost").value),
      shop: G("shop").value.trim() || null, notes: G("notes").value.trim() || null };
    if (estimated) body.notes = [body.notes, "Odometer estimated"].filter(Boolean).join(". ");
    err.textContent = "";
    const btn = $("#svc-save"); btn.disabled = true; btn.classList.add("saving");
    try {
      if (svcEditing) { await S.update(svcEditing.id, body, "services"); toast("Changes saved"); }
      else { await S.create(body, "services"); toast(estimated ? `Saved · odometer estimated at ${f(odo, 0)}` : "Service saved"); }
      const wasNew = !svcEditing;
      svcForm.hidden = true; svcEditing = null;
      await load();
      if (LOCAL && wasNew && S.settings.remind) offerBackup();
    } catch (ex) { err.textContent = ex.message; }
    finally { setTimeout(() => { btn.disabled = false; btn.classList.remove("saving"); }, 350); }
  });
  $("#svc-delete").addEventListener("click", async () => {
    if (!svcEditing || !confirm(`Delete the ${svcEditing.category} on ${fmtDate(svcEditing.date)}?`)) return;
    try { await S.remove(svcEditing.id, "services"); toast("Deleted"); svcForm.hidden = true; svcEditing = null; await load(); }
    catch (ex) { $("#svc-err").textContent = ex.message; }
  });
  $("#svc-scan-file").addEventListener("change", async (e) => {
    const file = e.target.files[0];
    e.target.value = "";
    if (!file) return;
    const box = $("#svc-scan"), thumb = $("#svc-scan-thumb");
    thumb.src = URL.createObjectURL(file); thumb.hidden = false; thumb.alt = "Scanned receipt";
    box.classList.add("busy");
    $("#svc-scan-title").textContent = "Reading receipt…";
    $("#svc-scan-sub").textContent = "First scan can take ~40 s";
    try {
      const r = await S.scan(await shrink(file).catch(() => file), "service");
      const got = [];
      if (r.date) { G("date").value = r.date; got.push("date"); svcOdoHint(); }
      if (r.odometer != null) { G("odometer").value = r.odometer; got.push("mileage"); }
      if (r.cost != null) { G("cost").value = r.cost.toFixed(2); got.push("cost"); }
      if (r.shop) { G("shop").value = r.shop; delete G("shop").dataset.auto; got.push("shop"); }
      if (r.category) { svcCat = r.category; G("category_other").hidden = true; renderSvcCats(); got.push("service"); }
      if (r.notes && !G("notes").value) { G("notes").value = r.notes; got.push("notes"); }
      $("#svc-scan-title").textContent = got.length ? `Filled ${got.join(", ")}` : "Couldn't read the receipt";
      $("#svc-scan-sub").textContent = got.length ? "Check them before saving" : "Try again flatter, with less glare";
    } catch (ex) {
      $("#svc-scan-title").textContent = "Scan failed";
      $("#svc-scan-sub").textContent = ex.message;
    } finally { box.classList.remove("busy"); }
  });

  function renderService() {
    const { items, missing, est, rate } = dueList();
    $("#svc-est").textContent = est ? `Odometer now ≈ ${f(est, 0)} mi${rate ? ` · ${f(rate, 0)} mi/day lately` : ""}` : "Log a fill-up or service with an odometer reading to start reminders.";
    $("#due").replaceChildren(
      ...items.map((d, i) => h("button", {
        type: "button", class: `due ${d.status}`, style: `--i:${i}`,
        "aria-label": `${d.cat}. ${dueText(d)}. Log one`,
        onclick: () => openSvcForm(null, d.cat),
      }, h("b", {}, d.cat), h("span", { class: "when" }, dueText(d)),
      h("span", { class: "last" }, `Every ${f(d.miles, 0)} mi · last ${fmtDate(d.last.date, { month: "short", day: "numeric", year: "numeric" })} at ${f(d.lastOdo, 0)} mi` +
        (d.eta ? ` · around ${d.eta.toLocaleDateString("en-US", { month: "short", day: "numeric" })}` : "")),
      h("span", { class: "bar", "aria-hidden": "true" }, h("i", { style: `width:${Math.max(3, Math.min(100, d.used * 100))}%` })))),
      ...(missing.length ? [h("p", { class: "due-none" }, `Not logged yet: ${missing.join(", ")}. Log one to start its reminder.`)] : []));
    $("#due").classList.remove("stagger"); void $("#due").offsetWidth; $("#due").classList.add("stagger");

    const year = String(new Date().getFullYear());
    const sum = (rs) => rs.reduce((a, r) => a + (r.cost || 0), 0);
    const rs = readings(), span = rs.length > 1 ? rs[rs.length - 1].odo - rs[0].odo : 0;
    const total = sum(services);
    $("#svc-tiles").replaceChildren(
      h("div", { class: "tile" }, h("p", { class: "eyebrow" }, "This year"), h("p", { class: "v" }, `$${f(sum(services.filter((r) => r.date.startsWith(year))), 0)}`),
        h("p", { class: "s" }, plural(services.filter((r) => r.date.startsWith(year)).length, "service"))),
      h("div", { class: "tile" }, h("p", { class: "eyebrow" }, "All time"), h("p", { class: "v" }, `$${f(total, 0)}`),
        h("p", { class: "s" }, span && total ? `${f((total / span) * 100, 1)}¢ per mile` : plural(services.length, "service"))));

    const list = $("#svc-list");
    if (!services.length) { list.replaceChildren(h("p", { class: "empty" }, "No services logged yet.")); return; }
    const out = []; let yr = null, i = 0;
    for (const r of services.slice().reverse()) {
      const y = r.date.slice(0, 4);
      if (y !== yr) {
        yr = y;
        const ys = services.filter((x) => x.date.startsWith(y));
        out.push(h("div", { class: "month" }, h("h2", {}, y), h("span", {}, `${ys.length} service${ys.length > 1 ? "s" : ""} · $${f(sum(ys), 0)}`)));
      }
      const meta = [r.odometer && `${f(r.odometer, 0)} mi`, r.shop].filter(Boolean).join(" · ");
      out.push(h("button", { type: "button", class: "entry svc-entry", style: `--i:${Math.min(i++, 12)}`,
        "aria-label": `${r.category}, ${fmtDate(r.date, { month: "long", day: "numeric", year: "numeric" })}. Edit`,
        onclick: () => openSvcForm(r) },
        h("span", { class: "e-date" }, h("b", {}, String(toDate(r.date).getDate())), toDate(r.date).toLocaleDateString("en-US", { month: "short" })),
        h("span", { class: "e-mid" }, h("span", { class: "e-meta", style: "display:block;color:var(--text);font-family:var(--sans);font-weight:650" }, r.category || "Service"),
          meta ? h("span", { class: "e-note", style: "display:block" }, meta) : null,
          r.notes ? h("span", { class: "e-note", style: "display:block" }, r.notes) : null),
        h("span", { class: "e-mpg" }, r.cost != null ? `$${f(r.cost, 0)}` : "–")));
    }
    list.replaceChildren(...out);

    const iv = intervals();
    const cats = [...new Set([...Object.keys(SERVICE_DEFAULTS), ...services.map((r) => r.category).filter(Boolean)])];
    if (!$("#intervals").contains(document.activeElement)) $("#intervals").replaceChildren(...cats.map((c) => {
      const inp = h("input", { inputmode: "numeric", placeholder: "off", value: iv[c] ?? "", "aria-label": `${c} reminder interval in miles` });
      inp.addEventListener("change", async () => {
        const v = num(inp.value);
        const next = { ...(appSettings.intervals || {}), [c]: v && v > 0 ? Math.round(v) : null };
        try { appSettings = { ...appSettings, ...(await S.saveSettings({ intervals: next })) }; appSettings.intervals = next; toast(v ? `${c}: every ${f(v, 0)} mi` : `${c}: no reminder`); renderService(); renderDueBanner(); }
        catch (ex) { toast(ex.message); }
      });
      return h("label", { class: "int-row" }, h("span", {}, c), inp);
    }));
  }

  // ------------------------------------------------------------ nav
  const views = ["log", "history", "service", "stats"];
  function show(name) {
    views.forEach((v) => { $(`#view-${v}`).hidden = v !== name; });
    $$(".tab").forEach((t) => (t.dataset.view === name ? t.setAttribute("aria-current", "page") : t.removeAttribute("aria-current")));
    $(".tab-pill").style.transform = `translateX(${views.indexOf(name) * 100}%)`;
    const v = $(`#view-${name}`);
    v.style.animation = "none"; void v.offsetWidth; v.style.animation = "";
    if (name === "stats") renderStats();
    if (name === "history") renderHistory();
    if (name === "service") renderService();
    history.replaceState(null, "", name === "log" ? location.pathname : `#${name}`);
  }
  $$(".tab").forEach((t) => t.addEventListener("click", () => { if (t.dataset.view === "log" && editing) fillForm(null); show(t.dataset.view); window.scrollTo({ top: 0 }); }));

  $("#import").addEventListener("change", async (e) => {
    const file = e.target.files[0];
    if (!file) return;
    try { const r = await S.importFile(file); toast(`Imported ${r.added}, skipped ${r.skipped} duplicates`); await load(); }
    catch (err) { toast(err.message); }
    e.target.value = "";
  });

  // parallax: hero drifts up slower than content and fades as you scroll
  const hero = $(".hero");
  let sy = 0, ticking = false;
  if (!reduced) addEventListener("scroll", () => {
    sy = scrollY;
    if (ticking) return; ticking = true;
    requestAnimationFrame(() => {
      ticking = false;
      const p = Math.min(1, sy / 220);
      hero.style.transform = `translateY(${sy * 0.35}px) scale(${1 - p * 0.08})`;
      hero.style.opacity = String(1 - p * 0.85);
    });
  }, { passive: true });

  let rz = 0;
  addEventListener("resize", () => { cancelAnimationFrame(rz); rz = requestAnimationFrame(() => { if (!$("#view-stats").hidden) renderChart(inRange().filter((r) => r.dash_mpg)); }); });

  function renderAll() {
    renderHero();
    renderDueBanner();
    if (!$("#view-service").hidden) renderService();
    renderSettings();
    prepareExport();
    if (!editing && !dirty) fillForm(null);
    else { renderChoices(); preview(); }
    if (!$("#view-history").hidden) renderHistory();
    if (!$("#view-stats").hidden) renderStats();
  }

  addEventListener("online", flush);
  document.addEventListener("visibilitychange", () => { if (!document.hidden) { flush(); load(); } });
  if ("serviceWorker" in navigator && location.protocol === "https:") navigator.serviceWorker.register("sw.js").catch(() => {});

  renderSync();
  const start = location.hash.slice(1);
  show(views.includes(start) ? start : "log");
  load().then(flush);
})();
