/* Storage for MPG Tracker.
   Server mode: the Python app (api/...). Local mode: IndexedDB on this device,
   with Excel export/import done in the browser. Both expose the same methods. */
(function () {
  "use strict";

  // config.js is served by the Python app in server mode and is a static
  // {mode:"local"} file everywhere else. Remember the mode in case it's ever
  // missing offline, so a server install never silently falls back to local.
  const saved = localStorage.getItem("mpg-mode");
  const cfg = Object.assign({ mode: saved || "local", vehicle: null, scan: false }, window.MPG_CONFIG || {});
  localStorage.setItem("mpg-mode", cfg.mode);
  const LOCAL = cfg.mode !== "server";

  const FIELDS = ["date", "odometer", "trip_miles", "dash_mpg", "gallons", "drive_minutes", "octane",
    "drive_mode", "price_per_gal", "total_cost", "station", "tags", "notes"];
  const sortRows = (rows) => rows.sort((a, b) =>
    a.date < b.date ? -1 : a.date > b.date ? 1 : (a.odometer || 0) - (b.odometer || 0) || a.id - b.id);
  const nowISO = () => new Date().toISOString().slice(0, 19);
  const settings = {
    get vehicle() { return localStorage.getItem("mpg-vehicle") || "My Truck"; },
    set vehicle(v) { localStorage.setItem("mpg-vehicle", v); },
    get remind() { return localStorage.getItem("mpg-remind") !== "0"; },
    set remind(v) { localStorage.setItem("mpg-remind", v ? "1" : "0"); },
    get lastBackup() { return localStorage.getItem("mpg-last-backup"); },
    set lastBackup(v) { localStorage.setItem("mpg-last-backup", v); },
  };

  // ---------------------------------------------------------------- server
  async function api(method, path, body) {
    const res = await fetch(path, {
      method,
      headers: { "X-MPG": "1", ...(body && !(body instanceof Blob) ? { "Content-Type": "application/json" } : {}) },
      body: body instanceof Blob ? body : body ? JSON.stringify(body) : undefined,
    });
    if (path.endsWith(".xlsx") && res.ok) return res.blob();
    const j = await res.json().catch(() => ({}));
    if (!res.ok) { const e = new Error(j.error || `HTTP ${res.status}`); e.http = res.status; throw e; }
    return j;
  }
  const server = {
    list: () => api("GET", "api/fillups"),
    create: (b) => api("POST", "api/fillups", b),
    update: (id, b) => api("PUT", `api/fillups/${id}`, b),
    remove: (id) => api("DELETE", `api/fillups/${id}`),
    importFile: (file) => api("POST", "api/import", file),
    exportBlob: () => api("GET", "api/export.xlsx"),
    scan: (blob) => api("POST", "api/scan", blob),
  };

  // ---------------------------------------------------------------- IndexedDB
  let dbP = null;
  function db() {
    if (!dbP) dbP = new Promise((res, rej) => {
      const r = indexedDB.open("mpg-tracker", 1);
      r.onupgradeneeded = () => r.result.createObjectStore("fillups", { keyPath: "id", autoIncrement: true });
      r.onsuccess = () => res(r.result);
      r.onerror = () => rej(r.error);
    });
    return dbP;
  }
  async function tx(mode, fn) {
    const d = await db();
    return new Promise((res, rej) => {
      const t = d.transaction("fillups", mode), st = t.objectStore("fillups");
      let out;
      Promise.resolve(fn(st, (v) => (out = v))).catch(rej);
      t.oncomplete = () => res(out);
      t.onerror = t.onabort = () => rej(t.error || new Error("Storage error"));
    });
  }
  const req = (r) => new Promise((res, rej) => { r.onsuccess = () => res(r.result); r.onerror = () => rej(r.error); });
  function clean(b) {
    const o = {};
    for (const k of FIELDS) if (k in b) o[k] = b[k] === "" || b[k] === undefined ? null : b[k];
    if ("date" in o && !/^\d{4}-\d{2}-\d{2}$/.test(o.date || "")) throw new Error("date must be YYYY-MM-DD");
    return o;
  }
  const allRows = () => tx("readonly", (st, done) => req(st.getAll()).then(done));
  const local = {
    async list() { return { vehicle: settings.vehicle, fillups: sortRows(await allRows()) }; },
    async create(b) {
      const row = { ...clean(b), created_at: nowISO(), updated_at: nowISO() };
      if (!row.date) throw new Error("date is required");
      row.id = await tx("readwrite", (st, done) => req(st.add(row)).then(done));
      return row;
    },
    async update(id, b) {
      return tx("readwrite", async (st, done) => {
        const cur = await req(st.get(id));
        if (!cur) throw new Error("not found");
        const row = { ...cur, ...clean(b), updated_at: nowISO() };
        await req(st.put(row));
        done(row);
      });
    },
    remove: (id) => tx("readwrite", (st, done) => req(st.delete(id)).then(() => done({ deleted: 1 }))),
    clear: () => tx("readwrite", (st, done) => req(st.clear()).then(() => done(true))),
    async importFile(file) {
      const entries = await readXlsx(await file.arrayBuffer());
      const have = new Set((await allRows()).map((r) => `${r.date}|${r.odometer}|${r.trip_miles}`));
      let added = 0;
      for (const e of entries) {
        const key = `${e.date}|${e.odometer}|${e.trip_miles}`;
        if (have.has(key)) continue;
        have.add(key);
        await local.create(e);
        added++;
      }
      return { added, skipped: entries.length - added };
    },
    async exportBlob() { return writeXlsx(sortRows(await allRows()), settings.vehicle); },
  };

  // ---------------------------------------------------------------- zip
  const CRC = (() => { const t = new Uint32Array(256); for (let n = 0; n < 256; n++) { let c = n; for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1; t[n] = c >>> 0; } return t; })();
  const crc32 = (b) => { let c = 0xffffffff; for (let i = 0; i < b.length; i++) c = CRC[(c ^ b[i]) & 0xff] ^ (c >>> 8); return (c ^ 0xffffffff) >>> 0; };
  function zip(files) { // stored (uncompressed) zip; plenty for a few hundred rows
    const enc = new TextEncoder(), parts = [], central = [];
    let off = 0;
    for (const [name, text] of files) {
      const nb = enc.encode(name), data = enc.encode(text), crc = crc32(data);
      const h = new DataView(new ArrayBuffer(30));
      h.setUint32(0, 0x04034b50, true); h.setUint16(4, 20, true); h.setUint16(8, 0, true);
      h.setUint32(14, crc, true); h.setUint32(18, data.length, true); h.setUint32(22, data.length, true);
      h.setUint16(26, nb.length, true);
      parts.push(h, nb, data);
      const c = new DataView(new ArrayBuffer(46));
      c.setUint32(0, 0x02014b50, true); c.setUint16(4, 20, true); c.setUint16(6, 20, true);
      c.setUint32(16, crc, true); c.setUint32(20, data.length, true); c.setUint32(24, data.length, true);
      c.setUint16(28, nb.length, true); c.setUint32(42, off, true);
      central.push(c, nb);
      off += 30 + nb.length + data.length;
    }
    const size = central.reduce((s, p) => s + p.byteLength, 0);
    const e = new DataView(new ArrayBuffer(22));
    e.setUint32(0, 0x06054b50, true); e.setUint16(8, files.length, true); e.setUint16(10, files.length, true);
    e.setUint32(12, size, true); e.setUint32(16, off, true);
    return new Blob([...parts, ...central, e], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
  }
  async function unzip(buf) {
    const dv = new DataView(buf), u8 = new Uint8Array(buf), dec = new TextDecoder();
    let e = buf.byteLength - 22;
    while (e >= 0 && dv.getUint32(e, true) !== 0x06054b50) e--;
    if (e < 0) throw new Error("Not an .xlsx file");
    const count = dv.getUint16(e + 10, true);
    let p = dv.getUint32(e + 16, true);
    const out = {};
    for (let i = 0; i < count; i++) {
      const method = dv.getUint16(p + 10, true), csize = dv.getUint32(p + 20, true);
      const nlen = dv.getUint16(p + 28, true), xlen = dv.getUint16(p + 30, true), clen = dv.getUint16(p + 32, true);
      const lho = dv.getUint32(p + 42, true), name = dec.decode(u8.subarray(p + 46, p + 46 + nlen));
      p += 46 + nlen + xlen + clen;
      if (!/\.(xml|rels)$/.test(name)) continue;
      const start = lho + 30 + dv.getUint16(lho + 26, true) + dv.getUint16(lho + 28, true);
      const raw = u8.subarray(start, start + csize);
      let data = raw;
      if (method === 8) {
        if (typeof DecompressionStream === "undefined") throw new Error("This browser can't read .xlsx files");
        data = new Uint8Array(await new Response(new Blob([raw]).stream().pipeThrough(new DecompressionStream("deflate-raw"))).arrayBuffer());
      } else if (method !== 0) continue;
      out[name] = dec.decode(data);
    }
    return out;
  }

  // ---------------------------------------------------------------- xlsx
  const COLS = [["Date", 12], ["MPG", 7], ["Calc MPG", 9], ["Gallons", 9], ["Time", 8], ["Miles", 8], ["Fuel", 6],
    ["Drive Mode", 12], ["Cum Miles", 11], ["Price/gal", 10], ["Total $", 9], ["Station", 16], ["Tags", 18], ["Notes", 48]];
  const esc = (s) => String(s).replace(/[<>&"]/g, (c) => ({ "<": "&lt;", ">": "&gt;", "&": "&amp;", '"': "&quot;" }[c]));
  const colName = (i) => { let s = ""; i++; while (i) { const m = (i - 1) % 26; s = String.fromCharCode(65 + m) + s; i = (i - m - 1) / 26; } return s; };
  const serial = (iso) => (Date.UTC(+iso.slice(0, 4), +iso.slice(5, 7) - 1, +iso.slice(8, 10)) - Date.UTC(1899, 11, 30)) / 864e5;

  function writeXlsx(rows, vehicle) {
    const cell = (r, c, v, style) => {
      const ref = colName(c) + r, s = style ? ` s="${style}"` : "";
      if (v == null || v === "") return "";
      if (typeof v === "number") return `<c r="${ref}"${s}><v>${v}</v></c>`;
      return `<c r="${ref}"${s} t="inlineStr"><is><t xml:space="preserve">${esc(v)}</t></is></c>`;
    };
    const lines = [`<row r="1">${COLS.map(([h], i) => cell(1, i, h, 5)).join("")}</row>`];
    rows.forEach((r, i) => {
      const n = i + 2, calc = r.trip_miles && r.gallons ? Math.round((r.trip_miles / r.gallons) * 100) / 100 : null;
      const vals = [[serial(r.date), 1], [r.dash_mpg], [calc], [r.gallons], [r.drive_minutes != null ? r.drive_minutes / 1440 : null, 2],
        [r.trip_miles], [r.octane], [r.drive_mode], [r.odometer], [r.price_per_gal, 3], [r.total_cost, 4], [r.station], [r.tags], [r.notes]];
      lines.push(`<row r="${n}">${vals.map(([v, s], c) => cell(n, c, v, s)).join("")}</row>`);
    });
    const sheet = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews><cols>${COLS.map(([, w], i) => `<col min="${i + 1}" max="${i + 1}" width="${w}" customWidth="1"/>`).join("")}</cols><sheetData>${lines.join("")}</sheetData></worksheet>`;
    const styles = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><numFmts count="4"><numFmt numFmtId="164" formatCode="yyyy-mm-dd"/><numFmt numFmtId="165" formatCode="[h]:mm"/><numFmt numFmtId="166" formatCode="&quot;$&quot;0.000"/><numFmt numFmtId="167" formatCode="&quot;$&quot;0.00"/></numFmts><fonts count="2"><font><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font></fonts><fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FF2B2A27"/></patternFill></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="6"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/><xf numFmtId="165" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/><xf numFmtId="166" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/><xf numFmtId="167" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/><xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1"/></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>`;
    const title = esc(String(vehicle || "MPG").slice(0, 31).replace(/[\\/?*[\]:]/g, " "));
    return zip([
      ["[Content_Types].xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>`],
      ["_rels/.rels", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`],
      ["xl/workbook.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="${title}" sheetId="1" r:id="rId1"/></sheets></workbook>`],
      ["xl/_rels/workbook.xml.rels", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>`],
      ["xl/styles.xml", styles],
      ["xl/worksheets/sheet1.xml", sheet],
    ]);
  }

  // Reads the first sheet of an .xlsx with the original headers (Date, MPG,
  // Gallons, Time, Miles, Fuel, Drive Mode, Cum Miles, Notes) or this app's export.
  async function readXlsx(buf) {
    const files = await unzip(buf);
    const xml = (s) => new DOMParser().parseFromString(s, "application/xml");
    const byTag = (doc, tag) => [...doc.getElementsByTagNameNS("*", tag)];
    const shared = files["xl/sharedStrings.xml"]
      ? byTag(xml(files["xl/sharedStrings.xml"]), "si").map((si) => byTag(si, "t").map((t) => t.textContent).join(""))
      : [];
    let sheetPath = "xl/worksheets/sheet1.xml";
    if (files["xl/workbook.xml"] && files["xl/_rels/workbook.xml.rels"]) {
      const first = byTag(xml(files["xl/workbook.xml"]), "sheet")[0];
      const rid = first && (first.getAttribute("r:id") || first.getAttributeNS("http://schemas.openxmlformats.org/officeDocument/2006/relationships", "id"));
      const rel = byTag(xml(files["xl/_rels/workbook.xml.rels"]), "Relationship").find((r) => r.getAttribute("Id") === rid);
      if (rel) { const t = rel.getAttribute("Target").replace(/^\//, ""); sheetPath = t.startsWith("xl/") ? t : "xl/" + t; }
    }
    if (!files[sheetPath]) throw new Error("Couldn't find a sheet in that file");
    const rows = byTag(xml(files[sheetPath]), "row").map((row) => {
      const out = {};
      for (const c of byTag(row, "c")) {
        const col = (c.getAttribute("r") || "").replace(/\d+/g, ""), t = c.getAttribute("t");
        const v = byTag(c, "v")[0]?.textContent;
        if (t === "s") out[col] = shared[+v];
        else if (t === "inlineStr") out[col] = byTag(c, "t").map((x) => x.textContent).join("");
        else if (t === "str" || t === "b" || t === "e") out[col] = v;
        else if (v != null) out[col] = Number(v);
      }
      return out;
    });
    if (!rows.length) return [];
    const head = {};
    for (const [col, h] of Object.entries(rows[0])) head[String(h).trim().toLowerCase()] = col;
    const get = (r, ...names) => { for (const n of names) if (head[n] && r[head[n]] != null && r[head[n]] !== "") return r[head[n]]; return null; };
    const num = (v) => { if (v == null || v === "") return null; if (typeof v === "number") return v; const x = parseFloat(String(v).replace(/[,$]/g, "")); return Number.isFinite(x) && /^\s*\$?\d[\d,]*(\.\d+)?\s*$/.test(String(v)) ? x : null; };
    const out = [];
    for (const r of rows.slice(1)) {
      let d = get(r, "date");
      if (d == null) continue;
      if (typeof d === "number") d = new Date(Date.UTC(1899, 11, 30) + Math.round(d) * 864e5).toISOString().slice(0, 10);
      else { d = String(d).trim().slice(0, 10); if (!/^\d{4}-\d{2}-\d{2}$/.test(d)) { const p = new Date(d); if (isNaN(p)) continue; d = p.toISOString().slice(0, 10); } }
      let notes = get(r, "notes"); notes = notes != null ? String(notes).trim() : null;
      const galRaw = get(r, "gallons"), gal = num(galRaw);
      if (gal == null && galRaw != null) { const extra = `Gallons recorded as '${galRaw}'`; notes = notes ? `${notes}. ${extra}` : extra; }
      const t = get(r, "time");
      let mins = null;
      if (typeof t === "number") mins = Math.round(t * 1440);
      else if (t != null) { const m = String(t).match(/^(\d+):(\d{2})/); if (m) mins = +m[1] * 60 + +m[2]; }
      const oct = num(get(r, "fuel"));
      const s = (v) => (v == null ? null : String(v).trim() || null);
      out.push({
        date: d, dash_mpg: num(get(r, "mpg")), gallons: gal, drive_minutes: mins, trip_miles: num(get(r, "miles")),
        octane: oct ? Math.round(oct) : null, drive_mode: s(get(r, "drive mode")), odometer: num(get(r, "cum miles", "odometer")),
        price_per_gal: num(get(r, "price/gal")), total_cost: num(get(r, "total $")), station: s(get(r, "station")),
        tags: s(get(r, "tags")), notes,
      });
    }
    return out;
  }

  window.MPGStore = { cfg, LOCAL, settings, xlsx: { write: writeXlsx, read: readXlsx }, ...(LOCAL ? local : server) };
})();
