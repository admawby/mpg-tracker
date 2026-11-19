"""MPG Tracker: fuel fill-up log.

Stdlib HTTP server + SQLite; openpyxl only for Excel import/export.
Run:  python app.py               (serve on $PORT, default 8840)
      python app.py import FILE   (load an .xlsx in the original sheet layout)
"""
import datetime as dt
import io
import ipaddress
import json
import os
import re
import shutil
import sqlite3
import sys
import threading
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from urllib.parse import urlparse

ROOT = os.path.dirname(os.path.abspath(__file__))
STATIC = os.path.join(ROOT, "static")
DB_PATH = os.environ.get("MPG_DB", os.path.join(ROOT, "data", "mpg.sqlite"))
PORT = int(os.environ.get("PORT", "8840"))
VEHICLE = os.environ.get("MPG_VEHICLE", "My Truck")
BACKUP_DIR = os.environ.get("MPG_BACKUP_DIR")  # copied to after each new fill-up
BACKUP_KEEP = int(os.environ.get("MPG_BACKUP_KEEP", "60"))
OLLAMA_URL = os.environ.get("OLLAMA_URL", "")  # empty = photo scan off
VISION_MODEL = os.environ.get("VISION_MODEL", "qwen3-vl:8b")

# Private, CGNAT/Tailscale and Docker networks only: keep this off the open internet.
ALLOWED_NETS = [ipaddress.ip_network(n) for n in (
    "127.0.0.0/8", "10.0.0.0/8", "172.16.0.0/12", "192.168.0.0/16",
    "100.64.0.0/10", "::1/128", "fd7a:115c:a1e0::/48", "fc00::/7")]

FIELDS = {  # column -> type
    "date": str, "odometer": float, "trip_miles": float, "dash_mpg": float,
    "gallons": float, "drive_minutes": int, "octane": int, "drive_mode": str,
    "price_per_gal": float, "total_cost": float, "station": str,
    "tags": str, "notes": str,
}

SCHEMA = """
CREATE TABLE IF NOT EXISTS fillups (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  date TEXT NOT NULL,
  odometer REAL, trip_miles REAL, dash_mpg REAL, gallons REAL,
  drive_minutes INTEGER, octane INTEGER, drive_mode TEXT,
  price_per_gal REAL, total_cost REAL, station TEXT,
  tags TEXT, notes TEXT,
  created_at TEXT NOT NULL, updated_at TEXT NOT NULL
);
CREATE INDEX IF NOT EXISTS fillups_date ON fillups(date);
"""

_lock = threading.Lock()
_backup_lock = threading.Lock()


def db():
    os.makedirs(os.path.dirname(DB_PATH), exist_ok=True)
    con = sqlite3.connect(DB_PATH)
    con.row_factory = sqlite3.Row
    con.executescript(SCHEMA)
    return con


def now():
    return dt.datetime.now().isoformat(timespec="seconds")


def clean(payload):
    """Coerce an incoming JSON object to known columns; '' -> NULL."""
    out = {}
    for k, typ in FIELDS.items():
        if k not in payload:
            continue
        v = payload[k]
        if v is None or (isinstance(v, str) and v.strip() == ""):
            out[k] = None
            continue
        if typ is str:
            out[k] = str(v).strip()[:2000]
        else:
            try:
                out[k] = typ(float(v)) if typ is int else float(v)
            except (TypeError, ValueError):
                raise ValueError(f"{k} must be a number")
    if "date" in out:
        if not out["date"] or not re.fullmatch(r"\d{4}-\d{2}-\d{2}", out["date"]):
            raise ValueError("date must be YYYY-MM-DD")
    return out


def rows(con):
    return [dict(r) for r in con.execute(
        "SELECT * FROM fillups ORDER BY date, COALESCE(odometer, 0), id")]


# ---------------------------------------------------------------- Excel
EXPORT_COLS = [
    ("Date", 12), ("MPG", 7), ("Calc MPG", 9), ("Gallons", 9), ("Time", 8),
    ("Miles", 8), ("Fuel", 6), ("Drive Mode", 12), ("Cum Miles", 11),
    ("Price/gal", 10), ("Total $", 9), ("Station", 16), ("Tags", 18), ("Notes", 48),
]


def export_xlsx(data):
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill
    wb = Workbook()
    ws = wb.active
    ws.title = VEHICLE
    ws.append([c for c, _ in EXPORT_COLS])
    for i, (_, w) in enumerate(EXPORT_COLS, start=1):
        ws.column_dimensions[ws.cell(1, i).column_letter].width = w
        ws.cell(1, i).font = Font(bold=True, color="FFFFFF")
        ws.cell(1, i).fill = PatternFill("solid", fgColor="2B2A27")
    ws.freeze_panes = "A2"
    for r in data:
        calc = (round(r["trip_miles"] / r["gallons"], 2)
                if r["trip_miles"] and r["gallons"] else None)
        t = r["drive_minutes"] / 1440 if r["drive_minutes"] is not None else None
        ws.append([
            dt.date.fromisoformat(r["date"]), r["dash_mpg"], calc, r["gallons"], t,
            r["trip_miles"], r["octane"], r["drive_mode"], r["odometer"],
            r["price_per_gal"], r["total_cost"], r["station"], r["tags"], r["notes"],
        ])
        n = ws.max_row
        ws.cell(n, 1).number_format = "yyyy-mm-dd"
        ws.cell(n, 5).number_format = "[h]:mm"
        ws.cell(n, 10).number_format = '"$"0.000'
        ws.cell(n, 11).number_format = '"$"0.00'
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def _num(v):
    if v is None or v == "":
        return None
    if isinstance(v, (int, float)):
        return float(v)
    try:
        return float(str(v).replace(",", "").replace("$", ""))
    except ValueError:
        return None


def import_xlsx(blob):
    """Parse a sheet with the original headers (Date, MPG, Gallons, Time, Miles,
    Fuel, Drive Mode, Cum Miles, Notes) or this app's export. Returns entries."""
    from openpyxl import load_workbook
    ws = load_workbook(io.BytesIO(blob), data_only=True).worksheets[0]
    it = ws.iter_rows(values_only=True)
    head = [str(h).strip().lower() if h is not None else "" for h in next(it)]
    col = {h: i for i, h in enumerate(head)}

    def get(row, *names):
        for n in names:
            if n in col and col[n] < len(row):
                return row[col[n]]
        return None

    out = []
    for row in it:
        d = get(row, "date")
        if d is None:
            continue
        if isinstance(d, dt.datetime):
            d = d.date()
        elif isinstance(d, (int, float)):
            d = dt.date(1899, 12, 30) + dt.timedelta(days=int(d))
        elif isinstance(d, str):
            try:
                d = dt.date.fromisoformat(d.strip()[:10])
            except ValueError:
                continue
        notes = get(row, "notes")
        notes = str(notes).strip() if notes else None
        gal_raw = get(row, "gallons")
        gal = _num(gal_raw)
        if gal is None and gal_raw not in (None, ""):
            # e.g. "28..." in the original sheet: keep what was written
            extra = f"Gallons recorded as '{gal_raw}'"
            notes = f"{notes}. {extra}" if notes else extra
        t = get(row, "time")
        if isinstance(t, dt.time):
            mins = t.hour * 60 + t.minute
        elif isinstance(t, dt.timedelta):
            mins = round(t.total_seconds() / 60)
        elif isinstance(t, (int, float)):
            mins = round(float(t) * 1440)
        else:
            mins = None
        oct_ = _num(get(row, "fuel"))
        out.append({
            "date": d.isoformat(),
            "dash_mpg": _num(get(row, "mpg")),
            "gallons": gal,
            "drive_minutes": mins,
            "trip_miles": _num(get(row, "miles")),
            "octane": int(oct_) if oct_ else None,
            "drive_mode": (str(get(row, "drive mode")).strip()
                           if get(row, "drive mode") else None),
            "odometer": _num(get(row, "cum miles", "odometer")),
            "price_per_gal": _num(get(row, "price/gal")),
            "total_cost": _num(get(row, "total $")),
            "station": get(row, "station") or None,
            "tags": get(row, "tags") or None,
            "notes": notes,
        })
    return out


def insert(con, entry):
    e = clean(entry)
    if not e.get("date"):
        raise ValueError("date is required")
    e["created_at"] = e["updated_at"] = now()
    keys = list(e)
    cur = con.execute(
        f"INSERT INTO fillups ({','.join(keys)}) VALUES ({','.join('?' * len(keys))})",
        [e[k] for k in keys])
    return cur.lastrowid


# ---------------------------------------------------------------- NAS backup
def backup(reason):
    """Copy the DB (+ a readable .xlsx) to BACKUP_DIR. Runs in a thread so a
    slow or offline network share never holds up saving a fill-up."""
    if not BACKUP_DIR:
        return
    with _backup_lock:
        try:
            hist = os.path.join(BACKUP_DIR, "history")
            os.makedirs(hist, exist_ok=True)
            stamp = dt.datetime.now().strftime("%Y-%m-%d_%H%M%S")
            snap = os.path.join(hist, f"mpg_{stamp}.sqlite")
            # Snapshot locally first: SQLite locking can hang on SMB/CIFS shares.
            local = os.path.join(os.path.dirname(DB_PATH), ".backup.sqlite")
            src = sqlite3.connect(DB_PATH)
            dst = sqlite3.connect(local)
            with dst:
                src.backup(dst)
            data = [dict(r) for r in _rows_raw(src)]
            src.close()
            dst.close()
            shutil.copyfile(local, snap + ".tmp")
            os.replace(snap + ".tmp", snap)
            os.remove(local)
            xlsx = os.path.join(BACKUP_DIR, f"{VEHICLE}_MPG.xlsx")
            with open(xlsx + ".tmp", "wb") as f:
                f.write(export_xlsx(data))
            os.replace(xlsx + ".tmp", xlsx)
            snaps = sorted(n for n in os.listdir(hist) if n.endswith(".sqlite"))
            for old in snaps[:-BACKUP_KEEP]:
                os.remove(os.path.join(hist, old))
            print(f"backup ok ({reason}): {snap}", flush=True)
        except Exception as exc:
            print(f"backup FAILED ({reason}): {exc}", file=sys.stderr, flush=True)


def _rows_raw(con):
    con.row_factory = sqlite3.Row
    return con.execute("SELECT * FROM fillups ORDER BY date, COALESCE(odometer, 0), id").fetchall()


def backup_async(reason):
    threading.Thread(target=backup, args=(reason,), daemon=True).start()


# ---------------------------------------------------------------- photo scan
SCAN_PROMPT = """This photo is either a vehicle instrument cluster (trip computer), a gas pump display, or a fuel receipt.
Read only values that are clearly visible, exactly as displayed:
- kind: "dash", "pump", or "receipt"
- trip_mpg: the MPG (average) number in the Trip panel
- trip_time: the Time value in the Trip panel, exactly as shown (h:mm:ss or h:mm)
- trip_miles: the "mi" value in the Trip panel
- odometer: the odometer reading (number before "mi" at the bottom, not the trip miles)
- miles_to_empty: the number labelled "mi to E"
- gallons: gallons pumped (pump or receipt)
- price_per_gal: price per gallon (pump or receipt)
- total_cost: total sale amount in dollars (pump or receipt)
- octane: fuel grade octane if shown (87, 89, 93...)
Use null for anything not present or not readable. Reply with JSON only."""


def _f(v):
    if v is None:
        return None
    m = re.search(r"-?\d[\d,]*\.?\d*", str(v))
    return float(m.group(0).replace(",", "")) if m else None


def scan_photo(jpeg):
    import base64
    import urllib.request
    body = {"model": VISION_MODEL, "stream": False, "format": "json", "think": False,
            "keep_alive": "3m",  # free VRAM soon if the GPU is shared
            "options": {"temperature": 0, "num_ctx": 8192},
            "messages": [{"role": "user", "content": SCAN_PROMPT,
                          "images": [base64.b64encode(jpeg).decode()]}]}
    req = urllib.request.Request(f"{OLLAMA_URL}/api/chat", json.dumps(body).encode(),
                                 {"Content-Type": "application/json"})
    with urllib.request.urlopen(req, timeout=240) as r:
        raw = json.loads(json.load(r)["message"]["content"])
    out = {"kind": raw.get("kind")}
    rng = {"dash_mpg": ("trip_mpg", 3, 60), "trip_miles": ("trip_miles", 0, 1500),
           "odometer": ("odometer", 1, 2_000_000), "miles_to_empty": ("miles_to_empty", 0, 1500),
           "gallons": ("gallons", 0.5, 60), "price_per_gal": ("price_per_gal", 1, 15),
           "total_cost": ("total_cost", 1, 600)}
    for key, (src, lo, hi) in rng.items():
        v = _f(raw.get(src))
        if v is not None and lo <= v <= hi:
            out[key] = v
    t = raw.get("trip_time")
    if t:
        parts = [int(p) for p in re.findall(r"\d+", str(t))]
        if len(parts) == 3:
            out["drive_minutes"] = round(parts[0] * 60 + parts[1] + parts[2] / 60)
        elif len(parts) == 2:
            out["drive_minutes"] = parts[0] * 60 + parts[1]
    o = _f(raw.get("octane"))
    if o in (85, 86, 87, 88, 89, 90, 91, 92, 93, 94):
        out["octane"] = int(o)
    out["raw"] = raw
    return out


# ---------------------------------------------------------------- HTTP
MIME = {".html": "text/html; charset=utf-8", ".js": "text/javascript; charset=utf-8",
        ".css": "text/css; charset=utf-8", ".json": "application/json",
        ".webmanifest": "application/manifest+json", ".png": "image/png",
        ".svg": "image/svg+xml", ".woff2": "font/woff2", ".ico": "image/x-icon"}


class Handler(BaseHTTPRequestHandler):
    server_version = "mpg-tracker"

    def log_message(self, fmt, *args):
        sys.stderr.write("%s %s\n" % (self.client_address[0], fmt % args))

    def _allowed(self):
        try:
            ip = ipaddress.ip_address(self.client_address[0])
        except ValueError:
            return False
        if getattr(ip, "ipv4_mapped", None):
            ip = ip.ipv4_mapped
        return any(ip in n for n in ALLOWED_NETS)

    def _send(self, code, body=b"", ctype="application/json", headers=None):
        if isinstance(body, (dict, list)):
            body = json.dumps(body).encode()
        elif isinstance(body, str):
            body = body.encode()
        self.send_response(code)
        self.send_header("Content-Type", ctype)
        self.send_header("Content-Length", str(len(body)))
        self.send_header("X-Content-Type-Options", "nosniff")
        for k, v in (headers or {}).items():
            self.send_header(k, v)
        self.end_headers()
        if self.command != "HEAD":
            self.wfile.write(body)

    def _body(self, limit=20 * 1024 * 1024):
        n = int(self.headers.get("Content-Length") or 0)
        if n > limit:
            raise ValueError("too large")
        return self.rfile.read(n)

    def _guard(self):
        if not self._allowed():
            self._send(403, {"error": "forbidden"})
            return False
        return True

    def _csrf(self):
        if self.headers.get("X-MPG") != "1":
            self._send(403, {"error": "missing X-MPG header"})
            return False
        return True

    # -- routes
    def do_HEAD(self):
        self.do_GET()

    def do_GET(self):
        if not self._guard():
            return
        path = urlparse(self.path).path
        if path == "/api/fillups":
            with db() as con:
                return self._send(200, {"vehicle": VEHICLE, "fillups": rows(con)})
        if path == "/api/export.xlsx":
            with db() as con:
                blob = export_xlsx(rows(con))
            name = f"{VEHICLE}_MPG_{dt.date.today().isoformat()}.xlsx"
            return self._send(200, blob,
                              "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                              {"Content-Disposition": f'attachment; filename="{name}"'})
        if path == "/config.js":
            cfg = {"mode": "server", "vehicle": VEHICLE, "scan": bool(OLLAMA_URL)}
            return self._send(200, f"window.MPG_CONFIG = {json.dumps(cfg)};\n",
                              "text/javascript; charset=utf-8", {"Cache-Control": "no-cache"})
        if path == "/api/health":
            return self._send(200, {"ok": True})
        return self._static(path)

    def _static(self, path):
        if path in ("/", ""):
            path = "/index.html"
        full = os.path.realpath(os.path.join(STATIC, path.lstrip("/")))
        if not full.startswith(STATIC + os.sep) or not os.path.isfile(full):
            return self._send(404, {"error": "not found"})
        ext = os.path.splitext(full)[1]
        cache = ("public, max-age=31536000, immutable" if ext == ".woff2"
                 else "no-cache")
        with open(full, "rb") as f:
            self._send(200, f.read(), MIME.get(ext, "application/octet-stream"),
                       {"Cache-Control": cache})

    def do_POST(self):
        if not (self._guard() and self._csrf()):
            return
        path = urlparse(self.path).path
        try:
            if path == "/api/fillups":
                entry = json.loads(self._body() or b"{}")
                with _lock, db() as con:
                    new_id = insert(con, entry)
                    row = dict(con.execute("SELECT * FROM fillups WHERE id=?",
                                           (new_id,)).fetchone())
                backup_async(f"new fill-up {row['date']}")
                return self._send(201, row)
            if path == "/api/scan":
                if not OLLAMA_URL:
                    return self._send(404, {"error": "photo scan not configured"})
                blob = self._body(limit=15 * 1024 * 1024)
                if not blob:
                    return self._send(400, {"error": "no image"})
                try:
                    return self._send(200, scan_photo(blob))
                except OSError as exc:
                    return self._send(502, {"error": f"vision model unavailable: {exc}"})
            if path == "/api/import":
                entries = import_xlsx(self._body())
                with _lock, db() as con:
                    have = {(r["date"], r["odometer"], r["trip_miles"]) for r in rows(con)}
                    added = 0
                    for e in entries:
                        if (e["date"], e["odometer"], e["trip_miles"]) in have:
                            continue
                        insert(con, e)
                        added += 1
                return self._send(200, {"added": added, "skipped": len(entries) - added})
        except (ValueError, json.JSONDecodeError) as exc:
            return self._send(400, {"error": str(exc)})
        except Exception as exc:  # bad xlsx etc.
            return self._send(400, {"error": f"could not read file: {exc}"})
        self._send(404, {"error": "not found"})

    def _id(self):
        m = re.fullmatch(r"/api/fillups/(\d+)", urlparse(self.path).path)
        return int(m.group(1)) if m else None

    def do_PUT(self):
        if not (self._guard() and self._csrf()):
            return
        fid = self._id()
        if fid is None:
            return self._send(404, {"error": "not found"})
        try:
            e = clean(json.loads(self._body() or b"{}"))
        except (ValueError, json.JSONDecodeError) as exc:
            return self._send(400, {"error": str(exc)})
        if not e:
            return self._send(400, {"error": "nothing to update"})
        e["updated_at"] = now()
        with _lock, db() as con:
            cur = con.execute(
                f"UPDATE fillups SET {','.join(k + '=?' for k in e)} WHERE id=?",
                [*e.values(), fid])
            if not cur.rowcount:
                return self._send(404, {"error": "not found"})
            row = dict(con.execute("SELECT * FROM fillups WHERE id=?", (fid,)).fetchone())
        self._send(200, row)

    def do_DELETE(self):
        if not (self._guard() and self._csrf()):
            return
        fid = self._id()
        with _lock, db() as con:
            cur = con.execute("DELETE FROM fillups WHERE id=?", (fid,))
        self._send(200 if cur.rowcount else 404, {"deleted": cur.rowcount})


def main():
    if len(sys.argv) >= 3 and sys.argv[1] == "import":
        with open(sys.argv[2], "rb") as f:
            entries = import_xlsx(f.read())
        with db() as con:
            for e in entries:
                insert(con, e)
        print(f"imported {len(entries)} fill-ups into {DB_PATH}")
        return
    db().close()
    srv = ThreadingHTTPServer(("0.0.0.0", PORT), Handler)
    print(f"mpg-tracker on :{PORT}, db {DB_PATH}", flush=True)
    srv.serve_forever()


if __name__ == "__main__":
    main()
