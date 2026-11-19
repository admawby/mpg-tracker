# MPG Tracker

A phone-first fuel log: record each fill-up in a few taps, review your history and stats, and export everything to Excel.

- Fill-ups are pre-filled from your history: today's date, last octane, drive mode and station, and trip miles worked out from the odometer (or the other way round).
- Tag chips for things that affect mileage (Towing, Road trip, Cold weather…), notes, optional price and station.
- History by month with search, and tap to edit or delete.
- Stats: average and pump-math MPG, best and worst tank, miles, fuel, tank range, days between fills, averages by drive mode and tag, and an MPG-per-tank chart.
- **Service log** for maintenance (oil changes, tire rotations, brakes, filters…) with cost, shop and notes, plus **mileage reminders**: each category has an interval you can edit, and current mileage is estimated from your last odometer reading and recent miles per day. Due-soon items show on the Log screen, and services appear as markers on the MPG chart.
- Excel export, and import from an existing sheet (fuel columns `Date, MPG, Gallons, Time, Miles, Fuel, Drive Mode, Cum Miles, Notes`, plus an optional `Service` sheet with `Date, Odometer, Service, Cost, Shop, Notes`).
- Installable on an iPhone or Android home screen, and works offline.

There are two ways to run it.

## 1. No server (data stays on your phone)

Open **https://admawby.github.io/mpg-tracker/** in Safari (or Chrome on Android), then **Share → Add to Home Screen**.

- Fill-ups are stored in the browser's on-device storage (IndexedDB). Nothing is sent anywhere.
- **Use it from the Home Screen icon.** Data in a regular Safari tab can be cleared by iOS after a few weeks without use; the home-screen app is exempt. Removing the icon deletes the data.
- After each new fill-up or service the app offers **Back up**, which opens the share sheet with an `.xlsx` file. Save it to Files or iCloud Drive. You can turn the reminder off in Stats → Settings.
- To restore or move phones, use Stats → **Import .xlsx** with a backup file.
- Set your vehicle name in Stats → Settings.

You can host the `static/` folder anywhere that serves static files. `.github/workflows/pages.yml` publishes it to GitHub Pages on every push to `main` (enable Pages with source **GitHub Actions**).

## 2. Self-hosted server (one shared copy, automatic backups)

A small Python server (stdlib + SQLite, openpyxl for Excel) for running on a home server with Docker.

```sh
docker compose up -d --build      # http://<host>:8840
```

Settings come from environment variables (or a `.env` file next to `compose.yml`):

| Variable | Default | What it does |
|---|---|---|
| `MPG_VEHICLE` | `My Truck` | Name shown in the app and export |
| `MPG_BACKUP_PATH` | `./backup` | Host folder that gets `<vehicle>_MPG.xlsx` + `history/mpg_<time>.sqlite` (last 60) after each **new** fill-up or service. A NAS share works. |
| `TZ` | `UTC` | Time zone for timestamps |
| `OLLAMA_URL` | *(empty)* | Optional photo scan: an [Ollama](https://ollama.com) server with a vision model. When set, **Scan a photo** reads the trip computer, pump or fuel receipt, and **Scan a receipt** on the Service tab reads a shop invoice (date, shop, mileage, total, service). |
| `VISION_MODEL` | `qwen3-vl:8b` | Ollama model used for photo scan |

Anything specific to your machine (extra Docker networks, absolute paths) belongs in a `compose.override.yml`, which is gitignored.

The server only accepts requests from private, CGNAT/Tailscale and Docker networks. Keep it off the open internet: use a VPN such as Tailscale to reach it away from home. Service workers need HTTPS, so for offline use and home-screen install reach it over HTTPS (for example `tailscale serve`).

- Import from the command line: `docker compose run --rm mpg-tracker python app.py import /data/file.xlsx`
- API: `GET api/fillups` (returns fill-ups, services and settings), `POST api/fillups|services`, `PUT/DELETE api/fillups|services/<id>`, `PUT api/settings` (`{"intervals": {...}}`), `GET api/export.xlsx`, `POST api/import` (raw .xlsx body), `POST api/scan[?mode=service]` (JPEG body). Writes need the header `X-MPG: 1`.
- `deploy.sh` copies the app to a Docker host over SSH and rebuilds it. Set `DEPLOY_HOST` and `DEPLOY_DIR` in the environment or a gitignored `.deploy.env`.

## How the modes are chosen

`static/config.js` says `{mode: "local"}`. The Python server answers `config.js` itself with `{mode: "server", …}`, so the same front end runs either way. `static/store.js` holds both storage back ends and the in-browser Excel reader and writer.

When you change the UI, bump `VERSION` in `static/sw.js` so installed copies pick up the update.
