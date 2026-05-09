# Daily Top Stories Checker

This app checks top stories for a list of news domains and stores the latest results.

## What it does

- Accepts one website domain per line.
- Fetches each domain's top story via Google News RSS (`site:domain`).
- Saves each run to a SQLite database.
- Exports latest run to `exports/top_stories_*.csv`.
- Saves your custom domain list for reuse.
- Supports domain list import/export.
- Runs automatic daily checks on a schedule.
- Shows a mobile-friendly web page.
- Supports "Add to Home Screen" as a basic PWA.

## Run locally

1. Install Python 3.10+.
2. Open terminal in this folder.
3. Install dependencies:
   - `python -m pip install -r requirements.txt`
4. Run:
   - `python app.py`
5. Open:
   - `http://127.0.0.1:5000`

## One-click first run (recommended on Windows)

- Double-click:
  - `start_app_first_run.bat`

What it does:

- Creates `.venv` if needed.
- Installs dependencies.
- Imports `domains_full.txt` into app settings on first run only.
- Starts the app.

## Daily schedule

- Default run time is `08:00` (server local time).
- Change it with environment variables:
  - `DAILY_RUN_HOUR` (0-23)
  - `DAILY_RUN_MINUTE` (0-59)

Example:

- `DAILY_RUN_HOUR=6`
- `DAILY_RUN_MINUTE=30`

## Import/export domain list

- In UI:
  - Use **Save Domain List Only** to save a large domain list without running.
- API/helper endpoint:
  - `GET /export-domains` writes `exports/domains_latest.txt`

## Build Windows EXE

Run:

- `build_exe.bat`

After build, use:

- `dist/topstories.exe`

## One-click deploy (Render)

This repo includes `render.yaml` and `Procfile`.

1. Push this folder to GitHub.
2. In Render, create a new Blueprint/Web Service from the repo.
3. Render auto-detects config and deploys.
4. Open the deployed URL on your phone.
5. Tap **Add to Home Screen** in your mobile browser.

## Important note about phone install

- A Windows `.exe` does **not** install on a phone browser.
- For phone use, run this as a hosted web app (PWA) and open that URL on your phone.
- Then tap browser menu -> **Add to Home Screen**.
