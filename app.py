import csv
import datetime as dt
import os
import sqlite3
from pathlib import Path
from typing import Iterable

import feedparser
from apscheduler.schedulers.background import BackgroundScheduler
from flask import Flask, jsonify, redirect, render_template, request, url_for

BASE_DIR = Path(__file__).parent
DB_PATH = BASE_DIR / "top_stories.db"
DEFAULT_DOMAINS = [
    "cnn.com",
    "foxnews.com",
    "nytimes.com",
    "wsj.com",
    "washingtonpost.com",
    "latimes.com",
    "chicagotribune.com",
]

app = Flask(__name__)
scheduler = BackgroundScheduler(daemon=True)


def get_conn() -> sqlite3.Connection:
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


def init_db() -> None:
    conn = get_conn()
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS top_stories (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            fetched_at TEXT NOT NULL,
            domain TEXT NOT NULL,
            title TEXT NOT NULL,
            link TEXT NOT NULL,
            source TEXT NOT NULL
        )
        """
    )
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS runs (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            run_at TEXT NOT NULL,
            domains_count INTEGER NOT NULL
        )
        """
    )
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS settings (
            key TEXT PRIMARY KEY,
            value TEXT NOT NULL
        )
        """
    )
    conn.commit()
    conn.close()


def normalize_domains(raw_text: str) -> list[str]:
    domains = []
    for line in raw_text.splitlines():
        d = line.strip().lower()
        if not d:
            continue
        d = d.replace("https://", "").replace("http://", "").strip("/")
        domains.append(d)
    # Keep order, remove duplicates.
    return list(dict.fromkeys(domains))


def get_saved_domains() -> list[str]:
    conn = get_conn()
    row = conn.execute("SELECT value FROM settings WHERE key = 'domains'").fetchone()
    conn.close()
    if not row:
        return DEFAULT_DOMAINS
    domains = normalize_domains(row["value"])
    return domains or DEFAULT_DOMAINS


def save_domains(domains: list[str]) -> None:
    raw = "\n".join(domains)
    conn = get_conn()
    conn.execute(
        """
        INSERT INTO settings (key, value) VALUES ('domains', ?)
        ON CONFLICT(key) DO UPDATE SET value = excluded.value
        """,
        (raw,),
    )
    conn.commit()
    conn.close()


def google_news_rss_for_domain(domain: str) -> str:
    return (
        "https://news.google.com/rss/search?q="
        f"site:{domain}&hl=en-US&gl=US&ceid=US:en"
    )


def fetch_top_story(domain: str) -> dict:
    rss_url = google_news_rss_for_domain(domain)
    feed = feedparser.parse(rss_url)
    if feed.entries:
        top = feed.entries[0]
        return {
            "domain": domain,
            "title": top.get("title", "(No title)"),
            "link": top.get("link", ""),
            "source": "Google News RSS",
        }
    return {
        "domain": domain,
        "title": "No story found",
        "link": "",
        "source": "Google News RSS",
    }


def run_fetch_for_domains(domains: list[str]) -> None:
    rows = [fetch_top_story(d) for d in domains]
    save_results(rows)
    export_latest_csv()


def save_results(rows: Iterable[dict]) -> None:
    now = dt.datetime.now().isoformat(timespec="seconds")
    conn = get_conn()
    count = 0
    for row in rows:
        conn.execute(
            """
            INSERT INTO top_stories (fetched_at, domain, title, link, source)
            VALUES (?, ?, ?, ?, ?)
            """,
            (now, row["domain"], row["title"], row["link"], row["source"]),
        )
        count += 1
    conn.execute("INSERT INTO runs (run_at, domains_count) VALUES (?, ?)", (now, count))
    conn.commit()
    conn.close()


def latest_rows() -> list[sqlite3.Row]:
    conn = get_conn()
    run = conn.execute("SELECT run_at FROM runs ORDER BY id DESC LIMIT 1").fetchone()
    if not run:
        conn.close()
        return []
    rows = conn.execute(
        """
        SELECT fetched_at, domain, title, link, source
        FROM top_stories
        WHERE fetched_at = ?
        ORDER BY domain
        """,
        (run["run_at"],),
    ).fetchall()
    conn.close()
    return rows


def export_latest_csv() -> Path | None:
    rows = latest_rows()
    if not rows:
        return None
    out_dir = BASE_DIR / "exports"
    out_dir.mkdir(exist_ok=True)
    stamp = rows[0]["fetched_at"].replace(":", "-")
    out_path = out_dir / f"top_stories_{stamp}.csv"
    with out_path.open("w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow(["fetched_at", "domain", "title", "link", "source"])
        for r in rows:
            writer.writerow([r["fetched_at"], r["domain"], r["title"], r["link"], r["source"]])
    return out_path


@app.route("/", methods=["GET"])
def home():
    rows = latest_rows()
    conn = get_conn()
    last_run = conn.execute("SELECT run_at FROM runs ORDER BY id DESC LIMIT 1").fetchone()
    conn.close()
    domains = get_saved_domains()
    schedule_time = f"{os.getenv('DAILY_RUN_HOUR', '08')}:{os.getenv('DAILY_RUN_MINUTE', '00')}"
    return render_template(
        "index.html",
        default_domains="\n".join(domains),
        rows=rows,
        last_run=last_run["run_at"] if last_run else None,
        schedule_time=schedule_time,
    )


@app.route("/run", methods=["POST"])
def run_now():
    domains_raw = request.form.get("domains", "")
    domains = normalize_domains(domains_raw) or get_saved_domains()
    save_domains(domains)
    run_fetch_for_domains(domains)
    return redirect(url_for("home"))


@app.route("/import", methods=["POST"])
def import_domains():
    domains_text = request.form.get("domains", "")
    domains = normalize_domains(domains_text)
    if domains:
        save_domains(domains)
    return redirect(url_for("home"))


@app.route("/export-domains", methods=["GET"])
def export_domains():
    domains = get_saved_domains()
    out_dir = BASE_DIR / "exports"
    out_dir.mkdir(exist_ok=True)
    out_path = out_dir / "domains_latest.txt"
    out_path.write_text("\n".join(domains), encoding="utf-8")
    return jsonify({"file": str(out_path), "count": len(domains)})


@app.route("/api/top-stories", methods=["GET"])
def top_stories_api():
    rows = latest_rows()
    data = [dict(r) for r in rows]
    return jsonify(data)


def start_scheduler() -> None:
    hour = int(os.getenv("DAILY_RUN_HOUR", "8"))
    minute = int(os.getenv("DAILY_RUN_MINUTE", "0"))

    def scheduled_job() -> None:
        domains = get_saved_domains()
        run_fetch_for_domains(domains)

    if not scheduler.running:
        scheduler.add_job(
            scheduled_job,
            "cron",
            id="daily_top_stories_run",
            hour=hour,
            minute=minute,
            replace_existing=True,
        )
        scheduler.start()


def boot() -> None:
    init_db()
    save_domains(get_saved_domains())
    start_scheduler()


if __name__ == "__main__":
    boot()
    app.run(host="0.0.0.0", port=5000, debug=False)


if __name__ != "__main__":
    boot()
