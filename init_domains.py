import sqlite3
from pathlib import Path

BASE_DIR = Path(__file__).parent
DB_PATH = BASE_DIR / "top_stories.db"
DOMAINS_FILE = BASE_DIR / "domains_full.txt"


def normalize_domains(raw_text: str) -> list[str]:
    domains = []
    for line in raw_text.splitlines():
        d = line.strip().lower()
        if not d:
            continue
        d = d.replace("https://", "").replace("http://", "").strip("/")
        domains.append(d)
    return list(dict.fromkeys(domains))


def main() -> None:
    if not DOMAINS_FILE.exists():
        raise FileNotFoundError(f"Missing domains file: {DOMAINS_FILE}")

    domains = normalize_domains(DOMAINS_FILE.read_text(encoding="utf-8"))
    if not domains:
        raise RuntimeError("No domains found in domains_full.txt")

    conn = sqlite3.connect(DB_PATH)
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS settings (
            key TEXT PRIMARY KEY,
            value TEXT NOT NULL
        )
        """
    )
    conn.execute(
        """
        INSERT INTO settings (key, value) VALUES ('domains', ?)
        ON CONFLICT(key) DO UPDATE SET value = excluded.value
        """,
        ("\n".join(domains),),
    )
    conn.commit()
    conn.close()
    print(f"Imported {len(domains)} domains into settings.")


if __name__ == "__main__":
    main()
