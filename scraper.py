"""
Matka result scraper.
Fetches today's results from a configurable source URL and writes to matka.db.

Usage from CLI:
    python scraper.py              # fetch today
    python scraper.py 2026-06-01   # fetch a specific date

Usage from app:
    from scraper import run_fetch
    result = run_fetch(date_str=None, url=None)

Source compatibility:
- Targets dpboss-style HTML tables (the watermark seen on user's screenshots).
- Generic fallback: scans for "MarketName ... 123-45-678" patterns.
- Configurable via SOURCE_URL or `source_url` arg.

NOTE: Web scraping may violate the source site's Terms of Service.
Check the source's robots.txt and ToS before using in production.
"""
import re
import sys
import sqlite3
import datetime as dt
from urllib.request import Request, urlopen
from urllib.error import URLError, HTTPError
import json

DB_PATH = "matka.db"
DEFAULT_SOURCE_URL = "https://dpboss.services/"  # change in Settings if needed
USER_AGENT = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0 Safari/537.36"

# Map source-site market name variations -> our canonical DB market name
ALIAS_MAP = {
    "sita morning": "Sita Morning",
    "sita day": "Sita Day",
    "sita night": "Sita Night",
    "geeta morning": "Geeta Morning",
    "star tara morning": "Star Tara Morning",
    "star tara day": "Star Tara Day",
    "star tara night": "Star Tara Night",
    "tulsi morning": "Tulsi Morning",
    "andhra morning": "Andra Morning",
    "andra morning": "Andra Morning",
    "andhra day": "Andra Day",
    "andra day": "Andra Day",
    "andhra night": "Andra Night",
    "andra night": "Andra Night",
    "meena morning": "Meena Morning",
    "meena bazar": "Meena Bazaar",
    "meena bazaar": "Meena Bazaar",
    "sridevi": "Sridevi",
    "sridevi day": "Sridevi",
    "sridevi night": "Sridevi Night",
    "mahadevi morning": "Mahadevi Morning",
    "mahadevi": "Mahadevi",
    "mahadevi day": "Mahadevi",
    "mahadevi night": "Mahadevi Night",
    "time bazar": "Time Bazaar",
    "time bazaar": "Time Bazaar",
    "madhur day": "Madhur Day",
    "madhur night": "Madhur Night",
    "srilakhsmi day": "Srilakhsmi Day",
    "srilaxmi day": "Srilakhsmi Day",
    "sridevi laxmi": "Srilakhsmi Day",
    "milan day": "Milan Day",
    "milan night": "Milan Night",
    "rajdhani day": "Rajdhani Day",
    "rajdhani night": "Rajdhani Night",
    "kalyan": "Kalyan",
    "kalyan day": "Kalyan",
    "kalyan night": "Kalyan Night",
    "super king": "Super King",
    "superking": "Super King",
    "super king night": "Superking Night",
    "superking night": "Superking Night",
    "main bazar": "Main Bazaar",
    "main bazaar": "Main Bazaar",
}


def get_conn():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


def ensure_fetch_log_table():
    conn = get_conn()
    conn.execute("""CREATE TABLE IF NOT EXISTS fetch_log (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        ts TEXT DEFAULT CURRENT_TIMESTAMP,
        date TEXT,
        url TEXT,
        markets_found INTEGER,
        markets_saved INTEGER,
        status TEXT,
        message TEXT
    )""")
    conn.commit()


def fetch_html(url):
    """Fetch a URL, return HTML text."""
    req = Request(url, headers={"User-Agent": USER_AGENT, "Accept-Language": "en-US,en;q=0.9"})
    with urlopen(req, timeout=20) as resp:
        return resp.read().decode("utf-8", errors="ignore")


def parse_results(html):
    """
    Parse HTML and extract market_name -> jodi (2 digits).
    Strategy 1: look for patterns "MarketName <stuff> 123-45-678" (panel-jodi-panel format).
    Strategy 2: simple "MarketName 45" pairs.
    Returns dict: {canonical_market_name: jodi_str}
    """
    results = {}
    # Strip HTML tags to plain text
    text = re.sub(r"<style.*?</style>", " ", html, flags=re.S | re.I)
    text = re.sub(r"<script.*?</script>", " ", text, flags=re.S | re.I)
    text = re.sub(r"<[^>]+>", " ", text)
    text = re.sub(r"\s+", " ", text)
    text = text.strip()

    # Pattern 1: market name followed by panel-jodi-panel (e.g. "Kalyan 145-08-466")
    pat1 = re.compile(
        r"([A-Z][A-Za-z\s]{2,30}?)\s+(\d{3})\s*[-–]\s*(\d{2})\s*[-–]\s*(\d{3})",
        re.IGNORECASE,
    )
    for m in pat1.finditer(text):
        raw_name = m.group(1).strip().lower()
        jodi = m.group(3)
        canonical = ALIAS_MAP.get(raw_name)
        if not canonical:
            # Try fuzzy match: strip extra words
            for alias, mapped in ALIAS_MAP.items():
                if alias in raw_name:
                    canonical = mapped
                    break
        if canonical and len(jodi) == 2 and jodi.isdigit():
            results[canonical] = jodi

    # Pattern 2: "MarketName : 45" or "MarketName - 45"
    pat2 = re.compile(
        r"([A-Z][A-Za-z\s]{2,30}?)\s*[:\-=]\s*(\d{2})\b(?!\d)",
    )
    for m in pat2.finditer(text):
        raw_name = m.group(1).strip().lower()
        jodi = m.group(2)
        if raw_name in ALIAS_MAP and ALIAS_MAP[raw_name] not in results:
            if len(jodi) == 2 and jodi.isdigit():
                results[ALIAS_MAP[raw_name]] = jodi

    return results


def save_results(results, date_str):
    """Save scraped {market: jodi} into DB. Returns count saved."""
    conn = get_conn()
    saved = 0
    for market_name, jodi in results.items():
        row = conn.execute("SELECT id FROM markets WHERE name=?", (market_name,)).fetchone()
        if not row:
            continue
        try:
            conn.execute(
                "INSERT OR IGNORE INTO jodis (market_id, date, jodi) VALUES (?,?,?)",
                (row["id"], date_str, jodi.zfill(2)),
            )
            if conn.total_changes > 0:
                saved += 1
            conn.execute(
                "INSERT INTO audit_log (user, action, details) VALUES (?,?,?)",
                ("scraper", "auto_fetch", f"{market_name} {date_str}={jodi}"),
            )
        except Exception:
            pass
    conn.commit()
    return saved


def run_fetch(date_str=None, url=None):
    """
    Main entry point. Returns dict with:
      ok, date, url, markets_found, markets_saved, status, message
    """
    ensure_fetch_log_table()
    if date_str is None:
        date_str = dt.date.today().isoformat()
    if url is None:
        url = DEFAULT_SOURCE_URL
    info = {"ok": False, "date": date_str, "url": url,
            "markets_found": 0, "markets_saved": 0,
            "status": "error", "message": "", "results": {}}
    try:
        html = fetch_html(url)
    except HTTPError as e:
        info["message"] = f"HTTP {e.code}: {e.reason}"
        _log_fetch(info); return info
    except URLError as e:
        info["message"] = f"URL error: {e.reason}"
        _log_fetch(info); return info
    except Exception as e:
        info["message"] = f"Fetch failed: {e}"
        _log_fetch(info); return info

    results = parse_results(html)
    info["markets_found"] = len(results)
    info["results"] = results
    if not results:
        info["status"] = "no_data"
        info["message"] = "Source returned 0 markets parsed. Page format may have changed."
        _log_fetch(info); return info

    saved = save_results(results, date_str)
    info["markets_saved"] = saved
    info["ok"] = True
    info["status"] = "ok"
    info["message"] = f"Found {len(results)}, saved {saved} new entries."
    _log_fetch(info)
    return info


def _log_fetch(info):
    conn = get_conn()
    conn.execute(
        "INSERT INTO fetch_log (date, url, markets_found, markets_saved, status, message) VALUES (?,?,?,?,?,?)",
        (info["date"], info["url"], info["markets_found"], info["markets_saved"],
         info["status"], info["message"][:500]),
    )
    conn.commit()


if __name__ == "__main__":
    date_arg = sys.argv[1] if len(sys.argv) > 1 else None
    url_arg = sys.argv[2] if len(sys.argv) > 2 else None
    result = run_fetch(date_arg, url_arg)
    print(json.dumps({k: v for k, v in result.items() if k != "results"}, indent=2))
    if result["results"]:
        print("\nResults:")
        for market, jodi in sorted(result["results"].items()):
            print(f"  {market}: {jodi}")
