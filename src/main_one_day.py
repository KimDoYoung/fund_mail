"""
main_one_day.py

사용법
-----
$ python main_one_day.py --date 2025-06-30
$ python main_one_day.py               # (defaults to today, KST)

날짜 2025-06-30과 같이 인자로 받아서, 인자가 없으면 오늘 날짜로 해서 받은 메일 전부를 다운로드받는다.
"""
from __future__ import annotations

import argparse
import sys
from datetime import datetime

from logger import get_logger

try:
    from zoneinfo import ZoneInfo  # Python 3.9+
    KST = ZoneInfo("Asia/Seoul")
except ModuleNotFoundError:  # pragma: no cover (≤3.8 fallback)
    import pytz
    KST = pytz.timezone("Asia/Seoul")

# ----------------------------------------------------------------------------
# fund_mail internals (import paths may differ in your project layout)
# ----------------------------------------------------------------------------
from config import load_config
from fetch_email import fetch_email_from_office365

logger = get_logger()
# ---------------------------------------------------------------------------
# CLI helpers
# ---------------------------------------------------------------------------

def _parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Fetch all messages that arrived during a single KST date.",
    )
    parser.add_argument(
        "-d",
        "--date",
        metavar="YYYY-MM-DD",
        help="Target date in KST (default: today)",
    )
    return parser.parse_args()


# ---------------------------------------------------------------------------
# Main routine
# ---------------------------------------------------------------------------

def main() -> None:
    args = _parse_args()

    # 1️⃣ Determine the date string (YYYY‑MM‑DD, in KST)
    if args.date:
        try:
            datetime.strptime(args.date, "%Y-%m-%d")  # validation only
        except ValueError as exc:
            print("ERROR: --date must be in 'YYYY-MM-DD' format (e.g. 2025-06-30)",
                  file=sys.stderr)
            sys.exit(1)
        date_str = args.date
    else:
        date_str = datetime.now(KST).strftime("%Y-%m-%d")

    logger.info(f"[fund_mail] Fetching e‑mails for {date_str} (KST)…")

    # 2️⃣ Load project configuration (credentials, paths, etc.)
    cfg = load_config()

    db_path = fetch_email_from_office365(cfg, date=date_str)

    logger.info(f"[fund_mail] Saved {date_str} e‑mails to → {db_path}")


if __name__ == "__main__":  # pragma: no cover
    main()
