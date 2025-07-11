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
from config import Config
from fetch_email import fetch_email_from_office365
from sftp_upload import upload_to_sftp  # noqa: E402

           # 기본 .env 로드
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
    logger.info(f"=" * 59)
    logger.info(f"✅ {date_str}(KST)의 모든 메일 가져오기")
    logger.info(f"=" * 59)
    try:
        # 2️⃣ Load project configuration (credentials, paths, etc.)
        cfg = Config.load() 
        db_path = fetch_email_from_office365(cfg, one_day=date_str)
        if db_path:
            upload_to_sftp(cfg, db_path)    
        logger.info("=" * 59)
        logger.info("✅ {date_str}(KST) 완료: %s", datetime.now())
        logger.info("=" * 59)
    except Exception as exc:
        logger.info(f"=" * 59)
        logger.exception("⛔ fund메일 작업 중 예외 발생 – 프로세스 종료")
        logger.info(f"=" * 59)
        sys.exit(1)


if __name__ == "__main__":  # pragma: no cover
    main()
