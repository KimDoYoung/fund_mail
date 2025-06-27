import shutil
import sys
from datetime import datetime

from config import Config  # noqa: E402
from fetch_email import fetch_email_from_office365
from sftp_upload import upload_to_sftp  # noqa: E402
from logger import logger


def main():
    """단발성(fire‑and‑exit) fund 메일 수집/업로드 진입점."""
    cfg = Config.load()
    backup_path = None

    logger.info("=" * 59)
    logger.info("⏺️ fund메일 수집 시작: %s", datetime.now())
    logger.info("=" * 59)

    try:
        # LAST_TIME.json 백업
        if cfg.last_time_file.exists():
            backup_path = cfg.last_time_file.with_suffix(cfg.last_time_file.suffix + ".previous")
            shutil.copy2(cfg.last_time_file, backup_path)

        # 메일 수집 후 SFTP 업로드
        db_path = fetch_email_from_office365(cfg)
        if db_path:
            upload_to_sftp(cfg, db_path)

        # 정상 완료 시 백업 삭제
        if backup_path and backup_path.exists():
            backup_path.unlink()

        logger.info("=" * 59)
        logger.info("✅ fund메일 수집 완료: %s", datetime.now())
        logger.info("=" * 59)

    except Exception:
        logger.exception("⛔ fund메일 작업 중 예외 발생 – 프로세스 종료")
        # 실패 시 백업 복구
        if backup_path and backup_path.exists():
            shutil.copy2(backup_path, cfg.last_time_file)
            logger.warning("⚠️ 백업된 LAST_TIME.json 복구: %s", backup_path)
        sys.exit(1)


if __name__ == "__main__":
    main()
