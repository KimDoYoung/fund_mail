"""logger.py — 통합 회전 로거(5 MB)
========================================
• 로그 디렉터리: `Config.LOG_DIR` (`.env` 의 `LOG_DIR`) 값을 우선 사용하고,
  지정되지 않은 경우 `./logs` 로 자동 대체합니다.
• 디렉터리가 없으면 `mkdir(parents=True, exist_ok=True)` 로 자동 생성합니다.
• 로그 파일이 5 MB 를 초과하면 최대 3개까지 순환 보관(`RotatingFileHandler`).
• `get_logger(name)` 로 원하는 이름의 로거를 얻거나, 모듈 전역 `logger`
  를 바로 사용하세요.

사용 예::

    from logger import logger, get_logger

    logger.info("Fund‑Mail 시작!")

    mail_logger = get_logger("fund_mail.worker")
    mail_logger.debug("Processing batch…")
"""
from __future__ import annotations

import logging
from logging.handlers import RotatingFileHandler
from pathlib import Path
from typing import Optional

# ─────────────────────────────── 설정 상수 ────────────────────────────────
_MAX_BYTES: int = 5 * 1024 * 1024   # 5 MB
_BACKUP_COUNT: int = 3

# ────────────────────────────── 내부 헬퍼 ─────────────────────────────────

def _determine_log_dir() -> Path:
    """`Config.LOG_DIR` 값을 확인하여 Path 를 반환한다. 지정되지 않았거나
    예외가 발생하면 ./logs 로 폴백한다."""
    try:
        from config import Config  # runtime import — 순환 의존 방지
        cfg = Config.load()
        log_dir: Optional[Path] = getattr(cfg, "log_dir", None)
        if log_dir:
            return Path(log_dir).expanduser().resolve()
    except Exception:  # pylint: disable=broad-except
        pass
    return Path("./logs").expanduser().resolve()


def _make_handlers(log_file: Path) -> list[logging.Handler]:
    fmt = logging.Formatter("%(asctime)s [%(levelname).1s] %(name)s: %(message)s")

    file_handler = RotatingFileHandler(
        log_file,
        maxBytes=_MAX_BYTES,
        backupCount=_BACKUP_COUNT,
        encoding="utf-8",
    )
    file_handler.setFormatter(fmt)

    console_handler = logging.StreamHandler()
    console_handler.setFormatter(fmt)

    return [file_handler, console_handler]


# ─────────────────────────────── API ─────────────────────────────────────

def get_logger(name: str = "fund_mail") -> logging.Logger:
    """콘솔과 회전 로그 파일에 동시에 기록하는 로거를 반환한다.

    동일한 이름으로 반복 호출하면 중복 핸들러 없이 동일 로거를
    반환한다."""
    logger = logging.getLogger(name)
    if logger.handlers:  # 이미 초기화 완료
        return logger

    logger.setLevel(logging.INFO)

    log_dir = _determine_log_dir()
    log_dir.mkdir(parents=True, exist_ok=True)  # 디렉터리 자동 생성

    handlers = _make_handlers(log_dir / f"{name}.log")
    for h in handlers:
        logger.addHandler(h)

    logger.propagate = False  # 루트로 전파 방지
    return logger


# 모듈 전역 기본 로거 (이 파일 import 시 즉시 준비)
logger = get_logger()
