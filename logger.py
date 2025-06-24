"""logger.py — Unified logging helper
====================================
Rotating logger that rolls over at 5 MB.

• Log directory defaults to Config.LOG_DIR read from config.py; falls
  back to ./logs if config is unavailable.
• Uses UTF-8 RotatingFileHandler (5 MB, 3 backups) + console stream.
"""

from __future__ import annotations

import logging
from logging.handlers import RotatingFileHandler
from pathlib import Path
from typing import Any


def _default_log_dir() -> Path:
    """Return log directory from Config.LOG_DIR or fallback to ./logs."""
    try:
        from config import Config  # local import to avoid circular dependency
        cfg = Config.load()
        return Path(cfg.LOG_DIR)
    except Exception:  # pragma: no cover
        return Path("logs")


class Logger:  # pylint: disable=too-few-public-methods
    """Thin wrapper around logging configured for Fund-Mail."""

    def __init__(
        self,
        name: str = "fund_mail",
        log_dir: str | Path | None = None,
        level: int = logging.INFO,
        max_bytes: int = 5_000_000,  # 5 MB
        backup_count: int = 3,
    ) -> None:
        self.log_dir = Path(log_dir) if log_dir else _default_log_dir()
        self.log_dir.mkdir(parents=True, exist_ok=True)

        log_file = self.log_dir / f"{name}.log"

        self._logger = logging.getLogger(name)
        self._logger.setLevel(level)
        self._logger.propagate = False  # do not bubble up to root

        # Prevent duplicate handlers when re-imported
        if not self._logger.handlers:
            formatter = logging.Formatter(
                "%(asctime)s | %(levelname)-8s | %(name)s | %(message)s",
                datefmt="%Y-%m-%d %H:%M:%S",
            )

            file_handler = RotatingFileHandler(
                log_file,
                maxBytes=max_bytes,
                backupCount=backup_count,
                encoding="utf-8",
            )
            file_handler.setFormatter(formatter)
            self._logger.addHandler(file_handler)

            console_handler = logging.StreamHandler()
            console_handler.setFormatter(formatter)
            self._logger.addHandler(console_handler)

    # ---- convenience passthroughs ----
    def __getattr__(self, item: str) -> Any:  # delegate to underlying logger
        return getattr(self._logger, item)

    @property
    def logger(self) -> logging.Logger:
        """Expose underlying logging.Logger."""
        return self._logger


# module-level default logger
logger: logging.Logger = Logger().logger
