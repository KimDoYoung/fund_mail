"""config.py — 환경 변수 로더
================================
사전에 `.env` 파일(또는 시스템 환경 변수)에 정의된 값들을 읽어들여
`Config` 데이터클래스로 제공한다.  
사용 예::

    from config import Config
    cfg = Config.load()            # 기본은 .env
    db_path = cfg.db_path_for()    # fm_yyyy_mm_dd_HHMM.db 전체 경로
"""
from __future__ import annotations

import os
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path

from dotenv import load_dotenv

__all__ = ["Config"]


@dataclass(slots=True, frozen=True)
class Config:
    # ──────────────────────────────── Office 365 ──────────────────────────────
    email_id: str
    email_pw: str
    tenant_id: str
    client_id: str
    client_secret: str

    # ──────────────────────────────── 로컬 저장소 ─────────────────────────────
    base_dir: Path           # ex) C:\fund_mail\data
    attach_base_dir: Path    # ex) C:\fund_mail\data\attach

    # ──────────────────────────────── SFTP ────────────────────────────────────
    sftp_host: str
    sftp_port: int
    sftp_id: str
    sftp_pw: str
    log_dir: Path            # ex) C:\fund_mail\logs
    # ──────────────────────────────── 클래스메서드 ────────────────────────────
    @classmethod
    def load(cls, env_file: str | Path = ".env") -> "Config":
        """`.env` → :class:`Config` 인스턴스 반환."""
        load_dotenv(env_file)

        required = {
            "EMAIL_ID": str,
            "EMAIL_PW": str,
            "TENANT_ID": str,
            "CLIENT_ID": str,
            "CLIENT_SECRET": str,
            "BASE_DIR": str,
            "ATTACH_BASE_DIR": str,
            "HOST": str,
            "PORT": int,
            "SFTP_ID": str,
            "SFTP_PW": str,
            "LOG_DIR": str,  # 로그 디렉터리 추가
        }

        missing = [k for k in required if os.getenv(k) is None]
        if missing:
            raise EnvironmentError(
                f"[Config] .env 에 다음 값이 없습니다: {', '.join(missing)}"
            )

        # 형 변환 -------------------------------------------------------------
        def _cast(key: str, to_type):
            val = os.getenv(key)
            return to_type(val) if to_type is int else val

        return cls(
            email_id=_cast("EMAIL_ID", str),
            email_pw=_cast("EMAIL_PW", str),
            tenant_id=_cast("TENANT_ID", str),
            client_id=_cast("CLIENT_ID", str),
            client_secret=_cast("CLIENT_SECRET", str),
            base_dir=Path(_cast("BASE_DIR", str)).expanduser().resolve(),
            attach_base_dir=Path(_cast("ATTACH_BASE_DIR", str)).expanduser().resolve(),
            sftp_host=_cast("HOST", str),
            sftp_port=_cast("PORT", int),
            sftp_id=_cast("SFTP_ID", str),
            sftp_pw=_cast("SFTP_PW", str),
            log_dir=Path(_cast("LOG_DIR", str)).expanduser().resolve(),
        )

    # ──────────────────────────────── 헬퍼 메서드 ─────────────────────────────
    def db_name_for(self, ts: datetime | None = None) -> str:
        ts = ts or datetime.utcnow()
        return f"fm_{ts:%Y_%m_%d_%H%M}.db"

    def db_path_for(self, ts: datetime | None = None) -> Path:
        return self.base_dir / self.db_name_for(ts)

    @property
    def cursor_file(self) -> Path:
        """마지막 수집 시각을 저장하는 `LAST_TIME.txt` 경로."""
        return self.base_dir / "LAST_TIME.txt"
