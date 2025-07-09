"""config.py — 환경 변수 로더
================================
`.env` 파일(또는 시스템 환경 변수) 값을 읽어 `Config` 데이터클래스로
제공합니다. DB·첨부·로그 경로, 마지막 메일 수집 시각 등 자주 쓰는
헬퍼를 모두 포함했습니다.

사용 예::

    from config import Config
    cfg = Config.load()            # 기본 .env 로드
    db_path = cfg.db_path_for()    # fm_YYYY_MM_DD_HHMM.db 전체 경로
    cursor  = cfg.last_mail_fetch_time  # datetime(UTC)
"""
from __future__ import annotations

import json
import os
from dataclasses import dataclass
from datetime import datetime, timedelta, timezone
from pathlib import Path
from typing import Any

from dotenv import load_dotenv

__all__ = ["Config"]
UTC = timezone.utc
KST = timezone(timedelta(hours=9))


@dataclass(slots=True, frozen=True)
class Config:
    # ───────────────────────────── Office 365 ──────────────────────────────
    email_user_id: str
    email_pw: str
    tenant_id: str
    client_id: str
    client_secret: str

    # ───────────────────────────── 로컬 저장소 ─────────────────────────────
    data_dir: Path           # 예) C:\fund_mail\data
    log_dir: Path            # 예) C:\fund_mail\logs

    # ───────────────────────────── SFTP ────────────────────────────────────
    sftp_host: str
    sftp_port: int
    sftp_id: str
    sftp_pw: str
    sftp_base_dir: str 

    # ──────────────────────────── 클래스메서드 ────────────────────────────
    @classmethod
    def load(cls, env_file: str | Path = ".env") -> "Config":
        """`.env` → :class:`Config` 인스턴스를 반환."""
        load_dotenv(env_file)

        required: dict[str, type] = {
            "EMAIL_ID": str,
            "EMAIL_PW": str,
            "TENANT_ID": str,
            "CLIENT_ID": str,
            "CLIENT_SECRET": str,
            "DATA_DIR": str,
            "LOG_DIR": str,
            "HOST": str,
            "PORT": int,
            "SFTP_ID": str,
            "SFTP_PW": str,
            "SFTP_BASE_DIR": str,  # SFTP 업로드할 기본 디렉토리
        }

        missing = [k for k in required if os.getenv(k) is None]
        if missing:
            raise EnvironmentError(
                f"[Config] .env 에 다음 값이 없습니다: {', '.join(missing)}"
            )

        def _cast(key: str, to_type):
            val = os.getenv(key)
            return to_type(val) if to_type is int else val

        return cls(
            email_user_id=_cast("EMAIL_ID", str),
            email_pw=_cast("EMAIL_PW", str),
            tenant_id=_cast("TENANT_ID", str),
            client_id=_cast("CLIENT_ID", str),
            client_secret=_cast("CLIENT_SECRET", str),
            data_dir=Path(_cast("DATA_DIR", str)).expanduser().resolve(),
            log_dir=Path(_cast("LOG_DIR", str)).expanduser().resolve(),
            sftp_host=_cast("HOST", str),
            sftp_port=_cast("PORT", int),
            sftp_id=_cast("SFTP_ID", str),
            sftp_pw=_cast("SFTP_PW", str),
            sftp_base_dir=_cast("SFTP_BASE_DIR", str),
        )

    # ───────────────────────────── 헬퍼 메서드 ─────────────────────────────
    def db_name_for(self, ts: datetime | None = None) -> str:
        ts = ts or datetime.utcnow()
        return f"fm_{ts:%Y_%m_%d_%H%M}.db"

    def db_path_for(self, ts: datetime | None = None) -> Path:
        return self.data_dir / self.db_name_for(ts)

    @property
    def last_time_file(self) -> Path:
        """`LAST_TIME.json` 전체 경로."""
        if not self.data_dir.exists():
            self.data_dir.mkdir(parents=True, exist_ok=True)
        return self.data_dir / "LAST_TIME.json"

    # ──────────────────────── 커서 로딩 헬퍼 ────────────────────────────
    @property
    def last_mail_fetch_time(self) -> datetime:
        """마지막 이메일 수집 시각을 UTC 로 반환.

        1. `LAST_TIME.json` 의 `{"last_fetch_time": "ISO‑UTC"}` 값을 읽습니다.
        2. 파일이 없거나 파싱에 실패하면 **오늘 00:00(KST)** 를 UTC 로 변환해
           반환합니다.
        """
        if self.last_time_file.exists():
            try:
                with self.last_time_file.open("r", encoding="utf-8") as fh:
                    data: dict[str, Any] = json.load(fh)
                ts_str = data.get("last_fetch_time")
                if ts_str:
                    ts = datetime.fromisoformat(ts_str.replace("Z", "+00:00"))
                    if ts.tzinfo is None:
                        ts = ts.replace(tzinfo=UTC)
                    return ts.astimezone(UTC)
            except Exception:  # noqa: BLE001 (폭넓게 잡고 폴백)
                pass  # 폴백으로 이동

        # 폴백: 오늘 00:00 KST → UTC
        midnight_local = datetime.now(KST).replace(hour=0, minute=0, second=0, microsecond=0)
        return midnight_local.astimezone(UTC)

    @property
    def last_email_id(self) -> str:
        """마지막 이메일 ID를 반환.

        `LAST_TIME.json` 의 `{"last_email_id": "ID"}` 값을 읽습니다.
        파일이 없거나 파싱에 실패하면 빈 문자열을 반환합니다.
        """
        if self.last_time_file.exists():
            try:
                with self.last_time_file.open("r", encoding="utf-8") as fh:
                    data: dict[str, Any] = json.load(fh)
                return data.get("last_email_id", "")
            except Exception:
                pass
