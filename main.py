import threading
import time
from datetime import datetime
from config import Config # noqa: E402
from fetch_email import fetch_email_from_office365, upload_to_sftp
from logger import logger


class TaskScheduler:
    def __init__(self, interval=300):
        self.interval = interval
        self._running = threading.Event()
        self._running.set()           # 실행 상태
        self._timer = None            # 마지막 Timer 레퍼런스
        self.config = Config.load()  # 환경 변수 로드

    def _run_task(self):
        try:
            logger.info(f"작업이 시작됩니다. 작업 시작: {datetime.now()}")
            # 실제 작업 로직
            db_path = fetch_email_from_office365(self.config)  # 이메일 가져오기
            upload_to_sftp(self.config, db_path)  # SFTP 업로드
            logger.info("작업이 완료되었습니다. 완료 시각: %s", datetime.now())
        except Exception as e:
            logger.error(f"[ERROR] {e}")
        finally:
            # 다음 실행 예약
            if self._running.is_set():
                self._timer = threading.Timer(self.interval, self._run_task)
                self._timer.start()

    def start(self):
        if not self._running.is_set():
            self._running.set()
        self._run_task()               # 첫 실행

    def stop(self):
        self._running.clear()          # 중단 플래그
        if self._timer is not None:
            self._timer.cancel()       # 예약된 타이머 취소



# 사용 예시
if __name__ == "__main__":
    scheduler = TaskScheduler()
    scheduler.start()
    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        scheduler.stop()
        logger.info("스케줄러가 중단되었습니다.")