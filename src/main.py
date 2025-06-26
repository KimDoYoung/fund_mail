import threading
import time
import sys
import shutil
from datetime import datetime
from config import Config # noqa: E402
from fetch_email import fetch_email_from_office365
from sftp_upload import upload_to_sftp  # noqa: E402
from logger import logger


class TaskScheduler:
    def __init__(self, interval=300):
        self.interval = interval
        self._running = threading.Event()
        self._running.set()           # 실행 상태
        self._timer = None            # 마지막 Timer 레퍼런스
        self.config = Config.load()  # 환경 변수 로드

    def _run_task(self):
        success = False                    # 실행 결과 플래그
        try:
            logger.info("=" * 59)
            logger.info("⏺️ fund메일 수집이 시작됩니다.   작업 시작: %s", datetime.now())
            logger.info("=" * 59)
            # 
            # self.config.last_time_file의 백업을 만든다.
            backup_path = self.config.last_time_file + ".bak"
            shutil.copy2(self.config.last_time_file, backup_path)
            db_path = fetch_email_from_office365(self.config)
            if db_path:                    # db_path가 None → 첫 실행(수집만, SFTP 생략)
                upload_to_sftp(self.config, db_path)

            logger.info("=" * 59)
            logger.info("⏺️ fund메일 작업이 완료되었습니다. 완료 시각: %s", datetime.now())
            logger.info("=" * 59)
            success = True                 # 여기까지 오면 정상
        except Exception:
            logger.info("=" * 59)
            logger.exception("⛔ fund메일 작업 중 예외 발생 - 프로그램을 종료합니다.")
            logger.info("=" * 59)
            self.stop()                    # 타이머 취소 및 플래그 클리어
            #백업을 복구
            if shutil.os.path.exists(backup_path):
                shutil.copy2(backup_path, self.config.last_time_file)
                logger.warning("⚠️백업된 LAST_TIME.json을 복구했습니다: %s", backup_path)
            else:
                logger.error("❌ 백업 파일이 존재하지 않습니다: %s", backup_path)
            raise                          # 메인 루프까지 예외 전파
        finally:
            # 정상 종료 + 스케줄러가 살아있을 때만 다음 실행 예약
            if success and self._running.is_set():
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

def fetch_fund_mail():
    scheduler = TaskScheduler()
    try:
        scheduler.start()      # start() 내부에서 _run_task() 처음 실행
        while True:
            time.sleep(1)      # Ctrl-C 처리용 루프
    except KeyboardInterrupt:
        scheduler.stop()
        logger.info("⏹️  사용자가 스케줄러를 중단했습니다.")
    except Exception:
        # _run_task 에서 올라온 예외 – 이미 로그 찍었으므로 종료만
        scheduler.stop()
        sys.exit(1)

# 사용 예시
if __name__ == "__main__":
    fetch_fund_mail()
    # scheduler = TaskScheduler()
    # scheduler.start()
    # try:
    #     while True:
    #         time.sleep(1)
    # except KeyboardInterrupt:
    #     scheduler.stop()
    #     logger.info("스케줄러가 중단되었습니다.")