import win32serviceutil
import win32service
import win32event
import servicemanager
import sys
import os
import logging
import threading
import time
from pathlib import Path

# 실제 작업 스크립트 import
try:
    from main import TaskScheduler
except ImportError as e:
    print(f"main.py 또는 TaskScheduler 클래스를 찾을 수 없습니다: {e}")
    print("main.py 파일과 TaskScheduler 클래스가 올바르게 구현되어 있는지 확인하세요.")
    sys.exit(1)

class FundEmailFetchService(win32serviceutil.ServiceFramework):
    _svc_name_ = "FundEmailFetchService"
    _svc_display_name_ = "Fund Email Fetch Service"
    _svc_description_ = "5분마다 이메일을 수집하여 DB에 저장하고 SFTP에 업로드하는 서비스"

    def __init__(self, args):
        win32serviceutil.ServiceFramework.__init__(self, args)
        self.hWaitStop = win32event.CreateEvent(None, 0, 0, None)
        self.is_alive = True
        self.scheduler = None
        self.scheduler_thread = None

        # 서비스 로그 설정
        self.setup_logging()

    def setup_logging(self):
        """서비스용 로깅 설정"""
        log_path = Path("C:/fund_mail/logs")
        log_path.mkdir(parents=True, exist_ok=True)
        
        logging.basicConfig(
            level=logging.INFO,
            format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
            handlers=[
                logging.FileHandler(log_path / "service.log"),
                logging.StreamHandler()
            ]
        )
        self.logger = logging.getLogger(__name__)

    def SvcStop(self):
        """서비스 중지"""
        self.logger.info("서비스 중지 요청 받음")
        self.ReportServiceStatus(win32service.SERVICE_STOP_PENDING)
        
        # 스케줄러 중지
        if self.scheduler:
            self.scheduler.stop()
            
        # 스케줄러 스레드 종료 대기
        if self.scheduler_thread and self.scheduler_thread.is_alive():
            self.scheduler_thread.join(timeout=10)
            
        self.is_alive = False
        win32event.SetEvent(self.hWaitStop)

    def SvcDoRun(self):
        """서비스 실행"""
        try:
            # 즉시 서비스 상태를 "실행 중"으로 보고
            self.ReportServiceStatus(win32service.SERVICE_RUNNING)
            
            self.logger.info("이메일 수집 서비스 시작")
            servicemanager.LogMsg(
                servicemanager.EVENTLOG_INFORMATION_TYPE,
                servicemanager.PYS_SERVICE_STARTED,
                (self._svc_name_, '')
            )

            # TaskScheduler를 별도 스레드에서 실행
            self.scheduler = TaskScheduler(interval=300)  # 5분
            self.scheduler_thread = threading.Thread(
                target=self._run_scheduler, 
                daemon=True
            )
            self.scheduler_thread.start()
            
            self.logger.info("스케줄러 스레드 시작 완료")

            # 서비스가 중지될 때까지 대기
            win32event.WaitForSingleObject(self.hWaitStop, win32event.INFINITE)
            
        except Exception as e:
            self.logger.error(f"서비스 실행 중 오류: {e}")
            servicemanager.LogErrorMsg(f"서비스 오류: {e}")
            # 오류 발생 시 서비스 상태를 "중지됨"으로 변경
            self.ReportServiceStatus(win32service.SERVICE_STOPPED)

    def _run_scheduler(self):
        """스케줄러를 실행하는 내부 메서드"""
        try:
            self.logger.info("TaskScheduler 초기화 시작")
            
            # 스케줄러 시작 전에 짧은 지연 (초기화 안정성)
            time.sleep(2)
            
            self.scheduler.start()
            self.logger.info("TaskScheduler 시작 완료")
            
            # 서비스가 중지될 때까지 대기
            while self.is_alive:
                time.sleep(1)
                
        except Exception as e:
            self.logger.error(f"스케줄러 실행 중 오류: {e}")
            # 스케줄러 오류 시 서비스도 중지
            self.SvcStop()
        finally:
            if self.scheduler:
                self.logger.info("TaskScheduler 중지 시작")
                self.scheduler.stop()
                self.logger.info("TaskScheduler 중지 완료")

def run_debug():
    """디버그 모드로 실행 (개발용)"""
    print("디버그 모드로 실행 중...")
    
    # 로깅 설정
    log_path = Path("logs")
    log_path.mkdir(exist_ok=True)
    
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
        handlers=[
            logging.FileHandler(log_path / "debug.log"),
            logging.StreamHandler()
        ]
    )
    logger = logging.getLogger(__name__)
    
    try:
        scheduler = TaskScheduler(interval=300)  # 5분
        scheduler.start()
        
        print("서비스가 실행 중입니다. Ctrl+C로 중지하세요.")
        logger.info("디버그 모드로 스케줄러 시작")
        
        # 무한 대기 (main.py의 fetch_fund_mail() 함수와 동일한 방식)
        while True:
            time.sleep(1)
            
    except KeyboardInterrupt:
        print("\n서비스를 중지합니다...")
        logger.info("사용자가 디버그 모드 스케줄러를 중단했습니다.")
        scheduler.stop()
        print("서비스가 중지되었습니다.")
    except Exception as e:
        logger.error(f"디버그 모드 실행 중 오류 발생: {e}")
        if 'scheduler' in locals():
            scheduler.stop()

if __name__ == '__main__':
    if len(sys.argv) == 1:
        # 인수가 없으면 서비스 모드로 실행
        servicemanager.Initialize()
        servicemanager.PrepareToHostSingle(FundEmailFetchService)
        servicemanager.StartServiceCtrlDispatcher()
    elif len(sys.argv) == 2 and sys.argv[1] == 'debug':
        # debug 인수가 있으면 디버그 모드로 실행
        run_debug()
    else:
        # 다른 인수들은 서비스 관리 명령어
        win32serviceutil.HandleCommandLine(FundEmailFetchService)