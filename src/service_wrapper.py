import win32serviceutil
import win32service
import win32event
import servicemanager
import sys
import os
import logging
from pathlib import Path

# 실제 작업 스크립트 import
from main  import TaskScheduler

class FundEmailFetchService(win32serviceutil.ServiceFramework):
    _svc_name_ = "EmailFetchService"
    _svc_display_name_ = "Fund Email Fetch Service"
    _svc_description_ = "5분마다 이메일을 수집하여 DB에 저장하고 SFTP에 업로드하는 서비스"

    def __init__(self, args):
        win32serviceutil.ServiceFramework.__init__(self, args)
        self.hWaitStop = win32event.CreateEvent(None, 0, 0, None)
        self.is_alive = True
        self.scheduler = None

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
        win32event.SetEvent(self.hWaitStop)
        self.is_alive = False
        
        if self.scheduler:
            self.scheduler.stop()

    def SvcDoRun(self):
        """서비스 실행"""
        try:
            self.logger.info("이메일 수집 서비스 시작")
            servicemanager.LogMsg(
                servicemanager.EVENTLOG_INFORMATION_TYPE,
                servicemanager.PYS_SERVICE_STARTED,
                (self._svc_name_, '')
            )

            # TaskScheduler 시작
            self.scheduler = TaskScheduler(interval=300)  # 5분
            self.scheduler.start()

            # 서비스가 중지될 때까지 대기
            win32event.WaitForSingleObject(self.hWaitStop, win32event.INFINITE)
            
        except Exception as e:
            self.logger.error(f"서비스 실행 중 오류: {e}")
            servicemanager.LogErrorMsg(f"서비스 오류: {e}")

if __name__ == '__main__':
    if len(sys.argv) == 1:
        servicemanager.Initialize()
        servicemanager.PrepareToHostSingle(FundEmailFetchService)
        servicemanager.StartServiceCtrlDispatcher()
    else:
        win32serviceutil.HandleCommandLine(FundEmailFetchService)