import win32serviceutil
import win32service
import win32event
import servicemanager
import sys
import logging
import time
from pathlib import Path

class TestService(win32serviceutil.ServiceFramework):
    _svc_name_ = "TestFundService"
    _svc_display_name_ = "Test Fund Service"
    _svc_description_ = "최소 테스트 서비스"

    def __init__(self, args):
        win32serviceutil.ServiceFramework.__init__(self, args)
        self.hWaitStop = win32event.CreateEvent(None, 0, 0, None)
        
        # 로그 설정
        log_path = Path("C:/fund_mail/logs")
        log_path.mkdir(parents=True, exist_ok=True)
        
        logging.basicConfig(
            level=logging.DEBUG,
            format='%(asctime)s - %(levelname)s - %(message)s',
            handlers=[
                logging.FileHandler(log_path / "test_service.log"),
            ]
        )
        self.logger = logging.getLogger(__name__)

    def SvcStop(self):
        self.logger.info("=== 서비스 중지 요청 ===")
        self.ReportServiceStatus(win32service.SERVICE_STOP_PENDING)
        win32event.SetEvent(self.hWaitStop)
        self.logger.info("=== 서비스 중지 완료 ===")

    def SvcDoRun(self):
        self.logger.info("=== 서비스 시작 ===")
        
        try:
            # 즉시 실행 중 상태 보고
            self.ReportServiceStatus(win32service.SERVICE_RUNNING)
            self.logger.info("SERVICE_RUNNING 상태 보고 완료")
            
            servicemanager.LogMsg(
                servicemanager.EVENTLOG_INFORMATION_TYPE,
                servicemanager.PYS_SERVICE_STARTED,
                (self._svc_name_, '')
            )
            self.logger.info("이벤트 로그 기록 완료")
            
            # 간단한 작업 시뮬레이션
            counter = 0
            while True:
                # 중지 신호 확인 (1초마다)
                if win32event.WaitForSingleObject(self.hWaitStop, 1000) == win32event.WAIT_OBJECT_0:
                    break
                    
                counter += 1
                if counter % 60 == 0:  # 1분마다 로그
                    self.logger.info(f"서비스 실행 중... ({counter}초)")
                    
        except Exception as e:
            self.logger.error(f"서비스 실행 중 오류: {e}")
            servicemanager.LogErrorMsg(f"서비스 오류: {e}")
        
        self.logger.info("=== 서비스 종료 ===")

if __name__ == '__main__':
    if len(sys.argv) == 1:
        servicemanager.Initialize()
        servicemanager.PrepareToHostSingle(TestService)
        servicemanager.StartServiceCtrlDispatcher()
    else:
        win32serviceutil.HandleCommandLine(TestService)