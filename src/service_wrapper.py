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
from dotenv import load_dotenv

class FundEmailFetchService(win32serviceutil.ServiceFramework):
    _svc_name_ = "FundEmailFetchService"
    _svc_display_name_ = "Fund Email Fetch Service"
    _svc_description_ = "정해진시간(5분)마다 fund 메일을 수집하여 DB에 저장하고 SFTP에 업로드하는 서비스"

    def __init__(self, args):
        win32serviceutil.ServiceFramework.__init__(self, args)
        self.hWaitStop = win32event.CreateEvent(None, 0, 0, None)
        self.is_alive = True
        self.scheduler = None
        self.scheduler_thread = None

        self.base_path = r'C:\fund_mail'
        self.venv_path = os.path.join(self.base_path, '.venv')
        # 서비스 실행 경로 설정
        self.setup_paths()

        # 서비스 로그 설정
        self.setup_logging()

    def setup_paths(self):
        """서비스 실행 시 경로 설정"""
        # 작업 디렉터리 설정
        os.chdir(self.base_path)
        
        # Python 경로 설정
        venv_scripts = os.path.join(self.venv_path, 'Scripts')
        venv_lib = os.path.join(self.venv_path, 'Lib', 'site-packages')
        
        # PATH 환경 변수 업데이트
        current_path = os.environ.get('PATH', '')
        os.environ['PATH'] = f"{venv_scripts};{current_path}"
        
        # Python 경로 추가
        sys.path.insert(0, venv_lib)
        sys.path.insert(0, os.path.join(self.base_path, 'src'))

    def setup_logging(self):
        """서비스용 로깅 설정"""
        log_path = Path("C:/fund_mail/logs")
        log_path.mkdir(parents=True, exist_ok=True)
        
        logging.basicConfig(
            level=logging.INFO,
            format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
            handlers=[
                logging.FileHandler(log_path / "service.log", encoding='utf-8'),
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
            try:
                self.scheduler.stop()
                self.logger.info("스케줄러 중지 완료")
            except Exception as e:
                self.logger.error(f"스케줄러 중지 중 오류: {e}")
            
        # 스케줄러 스레드 종료 대기
        if self.scheduler_thread and self.scheduler_thread.is_alive():
            self.scheduler_thread.join(timeout=10)
            if self.scheduler_thread.is_alive():
                self.logger.warning("스케줄러 스레드가 정상적으로 종료되지 않았습니다.")
            
        self.is_alive = False
        win32event.SetEvent(self.hWaitStop)

    def SvcDoRun(self):
        """서비스 실행"""
        try:
            self.logger.info("서비스 시작 중...")
            
            # 서비스 시작 로그
            servicemanager.LogMsg(
                servicemanager.EVENTLOG_INFORMATION_TYPE,
                servicemanager.PYS_SERVICE_STARTED,
                (self._svc_name_, '')
            )

            # ★ 먼저 서비스 상태를 "실행 중"으로 보고 (중요!)
            self.ReportServiceStatus(win32service.SERVICE_RUNNING)
            self.logger.info("서비스 상태를 RUNNING으로 보고함")

            # ★ 그 다음에 스케줄러를 별도 스레드에서 초기화
            self.scheduler_thread = threading.Thread(
                target=self._run_scheduler, 
                daemon=True
            )
            self.scheduler_thread.start()
            
            self.logger.info("스케줄러 스레드 시작 완료")

            # 서비스가 중지될 때까지 대기
            win32event.WaitForSingleObject(self.hWaitStop, win32event.INFINITE)
            
            self.logger.info("서비스 정상 종료")
            
        except Exception as e:
            error_msg = f"서비스 실행 중 오류: {e}"
            self.logger.error(error_msg, exc_info=True)
            servicemanager.LogErrorMsg(error_msg)
            # 오류 발생 시 서비스 상태를 "중지됨"으로 변경
            self.ReportServiceStatus(win32service.SERVICE_STOPPED)


    def _run_scheduler(self):
        """스케줄러를 실행하는 내부 메서드"""
        try:
            self.logger.info("TaskScheduler 초기화 시작")
            
            # ★ 짧은 지연 후 import 시도
            time.sleep(2)
            
            # ★ 여기서 import 시도하고 오류 발생 시 로그 기록
            try:
                from main import TaskScheduler
                self.logger.info("TaskScheduler 모듈 import 성공")
            except ImportError as e:
                error_msg = f"TaskScheduler import 실패: {e}"
                self.logger.error(error_msg, exc_info=True)
                # import 실패 시 경로 정보 로그
                self.logger.error(f"현재 작업 디렉토리: {os.getcwd()}")
                self.logger.error(f"Python 경로: {sys.path}")
                self.logger.error(f"main.py 파일 존재 여부: {Path('main.py').exists()}")
                return
            except Exception as e:
                error_msg = f"TaskScheduler import 중 예상치 못한 오류: {e}"
                self.logger.error(error_msg, exc_info=True)
                return
            
            # ★ 스케줄러 초기화
            self.scheduler = TaskScheduler(interval=300)  # 5분
            self.logger.info("TaskScheduler 객체 생성 완료")
            
            # ★ 스케줄러 시작
            self.scheduler.start()
            self.logger.info("TaskScheduler 시작 완료")
            
            # 서비스가 중지될 때까지 대기
            while self.is_alive:
                time.sleep(1)
                
        except Exception as e:
            self.logger.error(f"스케줄러 실행 중 오류: {e}", exc_info=True)
            # 스케줄러 오류 시 서비스는 계속 실행 (서비스 자체를 중지하지 않음)
        finally:
            if self.scheduler:
                try:
                    self.scheduler.stop()
                    self.logger.info("TaskScheduler 중지 완료")
                except Exception as e:
                    self.logger.error(f"TaskScheduler 중지 중 오류: {e}")

def run_debug():
    """디버그 모드로 실행 (개발용)"""
    print("디버그 모드로 실행 중...")
    
    # 경로 설정
    base_dir = Path(__file__).parent.absolute()
    os.chdir(base_dir)
    if str(base_dir) not in sys.path:
        sys.path.insert(0, str(base_dir))
    
    # 로깅 설정
    log_path = base_dir / "logs"
    log_path.mkdir(exist_ok=True)
    
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
        handlers=[
            logging.FileHandler(log_path / "debug.log", encoding='utf-8'),
            logging.StreamHandler()
        ]
    )
    logger = logging.getLogger(__name__)
    
    try:
        # TaskScheduler import
        from main import TaskScheduler
        
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
        logger.error(f"디버그 모드 실행 중 오류 발생: {e}", exc_info=True)
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
