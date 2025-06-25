# Windows Service 설치 가이드

## 1. 사전 준비

### Python 설치
```cmd
# Python 3.8+ 설치 (Microsoft Store 또는 python.org에서)
python --version
```

### 프로젝트 설정
```cmd
# 프로젝트 폴더로 이동
cd C:\fund_mail

# 가상환경 생성
python -m venv venv

# 가상환경 활성화
venv\Scripts\activate

# 의존성 설치
pip install -r requirements.txt
```

## 2. 서비스 설치

### 방법 1: 배치 스크립트 사용 (추천)
```cmd
# 관리자 권한으로 실행
manage_service.bat
# 메뉴에서 1번 선택하여 설치
```

### 방법 2: 수동 설치
```cmd
# 관리자 권한으로 실행
venv\Scripts\activate
python service_wrapper.py install
```

## 3. 서비스 시작
```cmd
net start EmailFetchService
```

## 4. 서비스 상태 확인
```cmd
sc query EmailFetchService
```

## 5. 로그 확인
- 위치: `C:\ProgramData\EmailFetchService\logs\service.log`
- Windows 이벤트 뷰어에서도 확인 가능

## 6. 서비스 제거
```cmd
net stop EmailFetchService
python service_wrapper.py remove
```

## 트러블슈팅

### 권한 문제
- 관리자 권한으로 명령 프롬프트 실행
- 서비스 계정에 필요한 폴더 권한 부여

### 경로 문제
- 절대 경로 사용
- config 파일 경로 확인

### 의존성 문제
- 가상환경이 제대로 활성화되었는지 확인
- pip list로 패키지 설치 확인