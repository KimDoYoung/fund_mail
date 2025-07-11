# fund_mail

## 목표
1. office365 메일을 사용하는데, 오늘 새로이 수신한 메일을 5분단위로 가져와서 DB에 저장한다.
2. office365 메일을 가져오기위해서 IMAP 프로토콜을 사용해야하는데. OAuth2로 인증을 받아야한다.
3. 5분단위 가져온 메일 내용으로 sqlite를 생성해서 넣는다.
4. 첨부파일을 다운로드한다.
5. 만들어진 sqlite db와 첨부파일을 sftp를 통해서 서버에 전송한다.

## 개요

3종류의 프로그램을 작성하다.

1. 과거 fund mail 백업받은 것으로 db와 첨부파일 추출, pst_utils라는 다른 프로젝트 생성
   - [pst_utils](https://github.com/KimDoYoung/pst-utils): fund mail을 백업받은 파일(확장자 pst)를 해석
   - 2018~2024년도 데이터까지는 이 프로그램으로 처리
2. one_day 용 : 지정한 하루의 메일데이터를 API를 이용해서 추출
3. 지정된시간(10분)용으로 윈도우 service를 이용해서 화면(cmd /c)없이 데몬형식으로 계속 도는 프로그램
   1. 원래는 exe로 만들어서  taskschd에 걸어서 사용하려고 했으나 auto_esafe와 충돌이 나서 service를 고려하게 됨


## 기술스택
1. python
2. microsoft office365
3. sqlitedb
4. sftp
5. dotenv
6. pywin32 

## 추가 요구사항

1. 테이블 필드를 pst-utils와 동일하게
2. 날짜 format을 2021-07-19 07:50:53.360111 식으로
3. 서비스로 돌려야 한다.
4. 첨부파일의 size추가

## 빌드

1. make_exe.sh
   1. fund_mail_once.exe를 만든다. 
   2. 배포pc에서 tashschd에 5분마다(또는 정해진 시간)마다 동작하도록 하면 된다. 
   3. .env와 만들어진 fund_mail_once.exe를 c:\fund_mail에 가져다 놓고.
   4. tashschd에서 매시간 동작하게 설정
   
2. make_svc.sh
   0. 관리자권한으로 cmd를 실행한다. main.py를 사용
   1. 전제조건 uv와 python 3.12이 설치되어 있다고 가정
   ```bash
      which python
      uv python list
      uv python install 3.12
   ``` 
   2. c:\fund_mail에 dist/의 모든 파일 copy
   3. uv venv
   4. source .venv/Scripts/activate (cmd에서는 .venv/Scripts/activate)
   5. uv sync (cmd : c:\Users\PC\.local\bin\uv sync)
   6. python src/service_wrapper.py install/start/stop/remove/debug
   7. 4개의 파일이 필요하다.
   8. 

```bash
#작업폴더에서
./make_svc.sh
cmd+r 관리자로
cd c:\fund_mail
c:\Users\PC\.local\bin\uv venv
.venv\Scripts\activate.bat
c:\Users\PC\.local\bin\uv sync
```

1. make_one_day.sh
   1. fund_mail_one_day.exe를 생성한다.
   ```bash
      fund_mail_one_day --date 2025-06-30
   ```
   2. 하루치를 모두 받음. 
   3. 인자로 지정한 날짜를 기준으로 폴더를 생성하고 upload 폴더를 만듬.
   4. run_one_day.sh은 bash shell로 from, to 2개의 인자로 날짜범위로 API를 이용해서 데이터를 받는다.


## 기능
1. microsoft의 office365 mail을 imap으로 가져온다.
2. sqlitedb에 보낸사람, 보낸시간, 제목, 내용을 db테이.블 fund_.mail에 넣는다. 
3. sqlitedb의 파일명은 fm_yyyy_mm_dd_HHMM.db로 한다.
4. 첨부파일은 지정된 폴더 .env에 기록된 attach_base_dir하위에 yyyy_mm_dd 밑에 넣는다. 
5. sqlitedb와 다운로드된 파일을 모두 지정된 server로 sftp를 통해서 upload한다.
6. 5분마다 위 동작이 이루어지며 lose된 메일은 절대로 없어야한다.

## 설계

### .env에 기술되어야할 사항들
```bash
#
# OFFICE365
#
EMAIL_ID=fund@123.
EMAIL_PW=123
TENANT_ID=abc
CLIENT_ID=def
CLIENT_SECRET=ghi
#
# Folder
#
BASE_DIR=c:\\fund_mail\\data

#
# SFTP
#
HOST=
PORT=
SFTP_ID=
SFTP_PW=
```

### from / sender

| 구분         | 언제 쓰나                                                             |
| ---------- | ----------------------------------------------------------------- |
| **from**   | 일반적으로 보이는 “보낸 사람”. 공유 사서함이나 위임 계정이면 mailbox 소유자 이름이 들어갈 수 있음      |
| **sender** | 실제로 SMTP 전송을 수행한 계정. “홍길동(대신 보냄)” 같은 경우 정확한 발신자 실명을 얻으려면 이 필드를 사용 |

> TIP from은 “보이는” 발신자, sender는 실제 SMTP 발송 계정입니다.
> 위임 / Send As / Send On Behalf 권한이 걸린 메일에서는 두 값이 달라질 수 있으니, 정확한 “누가 보냈는지”를 알고 싶다면 sender를 활용하세요.

### window service

window의 tashschd또는 cron을 이용한다.

### DB설계

테이블명: fund_mail
```text
CREATE TABLE fund_mail (
            id INTEGER PRIMARY KEY AUTOINCREMENT,  
            email_id TEXT ,  -- office365의 email_id
            subject TEXT,
            sender_address TEXT,
            sender_name TEXT,  
            from_address TEXT,  
            from_name TEXT,  
            to_recipients TEXT,  -- 수신자 목록
            cc_recipients TEXT,  -- 참조자 목록
            email_time TEXT,
            kst_time TEXT,
            content TEXT
        )
```
테이블명 : fund_mail_attach
```text
CREATE TABLE fund_mail_attach (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            parent_id INTEGER,
            email_id TEXT,  -- fund_mail 테이블의 id
            save_folder TEXT,
            file_name TEXT
        )
```

## 동작-Refactoring
1. LAST_TIME.json 에서 마지막 email_id,last_fetch_time를 읽어온다.
2. 만약 LAST_TIME.json이 존재하지 않는다면 가장 늦게 도착한 email 1개만 읽는다.
3. 그리고 LAST_TIME.json에 email_id, last_fetch_time을 저장하고 5분대기
4. 5분이 흘러서 1000개 를 시간역순으로 읽어서 저장해 두었던 email_id와 비교 만날때까지 읽는다.
5. 모두 db동작은 transaction처리한다
6. sftp로 올린다.
7. LAST_TIME에 최종 시각을 저장한다.
8. 5분 대기
9. 어떠한 이유로 실패하든지 LAST_TIME.json을 유지한다.
10. 실패하면 프로그램은 종료한다
   
## 배포
1. window pc에 배포한다
2. 설치문서 INSTALL.md 참조
   1. Window Service로 동작하게 한다.
   2. 기본 폴더 c:\fund_mail 로 한다

## office365 IMAP사용을 위한 OAuth 
1. [azure portal](https://portal.azure.com/#home)에 admin으로 login
   1. 이때 2차인증을 요구하는데, 연기하는 것 가능하다. 
   2. fund계정으로 login 가능하게 된다.
2. fund 계정으로 [azure portal](https://portal.azure.com/#home) 로그인
3. App등록 
   1. fund-imap, http://localhost/fund-imap 임의로 줘도 됨
   2. client_id (application_id), tenant_id  2개를 얻는다.
4. 관리->인증서등록
   1. client secret **값**을 얻는다.
5. 관리-> API 사용 권한
   1. Microsoft graph
      1. IMAP.AccessAsUser.All
      2. Mail.Read (관리자=admin에서 확인필요)
      3. offline_access
      4. User.Read
   2. Offcie 365 Exchange Online
      1. IMAP.AccessAsApp (관리자=admin에서 확인필요)
6. env.sample
   ```
        EMAIL_ID=fund@123.com
        EMAIL_PW=1234
        TENANT_ID=123
        CLIENT_ID=123
        CLIENT_SECRET=123
   ```
7. refresh 토큰은 사용자권한 기반 OAuth2 플로우
   1. 사용자가 로그인해서 인증하는 방식에서 사용됨
   2. fund_mail에서는 **Client Credentials Grant** 방식으로 refresh token은 사용치 않음.
8. [ms 공식 매뉴얼](https://learn.microsoft.com/en-us/graph/api/user-list-messages?view=graph-rest-1.0&tabs=http) 참조

9. 보낸 메일과 받은 메일이 포함되는가?
```text
/users/{id}/messages (또는 /me/messages) 호출은 “메일박스에 있는 모든 메시지” 를 가져오는 엔드포인트이기 때문에 Sent Items(보낸 편지함) 메일도 포함됩니다.

Microsoft 공식 Q&A에서 “/me/messages 엔드포인트는 모든 폴더·하위폴더의 메일을 반환한다” 고 명시 
learn.microsoft.com
.
즉 Inbox, Sent Items, Deleted Items, 사용자 정의 폴더 등 전체가 한 번에 모여 나옵니다.

Graph REST 문서도 “사용자의 mailbox 메시지를 가져온다(Deleted Items·Clutter 포함)”로 설명하며, mailbox 범위에는 Sent Items가 당연히 들어갑니다 
learn.microsoft.com
.
```   

## 유틸리티들

- src폴더에는 유틸리티성 py도 있음
### pst_extract.py

- 과거 백업받은 pst를 읽어서 sqlitedb에 넣는다.
- 기능 
   - 인자로 기술된 pst파일을 읽어서 target folder 하위에 pst파일명과 같은 sqlitedb를 생성하고 내용을 옮긴다.
   - 또한 target folder 하위에 attach폴더를 만들어서 첨부파일을 저장한다.
   - sqlitedb는 2개의 테이블이 있다 (db_actions.py를 사용)
   - pst파일은 37G 정도임.

### 기본 지식
- 리눅스에 pst-util 을 설치하면 readpst라는 실행파일이 생긴다. 이것으로 pst파일을 csv등으로 extract한후 db에 넣는방법도 있다.
- 윈도우에서는 [XstReader](https://github.com/Dijji/XstReader)라는 것도 있는데. 별로 신뢰감이 없다.
- 리눅스에서 libpff + 파이썬 바인딩을 이용해서 
```bash
# 1) libpff + 파이썬 바인딩
sudo apt update
sudo apt-get install -y libpff-dev python3-dev build-essential  # Debian/Ubuntu
sudo apt install python3-pypff libpff-dev pff-tools
uv add libpff-python

```
### chatgpt
```
가능 여부 — “된다, 단 스트리밍 방식으로!”
37 GB PST → SQLite 는 기술적으로 문제없습니다.

SQLite 는 수 TB 단위까지 지원하고, 단일 트랜잭션으로만 처리하지 않으면 메모리 사용량도 안정적입니다.

PST 는 트리 구조이므로 메시지 단위로 순회하며 바로 DB 에 flush 하면 8 ~ 16 GB RAM 환경에서도 충분히 돌릴 수 있습니다.

파이썬 쪽에서는 libpff(pypff) / libratom 가 대표 라이브러리입니다. 두 라이브러리 모두 내부적으로 동일한 C-코어(libpff)를 사용하므로 속도·메모리 특성은 비슷합니다.

pypff : 가장 가볍고 예제가 많음 
github.com

libratom : “pst → SQLite” 를 자동으로 해 주는 CLI( ratom scan --write-bodies --db … )가 있지만, 기존 스키마(fund_mail 등)과 다르므로 여기서는 pypff로 직접 파싱하는 스크립트를 제안합니다. 
github.com

libpst 의 readpst(pst-utils) 를 먼저 mbox 로 변환한 뒤 mbox 를 다시 파싱하는 방법도 있으나, 한 번 더 디스크 I/O 가 발생하고 첨부파일 메타데이터가 일부 유실되므로 권장하지 않습니다.
```
## 참고

1. [마이크로소프트 인증사이트](https://myaccount.microsoft.com/)
2. [Learn graph api](https://learn.microsoft.com/en-us/graph/api/user-list-messages?view=graph-rest-1.0&tabs=http)




## python window service 만들기

## 작업

1. python(ver 3.12), pywind32를 이용해서 윈도우 서비스형태로 프로그램이 동작하게 하고자 함
2. 다음과 같은 에러가 남 install은 성공함. 관리자권한으로 cmd.exe에서 수행해 봄
```
(svc_example) C:\Users\PC\Work\svc_example>python svc1.py start

Starting service MyService

Error starting service: 서비스가 시작이나 제어 요청에 빠르게 응답하지 않았습니다.
```

3. .venv에 4개의 파일을 둠. 
	- pythonservice.exe
	- python312.dll
	- pywintypes312.dll
	- pythoncom312.dll

4. svc1.py 코드
   ```python
   import win32serviceutil
   import servicemanager
   import ctypes
   import sys
   import time

   OutputDebugString = ctypes.windll.kernel32.OutputDebugStringW

   class MyServiceFramework(win32serviceutil.ServiceFramework):
   _svc_name_ = 'MyPythonService'
   _svc_display_name_ = 'My Python Service'
   is_running = False

   def SvcStop(self):
      OutputDebugString("MyServiceFramework __SvcStop__")
      self.is_running = False

   def SvcDoRun(self):
      self.is_running = True
      while self.is_running:
         OutputDebugString("MyServiceFramework __loop__")
         time.sleep(1)

   if '__main__' == __name__:
   win32serviceutil.HandleCommandLine(MyServiceFramework)
   ```	
## 주의

- 고급시스템에서 PATH에 등록해 줘야한다.


