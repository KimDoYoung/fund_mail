# fund_mail

## 목표
1. office365 메일을 사용하는데, 오늘 새로이 수신한 메일을 5분단위로 가져와서 DB에 저장한다.
2. office365 메일을 가져오기위해서 IMAP 프로토콜을 사용해야하는데. OAuth2로 인증을 받아야한다.


## 기술스택
1. python
2. microsoft office365
3. sqlitedb
4. sftp
5. dotenv

## 기능
1. microsoft의 office365 mail을 imap으로 가져온다.
2. sqlitedb에 보낸사람, 보낸시간, 제목, 내용을 db테이.블 fund_.mail에 넣는다. 
3. sqlitedb의 파일명은 fm_yyyy_mm_dd_HHMM.db로 한다.
4. 첨부파일은 지정된 폴더 .env에 기록된 attach_base_dir하위에 yyyy_mm_dd 밑에 넣는다. 
5. sqlitedb와 다운로드된 파일을 모두 지정된 server로 sftp를 통해서 upload한다.
중요기능
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

### 스케줄
window의 tashschd또는 cron을 이용한다.

### DB설계

테이블명: fund_mail
```text
id : email_id (office365의 email_id)
subject: email_title
sender: email 보낸사람
email_time: 받은 시각
content : 내용
```
테이블명 : fund_mail_attach
```text
id : auto_increment
parent_id : fund_main의 id
save_folder: 저장폴더명
file_name: 파일명
```

## 동작

1. LAST_TIME.txt에 time을 저장
2. fund_mail 수행시 LAST_TIME.txt 파일이 없으면 그 날의 00시,00분부터 메일을 가져옴
3. 가져온 메일로 table fund_mail과 fund_mail_attach를 채우고
4. sftp로 서버로 sql db와 첨부파일을 모두 sftp로 보냄
5. LAST_TIME.txt에 마지막 메일의 time을 저장폴더명
6. 5분후 last_time.txt의 시간을 읽어서 그 시각 이후의 메일을 가져온 후 3번부터 반복

## 동작-Refactoring
1. LAST_TIME.json 에서 마지막 email_id를 읽어온다.
2. 만약 LAST_TIME.json이 존재하지 않는다면 가장 늦게 도착한 email 1개만 읽는다.
3. 그리고 LAST_TIME.json에 email_id를 저장하고 5분대기
4. 5분이 흘러서 1000개 를 시간역순으로 읽어서 저장해 두었던 email_id와 비교 만날때까지 읽는다.
5. 모두 db동작은 transaction처리한다
6. sftp로 올린다.
7. LAST_TIME에 최종 시각을 저장한다.
8. 5분 대기
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
   

## 참고

1. [마이크로소프트 인증사이트](https://myaccount.microsoft.com/)
2. [Learn graph api](https://learn.microsoft.com/en-us/graph/api/user-list-messages?view=graph-rest-1.0&tabs=http)