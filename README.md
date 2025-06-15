# fund_mail

## 목표
1. office365 메일을 사용하는데, 오늘 새로이 수신한 메일을 5분단위로 가져와서 DB에 저장한다.
2. office365 메일을 가져오기위해서 IMAP 프로토콜을 사용해야하는데. OAuth2로 인증을 받아야한다.


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