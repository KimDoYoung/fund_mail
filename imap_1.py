import imaplib, base64, email, os
from msal import ConfidentialClientApplication
from dotenv import load_dotenv
load_dotenv()

# 환경 변수 또는 직접 설정
TENANT_ID     = os.getenv('TENANT_ID')
CLIENT_ID     = os.getenv('CLIENT_ID')
CLIENT_SECRET = os.getenv('CLIENT_SECRET')
MAIL_USER     = os.getenv('EMAIL_ID')

# 1. 토큰 발급 (client-credentials)
def get_token():
    authority = f'https://login.microsoftonline.com/{TENANT_ID}'
    app = ConfidentialClientApplication(
        CLIENT_ID, authority=authority, client_credential=CLIENT_SECRET)
    result = app.acquire_token_for_client(['https://outlook.office365.com/.default'])
    print("토큰 발급 결과:", result)
    return result.get('access_token')

# 2. SASL XOAUTH2 문자열 생성
def gen_auth_string(user, token):
    auth = f'user={user}\x01auth=Bearer {token}\x01\x01'
    return base64.b64encode(auth.encode()).decode()

# 3. 메일 가져오기
def fetch_emails():
    token = get_token()
    if not token:
        raise RuntimeError("토큰 발급 실패")

    imap = imaplib.IMAP4_SSL('outlook.office365.com', 993)
    imap.authenticate('XOAUTH2', lambda _: gen_auth_string(MAIL_USER, token))
    imap.select('INBOX')

    typ, data = imap.search(None, 'ALL')
    for num in data[0].split():
        typ, msg_data = imap.fetch(num, '(RFC822)')
        msg = email.message_from_bytes(msg_data[0][1])

        # 기본 정보 출력
        print("From:", msg.get('From'))
        print("Subject:", msg.get('Subject'))

        # 첨부파일 저장
        for part in msg.walk():
            if part.get_content_disposition() == 'attachment':
                fname = part.get_filename()
                payload = part.get_payload(decode=True)
                with open(fname, 'wb') as f:
                    f.write(payload)
                print("첨부파일 저장됨:", fname)

    imap.logout()

if __name__ == '__main__':
    fetch_emails()
