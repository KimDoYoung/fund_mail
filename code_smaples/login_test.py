# from imapclient import IMAPClient
import os
from dotenv import load_dotenv
from imapclient import IMAPClient

load_dotenv()

EMAIL_CONFIG = {
    "host": "outlook.office365.com",
    "port": 993,
    "email": os.getenv("EMAIL_ID"),
    "password": os.getenv("EMAIL_PW"),
}

# EMAIL_CONFIG = {
#     "host": "outlook.office365.com",  # 반드시 IMAP 서버
#     "port": 993,
#     "email": "abc@k-fs.co.kr",
#     "password": "1234",  # 앱 비밀번호 필요할 수 있음
# }

try:
    with IMAPClient(EMAIL_CONFIG['host'], port=EMAIL_CONFIG['port'], ssl=True) as client:
        client.login(EMAIL_CONFIG['email'], EMAIL_CONFIG['password'])
        client.select_folder('INBOX')
        print("✅ 로그인 성공 및 INBOX 접속 완료")
except Exception as e:
    print("❌ 로그인 실패:", e)
