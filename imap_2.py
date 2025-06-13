import imaplib, base64, email, os
from msal import ConfidentialClientApplication
from dotenv import load_dotenv
load_dotenv()

# 환경 변수 또는 직접 설정
TENANT_ID     = os.getenv('TENANT_ID')
CLIENT_ID     = os.getenv('CLIENT_ID')
CLIENT_SECRET = os.getenv('CLIENT_SECRET')
MAIL_USER     = os.getenv('EMAIL_ID')

print("TENANT_ID:", TENANT_ID)
print("CLIENT_ID:", CLIENT_ID) 
print("CLIENT_SECRET:", CLIENT_SECRET[:10] + "..." if CLIENT_SECRET else None)  # 보안상 일부만 출력
print("MAIL_USER:", MAIL_USER)

# 1. 토큰 발급 (올바른 스코프로 수정)
def get_token():
    authority = f'https://login.microsoftonline.com/{TENANT_ID}'
    app = ConfidentialClientApplication(
        CLIENT_ID, authority=authority, client_credential=CLIENT_SECRET)
    
    # IMAP 접근을 위한 올바른 스코프
    # scopes = ['https://outlook.office365.com/IMAP.AccessAsUser.All']
    scopes = ['https://outlook.office365.com/.default']
    
    result = app.acquire_token_for_client(scopes)
    print("토큰 발급 결과:", result)
    
    if 'access_token' in result:
        print("✅ 토큰 발급 성공")
        return result.get('access_token')
    else:
        print("❌ 토큰 발급 실패")
        print("오류:", result.get('error_description'))
        return None

# 2. SASL XOAUTH2 문자열 생성 (개선된 버전)
def gen_auth_string(user, token):
    auth = f'user={user}\x01auth=Bearer {token}\x01\x01'
    auth_bytes = base64.b64encode(auth.encode('utf-8')).decode('ascii')
    print(f"생성된 인증 문자열 길이: {len(auth_bytes)}")
    return auth_bytes

# 3. 연결 테스트 함수 추가
def test_connection():
    """IMAP 연결만 테스트"""
    token = get_token()
    if not token:
        print("❌ 토큰 발급 실패로 연결 테스트 중단")
        return False
    
    try:
        print("📡 IMAP 서버 연결 시도...")
        imap = imaplib.IMAP4_SSL('outlook.office365.com', 993)
        print("✅ IMAP 서버 연결 성공")
        
        print("🔐 OAuth2 인증 시도...")
        auth_string = gen_auth_string(MAIL_USER, token)
        
        # 디버그 정보
        print(f"사용자: {MAIL_USER}")
        print(f"토큰 길이: {len(token)}")
        
        # 인증 시도
        imap.authenticate('XOAUTH2', lambda _: auth_string)
        print("✅ OAuth2 인증 성공")
        
        # 사서함 선택
        imap.select('INBOX')
        print("✅ INBOX 선택 성공")
        
        # 이메일 개수 확인
        typ, data = imap.search(None, 'ALL')
        email_count = len(data[0].split()) if data[0] else 0
        print(f"📧 총 이메일 개수: {email_count}")
        
        imap.logout()
        return True
        
    except imaplib.IMAP4.error as e:
        print(f"❌ IMAP 오류: {e}")
        return False
    except Exception as e:
        print(f"❌ 일반 오류: {e}")
        return False

# 4. 메일 가져오기 (안전한 버전)
def fetch_emails(limit=5):
    """이메일 가져오기 (제한된 개수)"""
    if not test_connection():
        print("연결 테스트 실패로 이메일 가져오기 중단")
        return
    
    token = get_token()
    if not token:
        raise RuntimeError("토큰 발급 실패")

    try:
        imap = imaplib.IMAP4_SSL('outlook.office365.com', 993)
        imap.authenticate('XOAUTH2', lambda _: gen_auth_string(MAIL_USER, token))
        imap.select('INBOX')

        typ, data = imap.search(None, 'ALL')
        email_ids = data[0].split()
        
        # 최근 이메일만 처리 (역순으로 제한)
        recent_emails = email_ids[-limit:] if len(email_ids) > limit else email_ids
        recent_emails.reverse()  # 최신순으로
        
        print(f"📧 최근 {len(recent_emails)}개 이메일 처리 중...")
        
        for i, num in enumerate(recent_emails, 1):
            print(f"\n--- 이메일 {i}/{len(recent_emails)} ---")
            
            typ, msg_data = imap.fetch(num, '(RFC822)')
            msg = email.message_from_bytes(msg_data[0][1])

            # 기본 정보 출력
            print("From:", msg.get('From'))
            print("Subject:", msg.get('Subject'))
            print("Date:", msg.get('Date'))

            # 첨부파일 저장
            attachment_count = 0
            for part in msg.walk():
                if part.get_content_disposition() == 'attachment':
                    fname = part.get_filename()
                    if fname:
                        try:
                            payload = part.get_payload(decode=True)
                            if payload:
                                # 안전한 파일명 생성
                                safe_fname = f"{i}_{fname}"
                                with open(safe_fname, 'wb') as f:
                                    f.write(payload)
                                print(f"첨부파일 저장됨: {safe_fname} ({len(payload)} bytes)")
                                attachment_count += 1
                        except Exception as e:
                            print(f"첨부파일 저장 실패 ({fname}): {e}")
            
            if attachment_count == 0:
                print("첨부파일 없음")

        imap.logout()
        print(f"\n✅ 총 {len(recent_emails)}개 이메일 처리 완료")
        
    except Exception as e:
        print(f"❌ 이메일 가져오기 오류: {e}")

# 5. 단계별 실행
if __name__ == '__main__':
    print("=== 이메일 가져오기 시작 ===")
    
    # 1단계: 연결 테스트
    print("\n1단계: 연결 테스트")
    if test_connection():
        print("✅ 연결 테스트 성공")
        
        # 2단계: 이메일 가져오기
        print("\n2단계: 이메일 가져오기")
        fetch_emails(limit=3)  # 처음에는 3개만
    else:
        print("❌ 연결 테스트 실패")
        print("\n🔍 문제 해결 방법:")
        print("1. MAIL_USER가 정확한 이메일 주소인지 확인")
        print("2. Azure에서 IMAP.AccessAsUser.All 권한 확인")
        print("3. 관리자 동의 필요할 수 있음")
        print("4. 이메일 계정이 Exchange Online 사용하는지 확인")