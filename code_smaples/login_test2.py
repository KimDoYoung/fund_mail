import imaplib
import os
import ssl
import socket
from dotenv import load_dotenv
load_dotenv()
def test_email_connection():
    """이메일 서버 연결 테스트"""
    
    # 회사 메일 설정 (받으신 정보 기반)
    EMAIL_CONFIG = {
        'email': os.getenv('EMAIL_ID'),
        'password': os.getenv('EMAIL_PW'),
        # IMAP 서버 정보 (여러 가능성 테스트)
        'possible_servers': [
            {'host': 'outlook.office365.com', 'port': 993, 'ssl': True},
            {'host': 'imap-mail.outlook.com', 'port': 993, 'ssl': True}, 
            {'host': 'imap.k-fs.co.kr', 'port': 993, 'ssl': True},
            {'host': 'mail.k-fs.co.kr', 'port': 993, 'ssl': True},
            {'host': 'outlook.office365.com', 'port': 143, 'ssl': False},
            {'host': 'imap-mail.outlook.com', 'port': 143, 'ssl': False}
        ]
    }
    
    print("🔍 이메일 서버 연결 테스트 시작...")
    print(f"📧 테스트 계정: {EMAIL_CONFIG['email']}")
    print("=" * 60)
    
    for i, server in enumerate(EMAIL_CONFIG['possible_servers'], 1):
        print(f"\n{i}. {server['host']}:{server['port']} ({'SSL' if server['ssl'] else 'Plain'})")
        print("-" * 40)
        
        try:
            # SSL/비SSL 연결 시도
            if server['ssl']:
                # SSL 연결
                mail = imaplib.IMAP4_SSL(server['host'], server['port'])
                print("✅ SSL 연결 성공")
            else:
                # 일반 연결 후 STARTTLS
                mail = imaplib.IMAP4(server['host'], server['port'])
                print("✅ 일반 연결 성공")
                try:
                    mail.starttls()
                    print("✅ STARTTLS 업그레이드 성공")
                except Exception as e:
                    print(f"⚠️  STARTTLS 실패: {e}")
            
            # 로그인 시도
            try:
                mail.login(EMAIL_CONFIG['email'], EMAIL_CONFIG['password'])
                print("✅ 로그인 성공!")
                
                # 폴더 목록 조회
                try:
                    folders = mail.list()
                    print("✅ 폴더 목록 조회 성공")
                    
                    # INBOX 선택 시도
                    mail.select('INBOX')
                    print("✅ INBOX 선택 성공")
                    
                    # 메일 개수 확인
                    status, messages = mail.search(None, 'ALL')
                    if status == 'OK':
                        mail_count = len(messages[0].split()) if messages[0] else 0
                        print(f"📬 INBOX 메일 개수: {mail_count}")
                    
                    print(f"🎉 완전 성공! 이 설정을 사용하세요:")
                    print(f"   Host: {server['host']}")
                    print(f"   Port: {server['port']}")
                    print(f"   SSL: {server['ssl']}")
                    
                    mail.close()
                    mail.logout()
                    return server  # 성공한 설정 반환
                    
                except Exception as e:
                    print(f"❌ 폴더 접근 실패: {e}")
                    
            except Exception as e:
                print(f"❌ 로그인 실패: {e}")
                
        except socket.gaierror as e:
            print(f"❌ DNS 조회 실패: {e}")
        except ssl.SSLError as e:
            print(f"❌ SSL 연결 실패: {e}")
        except Exception as e:
            print(f"❌ 연결 실패: {e}")
        
        finally:
            try:
                mail.logout()
            except:
                pass
    
    print("\n" + "=" * 60)
    print("❌ 모든 서버 연결 시도 실패")
    print("\n💡 추가 확인사항:")
    print("1. 회사 IT 부서에 IMAP 서버 주소 문의")
    print("2. 방화벽에서 IMAP 포트(993, 143) 허용 확인")
    print("3. 이메일 계정에서 IMAP 사용 설정 확인")
    print("4. 2단계 인증 사용시 앱 비밀번호 생성 필요")
    return None

def test_specific_server(host, port, ssl_enabled, email, password):
    """특정 서버 설정으로 테스트"""
    print(f"🔍 특정 서버 테스트: {host}:{port}")
    
    try:
        if ssl_enabled:
            mail = imaplib.IMAP4_SSL(host, port)
            print("✅ SSL 연결 성공")
        else:
            mail = imaplib.IMAP4(host, port)
            print("✅ 일반 연결 성공")
            try:
                mail.starttls()
                print("✅ STARTTLS 적용 성공")
            except:
                print("⚠️  STARTTLS 적용 실패 (선택사항)")
        
        # 로그인
        mail.login(email, password)
        print("✅ 로그인 성공!")
        
        # 기본 작업 테스트
        mail.select('INBOX')
        print("✅ INBOX 선택 성공")
        
        # 최근 메일 1개 가져오기 테스트
        status, messages = mail.search(None, 'ALL')
        if status == 'OK' and messages[0]:
            latest_msg_id = messages[0].split()[-1]
            status, msg_data = mail.fetch(latest_msg_id, '(BODY[HEADER.FIELDS (FROM SUBJECT DATE)])')
            if status == 'OK':
                print("✅ 메일 헤더 가져오기 성공")
        
        mail.close()
        mail.logout()
        return True
        
    except Exception as e:
        print(f"❌ 테스트 실패: {e}")
        return False

def main():
    """메인 실행 함수"""
    print("📧 이메일 서버 연결 진단 도구")
    print("=" * 60)
    
    # 1. 자동 서버 탐지
    successful_config = test_email_connection()
    
    if not successful_config:
        # 2. 수동 서버 입력 옵션
        print("\n🔧 수동 서버 정보 입력 테스트")
        print("회사 IT 부서에서 제공받은 IMAP 서버 정보를 입력하세요:")
        
        try:
            host = input("IMAP 서버 주소: ").strip()
            port = int(input("IMAP 포트 (993 또는 143): ").strip())
            ssl_choice = input("SSL 사용? (y/n): ").strip().lower()
            ssl_enabled = ssl_choice in ['y', 'yes', '1']
            
            email = input("이메일 주소 [abc@k-fs.co.kr]: ").strip() or "abc@k-fs.co.kr"
            password = input("비밀번호 [1234]: ").strip() or "1234"
            
            success = test_specific_server(host, port, ssl_enabled, email, password)
            
            if success:
                print(f"\n🎉 수동 설정 성공!")
                print(f"✅ 사용할 설정:")
                print(f"   Host: {host}")
                print(f"   Port: {port}")
                print(f"   SSL: {ssl_enabled}")
                
        except KeyboardInterrupt:
            print("\n\n종료되었습니다.")
        except Exception as e:
            print(f"❌ 입력 오류: {e}")

if __name__ == "__main__":
    main()