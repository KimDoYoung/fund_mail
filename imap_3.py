import requests
import os
from msal import ConfidentialClientApplication
from dotenv import load_dotenv

load_dotenv()

TENANT_ID = os.getenv('TENANT_ID')
CLIENT_ID = os.getenv('CLIENT_ID')
CLIENT_SECRET = os.getenv('CLIENT_SECRET')
MAIL_USER = os.getenv('EMAIL_ID')

def get_graph_token():
    """Microsoft Graph API용 토큰 발급"""
    authority = f'https://login.microsoftonline.com/{TENANT_ID}'
    app = ConfidentialClientApplication(
        CLIENT_ID, authority=authority, client_credential=CLIENT_SECRET)
    
    # Graph API 스코프
    result = app.acquire_token_for_client(['https://graph.microsoft.com/.default'])
    
    if 'access_token' in result:
        return result.get('access_token')
    else:
        print("토큰 발급 실패:", result.get('error_description'))
        return None

def get_emails_via_graph():
    """Graph API로 이메일 가져오기"""
    token = get_graph_token()
    if not token:
        return
    
    headers = {
        'Authorization': f'Bearer {token}',
        'Content-Type': 'application/json'
    }
    
    # 특정 사용자의 이메일 가져오기
    url = f'https://graph.microsoft.com/v1.0/users/{MAIL_USER}/messages'
    params = {
        '$top': 5,  # 최근 5개만
        '$select': 'subject,from,receivedDateTime,hasAttachments,id'
    }
    
    try:
        response = requests.get(url, headers=headers, params=params)
        
        if response.status_code == 200:
            emails = response.json().get('value', [])
            print(f"✅ {len(emails)}개의 이메일을 가져왔습니다:")
            
            for email in emails:
                print(f"\n제목: {email.get('subject')}")
                print(f"발신자: {email.get('from', {}).get('emailAddress', {}).get('address')}")
                print(f"날짜: {email.get('receivedDateTime')}")
                print(f"첨부파일: {'있음' if email.get('hasAttachments') else '없음'}")
                
                # 첨부파일이 있는 경우 다운로드
                if email.get('hasAttachments'):
                    download_attachments(email['id'], headers)
        else:
            print(f"❌ API 호출 실패: {response.status_code}")
            print(response.text)
            
    except Exception as e:
        print(f"❌ 오류: {e}")

def download_attachments(email_id, headers):
    """첨부파일 다운로드"""
    url = f'https://graph.microsoft.com/v1.0/users/{MAIL_USER}/messages/{email_id}/attachments'
    
    try:
        response = requests.get(url, headers=headers)
        
        if response.status_code == 200:
            attachments = response.json().get('value', [])
            
            for attachment in attachments:
                if attachment.get('@odata.type') == '#microsoft.graph.fileAttachment':
                    filename = attachment.get('name')
                    content = attachment.get('contentBytes')
                    
                    if content:
                        import base64
                        file_data = base64.b64decode(content)
                        
                        with open(filename, 'wb') as f:
                            f.write(file_data)
                        print(f"📎 첨부파일 저장: {filename}")
                        
    except Exception as e:
        print(f"첨부파일 다운로드 오류: {e}")

if __name__ == '__main__':
    print("=== Microsoft Graph API로 이메일 가져오기 ===")
    get_emails_via_graph()