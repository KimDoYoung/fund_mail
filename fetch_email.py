import requests
import os
import sqlite3
import paramiko
import json
from msal import ConfidentialClientApplication
from datetime import datetime, time, timezone, timedelta

from logger import get_logger

logger = get_logger()

def get_graph_token(config):
    """Microsoft Graph API용 토큰 발급"""
    TENANT_ID = config.tenant_id
    CLIENT_ID = config.client_id
    CLIENT_SECRET = config.client_secret

    authority = f'https://login.microsoftonline.com/{TENANT_ID}'
    app = ConfidentialClientApplication(
        CLIENT_ID, authority=authority, client_credential=CLIENT_SECRET)
    
    # Graph API 스코프
    result = app.acquire_token_for_client(['https://graph.microsoft.com/.default'])
    
    if 'access_token' in result:
        return result.get('access_token')
    else:
        logger.error("토큰 발급 실패:", result.get('error_description'))
        return None
    
def init_db_path(config):
    """DB 파일 경로 초기화"""
    data_dir = config.data_dir
    ymd_time = datetime.now(timezone.utc).strftime('%Y_%m_%d_%H%M')
    ymd_path = data_dir / ymd_time[:10]  # '2025-06-23' 형태
    
    if not ymd_path.exists():
        ymd_path.mkdir(parents=True, exist_ok=True)
        logger.info(f"날짜별 폴더 생성: {ymd_path}")

    db_path =  ymd_path / f'fm_{ymd_time}.db'
    
    if not db_path.exists():
        conn = sqlite3.connect(db_path)
        cur = conn.cursor()
        cur.execute("""
            CREATE TABLE IF NOT EXISTS fund_mail (
                id TEXT PRIMARY KEY,  -- office365의 email_id
                subject TEXT,
                sender TEXT,
                email_time TEXT,
                content TEXT
            )
        """)
        cur.execute("""
            CREATE TABLE IF NOT EXISTS fund_mail_attach (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                parent_id TEXT,
                save_folder TEXT,
                file_name TEXT
            )
        """)        
        conn.commit()
        conn.close()
        logger.info(f"DB 파일 생성: {db_path}")
    else:
        logger.info(f"DB 파일 경로: {db_path}") 
    return db_path

def save_last_fetch_time(last_mail_time, config):
    """마지막 이메일 수집 시각을 LAST_TIME.json에 저장"""
    last_time_file = config.last_time_file
    if not last_time_file.exists():
        with open(last_time_file, 'w') as f:
            f.write('{}')  # 빈 JSON 객체 생성
    
    try:
        with open(last_time_file, 'r+') as f:
            data = json.load(f)
            data['last_fetch_time'] = last_mail_time.isoformat() if last_mail_time else datetime.now(timezone.utc).isoformat()
            f.seek(0)
            json.dump(data, f, indent=4)
            f.truncate()
        logger.info(f"마지막 이메일 수집 시각 저장: {data['last_fetch_time']}")
    except Exception as e:
        logger.error(f"마지막 이메일 수집 시각 저장 오류: {e}")    

def fetch_email_from_office365(config):
    """
    LAST_TIME.json 파일에서 마지막 이메일 수집 시각을 읽어오고,
    없으면 그날의 00:00:00 시각을 반환합니다.
    그 시각 이후의 메일을 모두 가져와서 db에 저장, attachments를 다운로드합니다.
    """
    MAIL_USER = config.email_id

    token = get_graph_token(config)
    if not token:
        return
    
    headers = {
        'Authorization': f'Bearer {token}',
        'Content-Type': 'application/json'
    }
    
    # 특정 사용자의 이메일 가져오기
    cursor = config.last_mail_fetch_time
    cursor_str = cursor.astimezone(timezone.utc)          \
                     .strftime('%Y-%m-%dT%H:%M:%SZ')  # → '2025-06-23T15:00:00Z'    
    # 예: '2025-06-23T15:00:00Z'    
    url = f'https://graph.microsoft.com/v1.0/users/{MAIL_USER}/messages'
    params = {
        # 오늘 00:00 이후 메일 전부
        "$filter": f"receivedDateTime ge {cursor_str}",
        # 페이지당 최대 1000건 (더 많으면 @odata.nextLink 로 자동 페이지 이동)
        "$orderby": "receivedDateTime asc",               # ← 정렬 추가
        "$top": 1000,
        # 원하는 필드만 선택
        "$select": "subject,from,receivedDateTime,hasAttachments,id"
    }
    
    try:
        response = requests.get(url, headers=headers, params=params)
        # 현재 시각을 UTC로 변환
        # ymd_time = datetime.now(timezone.utc).strftime('%Y-%m-%d_%H_%M_%S')
        db_path = init_db_path(config)
        if response.status_code == 200:
            emails = response.json().get('value', [])
            logger.info(f"✅ {len(emails)}개의 이메일을 가져왔습니다:")
            last_mail_time = None
            for email in emails:
                email_id = email.get('id', 'ID 없음')
                subject = email.get('subject', '제목 없음')
                sender = email.get('from', {}).get('emailAddress', {}).get('address', '발신자 없음')
                received_time = email.get('receivedDateTime', '날짜 없음')
                content = email.get('body', {}).get('content', '내용 없음')
                # DB에 저장
                try:
                    conn = sqlite3.connect(db_path)
                    cur = conn.cursor()
                    cur.execute("""
                        INSERT OR IGNORE INTO fund_mail (id, subject, sender, email_time, content)
                        VALUES (?, ?, ?, ?, ?)
                    """, (email_id, subject, sender, received_time, content))   
                    conn.commit()
                    conn.close()
                    logger.info(f"📧 이메일 저장: {subject} ({email_id})"
                                 f" - 발신자: {sender}, 날짜: {received_time}")
                except sqlite3.Error as e:
                    logger.error(f"DB 저장 오류: {e}")

                # 첨부파일이 있는 경우 다운로드
                if email.get('hasAttachments'):
                    download_attachments(email['id'], headers, db_path)
                last_mail_time = email.get('receivedDateTime')
                # 마지막 이메일 수집 시각을 저장
            save_last_fetch_time(last_mail_time, config)
            return db_path
        else:
            logger.error(f"❌ API 호출 실패: {response.status_code}")
            logger.error(response.text)
            
    except Exception as e:
        logger.error(f"❌ 오류: {e}")

def download_attachments(email_id, headers, db_path):
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
                        attach_path = db_path.parent / 'attach'
                        if not attach_path.exists():
                            attach_path.mkdir(parents=True, exist_ok=True)
                            logger.info(f"첨부파일 폴더 생성: {attach_path}")
                        filepath = os.path.join(attach_path, filename)
                        with open(filepath, 'wb') as f:
                            f.write(file_data)
                        logger.info(f"📎 첨부파일 저장: {filename}")
                        # DB에 첨부파일 정보 저장
                        try:
                            conn = sqlite3.connect(db_path)
                            cur = conn.cursor()
                            cur.execute("""
                                INSERT INTO fund_mail_attach (parent_id, save_folder, file_name)
                                VALUES (?, ?, ?)
                            """, (email_id, attach_path, filename))  # 현재 작업 디렉토리에 저장
                            conn.commit()
                            conn.close()
                            logger.info(f"📎 첨부파일 DB 저장: {filename} ({email_id})"
                                         f" - 저장 폴더: {attach_path}")
                        except sqlite3.Error as e:
                            logger.error(f"DB 첨부파일 저장 오류: {e}")
        else:
            logger.error(f"❌ 첨부파일 API 호출 실패: {response.status_code}")
            logger.error(response.text)
                        
    except Exception as e:
        logger.error(f"첨부파일 다운로드 오류: {e}")

def upload_to_sftp(config, db_path):
    """SFTP 서버에 DB 파일 업로드"""

    sftp_host = config.sftp_host
    sftp_port = config.sftp_port
    sftp_id = config.sftp_id
    sftp_pw = config.sftp_pw
    sftp_base_dir = config.sftp_base_dir

    try:
        transport = paramiko.Transport((sftp_host, sftp_port))
        transport.connect(username=sftp_id, password=sftp_pw)
        sftp = paramiko.SFTPClient.from_transport(transport)

        # DB 파일 업로드
        remote_path = f"{sftp_base_dir}/{db_path.name}"  # SFTP 서버의 경로
        sftp.put(str(db_path), remote_path)
        logger.info(f"DB 파일 SFTP 업로드 완료: {remote_path}")

        sftp.close()
        transport.close()
    except Exception as e:
        logger.error(f"SFTP 업로드 오류: {e}")