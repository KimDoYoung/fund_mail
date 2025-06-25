import requests
import os
import sqlite3
import paramiko
import json
from msal import ConfidentialClientApplication
from datetime import datetime, time, timezone, timedelta
from pathlib import Path
from zoneinfo import ZoneInfo

from logger import get_logger

logger = get_logger()

KST = ZoneInfo("Asia/Seoul")
IMAGE_SIZE_THRESHOLD = 20 * 1024     # 20 KB 이하 이미지 = 로고로 간주


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
    # ymd_time = datetime.now(timezone.utc).strftime('%Y_%m_%d_%H_%M')
    ymd_time = datetime.now().strftime('%Y_%m_%d_%H_%M')
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
                id INTEGER PRIMARY KEY AUTOINCREMENT,  
                email_id TEXT ,  -- office365의 email_id
                subject TEXT,
                sender TEXT,
                to_recipients TEXT,  -- 수신자 목록
                cc_recipients TEXT,  -- 참조자 목록
                email_time TEXT,
                kst_time TEXT,
                content TEXT
            )
        """)
        cur.execute("""
            CREATE TABLE IF NOT EXISTS fund_mail_attach (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                parent_id INTEGER,
                email_id TEXT,  -- fund_mail 테이블의 id
                save_folder TEXT,
                file_name TEXT
            )
        """)        
        conn.commit()
        conn.close()
        logger.info(f"✅ DB 파일 생성: {db_path}")
    else:
        logger.info(f"✅ DB 파일 경로: {db_path}") 
    return db_path


def save_last_fetch_time(last_mail_time, config):
    """
    마지막 이메일 수집 시각을 LAST_TIME.json에 저장
    """
    last_time_file: Path = config.last_time_file

    # 파일이 없으면 빈 JSON 생성
    if not last_time_file.exists():
        last_time_file.write_text("{}")

    try:
        # 1) 문자열이면 그대로, 2) datetime이면 ISO 문자열로, 3) None이면 현재 UTC
        if isinstance(last_mail_time, str):
            ts = last_mail_time
        elif last_mail_time is None:
            ts = datetime.now(timezone.utc).isoformat()
        else:
            ts = last_mail_time.isoformat()

        with last_time_file.open("r+", encoding="utf-8") as f:
            data = json.load(f)
            data["last_fetch_time"] = ts
            f.seek(0)
            json.dump(data, f, indent=4)
            f.truncate()

        logger.info(f"✅ 마지막 이메일 수집 시각 저장: {ts}")

    except Exception as e:
        logger.error(f"❌ 마지막 이메일 수집 시각 저장 오류: {e}")

def get_message_body(graph: requests.Session,MAIL_USER:str, message_id: str) -> str | None:
    """
    단건 조회로 본문 가져오기 (text 형식).
    graph 세션에는 반드시  Authorization: Bearer <token>  헤더가 포함돼 있어야 합니다.
    """
    # url = f"https://graph.microsoft.com/v1.0/me/messages/{message_id}"
    url = f"https://graph.microsoft.com/v1.0/users/{MAIL_USER}/messages/{message_id}"
    params  = {"$select": "body"}              # body 외 필드가 필요하면 , 로 추가
    # headers = {'Prefer': 'outlook.body-content-type="text"'}  # → 평문으로 받기
    headers = {'Prefer': 'outlook.body-content-type="html"'}  # → 평문으로 받기

    r = graph.get(url, params=params, headers=headers, timeout=30)

    if r.status_code != 200:
        logger.error(f"❌ 메일 본문 조회 실패 {r.status_code}: {r.text}")
        return None

    data = r.json()
    return data.get("body", {}).get("content") or None

def utc_to_kst(utc_str: str, *, as_iso: bool = True) -> str:
    """
    UTC ISO-8601 문자열(Z 포함 가능)을 KST(+09:00) 문자열로 변환한다.

    Parameters
    ----------
    utc_str : str
        예) "2025-06-25T04:47:45Z"  또는  "2025-06-25T04:47:45+00:00"
    as_iso : bool, default True
        True  → '2025-06-25T13:47:45+09:00'  (ISO-8601)
        False → '2025-06-25 13:47:45'        (가독성이 좋은 포맷)

    Returns
    -------
    str
        변환된 KST 시각 문자열. 입력이 None/'' 일 경우 빈 문자열 반환.
    """
    if not utc_str:
        return ""

    # 1) 'Z' 표기를 '+00:00' 로 바꿔야 fromisoformat() 이 읽을 수 있음
    if utc_str.endswith("Z"):
        utc_str = utc_str[:-1] + "+00:00"

    # 2) 문자열 → datetime(UTC)
    dt_utc = datetime.fromisoformat(utc_str).astimezone(timezone.utc)

    # 3) UTC → KST
    dt_kst = dt_utc.astimezone(KST)

    # 4) 원하는 형식으로 반환
    if as_iso:
        return dt_kst.isoformat()           # 2025-06-25T13:47:45+09:00
    else:
        return dt_kst.strftime("%Y-%m-%d %H:%M:%S")

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
        "$select": "subject,from,receivedDateTime,hasAttachments,id,toRecipients,ccRecipients"
    }
    
    try:
        graph = requests.Session()
        graph.headers.update(headers)  # 세션에 헤더 추가
        response = graph.get(url, headers=headers, params=params)
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
                to_recipients  = ', '.join(r['emailAddress']['address'] for r in email.get('toRecipients', [])) or '받는 사람 없음'
                cc_recipients  = ', '.join(r['emailAddress']['address'] for r in email.get('ccRecipients', [])) or '참조 없음'

                content = get_message_body(graph, MAIL_USER, email_id) or '내용 없음'                
                kst_time = utc_to_kst(received_time, as_iso=True)  # KST로 변환
                # DB에 저장
                try:
                    conn = sqlite3.connect(db_path)
                    cur = conn.cursor()
                                    
                    cur.execute("""
                        INSERT OR IGNORE INTO fund_mail (email_id, subject, sender, to_recipients, cc_recipients, email_time, kst_time, content)
                        VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                    """, (email_id, subject, sender, to_recipients, cc_recipients, received_time, kst_time, content))   
                    conn.commit()
                    # insert된 id를 가져오기
                    cur.execute("SELECT last_insert_rowid()")
                    parent_id = cur.fetchone()[0]  # 마지막으로 삽입된 행의

                    conn.close()
                    logger.info(f"✅ 이메일 저장: {subject} "
                                 f" - 발신자: {sender}, 날짜(KST): {kst_time}")
                except sqlite3.Error as e:
                    logger.error(f"❌ DB 저장 오류: {e}")

                # 첨부파일이 있는 경우 다운로드
                if email.get('hasAttachments'):
                    download_attachments(parent_id, MAIL_USER, email['id'], headers, db_path)
                last_mail_time = email.get('receivedDateTime')
                # 마지막 이메일 수집 시각을 저장
            save_last_fetch_time(last_mail_time, config)
            return db_path
        else:
            logger.error(f"❌ API 호출 실패: {response.status_code}")
            logger.error(response.text)
            
    except Exception as e:
        logger.error(f"❌ 오류: {e}")

def is_logo_like(attachment: dict) -> bool:
    """본문 삽입 이미지(로고 등)인지 판단하는 유틸 함수"""
    if attachment.get("isInline"):
        return True
    if attachment.get("contentType", "").startswith("image/") \
       and attachment.get("size", 0) < IMAGE_SIZE_THRESHOLD:
        return True
    return False

def if_exist_change_filename(filepath: str,
                             tz: str = "Asia/Seoul",
                             fmt: str = "%Y%m%d_%H%M%S") -> str:
    """
    같은 디렉터리에 동일한 이름의 파일이 이미 있으면
    `<원본파일명>_<날짜>.<확장자>` 형식으로 이름을 바꾼 새 경로를 돌려 준다.

    Parameters
    ----------
    filepath : str
        원본 파일 전체 경로
    tz : str, default "Asia/Seoul"
        타임스탬프에 사용할 타임존(Olson DB 이름)
    fmt : str, default "%Y%m%d_%H%M%S"
        datetime.strftime 포맷 문자열

    Returns
    -------
    str
        충돌이 없을 때는 filepath,
        존재할 경우에는 "file_YYYYMMDD_HHMMSS.ext" 와 같이 수정된 경로
    """
    if not os.path.exists(filepath):
        return filepath

    dirname, fname = os.path.split(filepath)
    stem, ext = os.path.splitext(fname)

    ts = datetime.now(ZoneInfo(tz)).strftime(fmt)
    candidate = os.path.join(dirname, f"{stem}_{ts}{ext}")

    # 혹시 동일 시각에 두 번 충돌할 수도 있으니 반복으로 안전 확보
    counter = 1
    while os.path.exists(candidate):
        candidate = os.path.join(dirname, f"{stem}_{ts}_{counter}{ext}")
        counter += 1

    return candidate

def download_attachments(parent_id, MAIL_USER, email_id, headers, db_path):
    """첨부파일 다운로드"""
    url = f'https://graph.microsoft.com/v1.0/users/{MAIL_USER}/messages/{email_id}/attachments'
    
    try:
        response = requests.get(url, headers=headers)
        
        if response.status_code == 200:
            attachments = response.json().get('value', [])
            
            for attachment in attachments:
                if attachment.get("@odata.type") != "#microsoft.graph.fileAttachment":
                    continue                        # 참조·메시지 첨부 등

                if is_logo_like(attachment):
                    logger.debug("본문 로고로 판단, 저장 생략: %s (%s)",
                                attachment.get("name"), attachment.get("contentType"))
                    continue                
                filename = attachment.get('name')
                content = attachment.get('contentBytes')
                
                if content:
                    import base64
                    file_data = base64.b64decode(content)
                    attach_path = db_path.parent / 'attach'
                    if not attach_path.exists():
                        attach_path.mkdir(parents=True, exist_ok=True)
                        logger.info(f"✅ 첨부파일 폴더 생성: {attach_path}")
                    filepath = os.path.join(attach_path, filename)
                    filepath = if_exist_change_filename(filepath)  # 중복 파일명 처리
                    with open(filepath, 'wb') as f:
                        f.write(file_data)
                    logger.info(f"✅ 첨부파일 저장: {filename}")
                    # DB에 첨부파일 정보 저장
                    try:
                        conn = sqlite3.connect(db_path)
                        cur = conn.cursor()
                        cur.execute("""
                            INSERT INTO fund_mail_attach (parent_id, email_id, save_folder, file_name)
                            VALUES (?, ?, ?, ?)
                        """, (parent_id, email_id, str(attach_path), filename))  # 현재 작업 디렉토리에 저장
                        conn.commit()
                        conn.close()
                        logger.info(f"✅ 첨부파일 DB 저장: {filename} ({email_id})"
                                        f" - 저장 폴더: {attach_path}")
                    except sqlite3.Error as e:
                        logger.error(f"❌ DB 첨부파일 저장 오류: {e}")
        else:
            logger.error(f"❌ 첨부파일 API 호출 실패: {response.status_code}")
            logger.error(response.text)
                        
    except Exception as e:
        logger.error(f"❌ 첨부파일 다운로드 오류: {e}")
