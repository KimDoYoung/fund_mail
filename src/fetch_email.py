import requests
import os
import sqlite3
import paramiko
import json
from msal import ConfidentialClientApplication
from datetime import datetime, time, timezone, timedelta
from pathlib import Path
from zoneinfo import ZoneInfo

from exceptions import AttachFileFetchError, TokenError
from exceptions import EmailFetchError
from logger import get_logger
from db_actions import create_db_tables, save_email_data_to_db
from utils import truncate_filepath  

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

def get_ymd_path_and_dbpath(config, one_day: str = None):
    ''' 현재 날짜를 'YYYY_MM_DD' 형식으로 반환 폴더 경로 및 DB명 생성'''
    data_dir = config.data_dir
    if one_day:
        ymd_time = one_day.replace("-","_") + datetime.now().strftime('_%H_%M')
        ymd_path = data_dir / ymd_time[:10]  # '2025-06-23' 형태
    else:    
        ymd_time = datetime.now().strftime('%Y_%m_%d_%H_%M')
        ymd_path = data_dir / ymd_time[:10]  # '2025-06-23' 형태
    
    if not ymd_path.exists():
        ymd_path.mkdir(parents=True, exist_ok=True)
        logger.info(f"날짜별 폴더 생성: {ymd_path}")

    db_path =  ymd_path / f'fm_{ymd_time}.db'
    return ymd_path, db_path

def save_last_email_id_and_time(last_mail_time, last_email_id, title, config):
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
            data["last_email_id"] = last_email_id
            data["last_fetch_time_kst"] = utc_to_kst(ts, as_iso=True)  # KST로 변환
            data["title"] = title  # 제목 추가
            f.seek(0)
            json.dump(data, f, indent=4, ensure_ascii=False)
            f.truncate()
        logger.info(f"✏️ 마지막 이메일 수집 시각 저장: [{json.dumps(data, ensure_ascii=False)}]")

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



def utc_day_range(date_str: str) -> tuple[str, str]:
    """
    Parameters
    ----------
    date_str : str
        날짜 문자열. 반드시 **'YYYY-MM-DD'** 형식이어야 하며,
        한국 표준시(KST, UTC+09:00)의 그날 0시~24시 범위를 의미한다.

    Returns
    -------
    tuple[str, str]
        (UTC 0시 ISO8601, 다음날 UTC 0시 ISO8601)
        → Microsoft Graph $filter 에 바로 사용할 수 있는 `'Z'` 접미사 ISO-8601.
    """
    # 1) 문자열 → KST 자정으로 변환
    try:
        local_midnight = datetime.strptime(date_str, "%Y-%m-%d").replace(tzinfo=KST)
    except ValueError as e:
        raise ValueError("date_str must be in 'YYYY-MM-DD' format") from e

    # 2) KST → UTC
    start_utc = local_midnight.astimezone(timezone.utc)
    end_utc   = start_utc + timedelta(days=1)

    # 3) 'Z' 접미사 ISO 8601
    to_iso_z = lambda dt: dt.isoformat().replace("+00:00", "Z")
    return to_iso_z(start_utc), to_iso_z(end_utc)

def build_params_for_one_day(target_date: str, top: int = 1000) -> dict:
    start_iso, end_iso = utc_day_range(target_date)

    filter_expr = (
        f"receivedDateTime ge {start_iso} "
        f"and receivedDateTime lt {end_iso}"
    )

    return {
        "$filter":  filter_expr,
        "$orderby": "receivedDateTime desc",
        "$top":     top,
        "$select": ("subject,from,sender,receivedDateTime,hasAttachments,"
                    "id,toRecipients,ccRecipients"),
    }	

def fetch_email_from_office365(config, one_day:str = None):
    """
    LAST_TIME.json 파일에서 마지막 이메일 수집 시각을 읽어오고,
    없으면 그날의 00:00:00 시각을 반환합니다.
    그 시각 이후의 메일을 모두 가져와서 db에 저장, attachments를 다운로드합니다.
    """
    MAIL_USER = config.email_id

    token = get_graph_token(config)
    if not token:
        raise TokenError("❌ Graph API 토큰 발급 실패")

    headers = {
        'Authorization': f'Bearer {token}',
        'Content-Type': 'application/json'
    }
    
    # LAST_TIME.json에서 최종 email_id를 가져온다.
    last_email_id = config.last_email_id
    is_first_fetch = False
    if one_day: # 하루동안의 메일
        logger.info(f"하루동안의 메일을 가져옵니다: {one_day}")
        params = build_params_for_one_day(one_day)
        last_email_id = None  # 하루 단위로 가져오면 마지막 ID는 의미 없음
    elif not last_email_id:
        is_first_fetch = True
        logger.warning("⚠️마지막 이메일 ID가 없으므로(LAST_TIME.json 없는지 체크), 1건만 가져옵니다.")
        params = {
            "$orderby": "receivedDateTime desc",                    # 최신부터
            "$top": 1,                                             # 1개만
            "$select": (
                "subject,from,sender,receivedDateTime,hasAttachments,id,"
                "toRecipients,ccRecipients"
            )
        }    
    else:
        cursor = config.last_mail_fetch_time
        cursor_str = cursor.astimezone(timezone.utc)          \
                        .strftime('%Y-%m-%dT%H:%M:%SZ')  # → '2025-06-23T15:00:00Z'    
        params = {
            # 오늘 00:00 이후 메일 전부
            "$filter": f"receivedDateTime ge {cursor_str}",
            # 페이지당 최대 1000건 (더 많으면 @odata.nextLink 로 자동 페이지 이동)
            "$orderby": "receivedDateTime desc",               # ← 정렬 추가
            "$top": 1000,
            # 원하는 필드만 선택
            "$select": "subject,from,sender, receivedDateTime,hasAttachments,id,toRecipients,ccRecipients"
        }
    
    # cursor = config.last_mail_fetch_time
    # cursor_str = cursor.astimezone(timezone.utc)          \
    #                  .strftime('%Y-%m-%dT%H:%M:%SZ')  # → '2025-06-23T15:00:00Z'    
    # 예: '2025-06-23T15:00:00Z'    
    url = f'https://graph.microsoft.com/v1.0/users/{MAIL_USER}/messages'
    
    try:
        graph = requests.Session()
        graph.headers.update(headers)  # 세션에 헤더 추가
        response = graph.get(url, headers=headers, params=params)
        # 현재 시각을 UTC로 변환
        db_path = None
        ymd_path = None
        if not is_first_fetch:
            ymd_path, db_path = get_ymd_path_and_dbpath(config, one_day)  # ymd_path, db_path 생성
        email_data_list = []
        if response.status_code == 200:
            emails = response.json().get('value', [])
            logger.info(f"✅ {len(emails)}개의 이메일을 가져왔습니다. 새로운 메일:{len(emails)-1}, 마지막 1개는 체크용임")
            last_mail_time = None
            count = 0
            logger.info("--------------------------------------------------------")
            for email in emails:
                email_id = email.get('id', 'ID 없음')
                subject = email.get('subject', '제목 없음')
                from_address = email.get('from', {}).get('emailAddress', {}).get('address', '')
                from_name = email.get('from', {}).get('emailAddress', {}).get('name', '')
                sender_address = email.get('sender', {}).get('emailAddress', {}).get('address', '')
                sender_name = email.get('sender', {}).get('emailAddress', {}).get('name', '')
                received_time = email.get('receivedDateTime', '날짜 없음')
                to_recipients  = ', '.join(r.get('emailAddress', {}).get('address') for r in email.get('toRecipients', []) if r.get('emailAddress', {}).get('address')) or '받는 사람 없음'
                cc_recipients  = ', '.join(r.get('emailAddress', {}).get('address') for r in email.get('ccRecipients', []) if r.get('emailAddress', {}).get('address')) or '참조 없음'

                content = get_message_body(graph, MAIL_USER, email_id) or '내용 없음'                
                kst_time = utc_to_kst(received_time, as_iso=True)  # KST로 변환

                # last_mail_time은 가장 최근 시각으로 설정
                if last_mail_time is None or received_time > last_mail_time:
                    last_mail_time = received_time

                if is_first_fetch:
                    last_email_id = email_id
                    logger.info(f"가장 최근 이메일 수집: {email_id} - {subject}")
                    break
                if last_email_id == email_id:
                    logger.info(f"마지막 이메일 ID와 일치합니다. 수집을 중단합니다.")
                    break    
                # 첨부파일이 있는 경우 다운로드
                attach_files= []
                if email.get('hasAttachments'):
                    attach_files = download_attachments( MAIL_USER, email['id'], headers, ymd_path)
                # 데이터 메모리에 저장
                email_data_list.append({
                    'email_id': email_id,
                    'subject': subject,
                    'sender_address': sender_address,
                    'sender_name': sender_name,
                    'from_address': from_address,
                    'from_name': from_name,
                    'to_recipients': to_recipients,
                    'cc_recipients': cc_recipients,
                    'email_time': received_time,  # UTC 시각
                    'kst_time': kst_time,          # KST 시각
                    'content': content,
                    'attach_files': attach_files if not is_first_fetch else []
                })
                count += 1
                logger.info(f"{count} : {subject} ({kst_time}), 첨부파일 개수: {len(attach_files) if not is_first_fetch else 0}")
            logger.info("--------------------------------------------------------")
            # 처음이면 last_time.json저장    
            if is_first_fetch:
                save_last_email_id_and_time(last_mail_time, last_email_id, subject, config)
            elif email_data_list:
                    last_email_id = email_data_list[0]['email_id']
                    title = email_data_list[0]['subject']
                    create_db_tables(db_path)  # DB 초기화
                    # 마지막 이메일 ID와 시각 저장
                    save_last_email_id_and_time(last_mail_time, last_email_id, title, config)
                    # DB에 저장 
                    db_path = save_email_data_to_db(email_data_list, db_path)
            else:
                kst = utc_to_kst(last_mail_time, as_iso=True) if last_mail_time else '알 수 없음'
                logger.warning(f"⚠️ 시각: {kst} 으로부터 수신된 이메일이 없습니다.")
                db_path = None
            return db_path
        else:
            raise EmailFetchError(f"❌ API 호출 실패: {response.status_code} - {response.text}")
            
    except Exception as e:
        raise EmailFetchError(f"❌ 이메일 수집시 알려지지 않은 오류: {e}")


def download_attachments(MAIL_USER, email_id, headers, ymd_path):
    """첨부파일 다운로드"""
    url = f'https://graph.microsoft.com/v1.0/users/{MAIL_USER}/messages/{email_id}/attachments'
    
    try:
        response = requests.get(url, headers=headers)
        
        if response.status_code == 200:
            attachments = response.json().get('value', [])
            attach_count = 0
            attach_files = []
            # 첨부파일 저장 폴더 생성
            attach_path = ymd_path / 'attach'
            if not attach_path.exists():
                attach_path.mkdir(parents=True, exist_ok=True)
                logger.info(f"✅ 첨부파일 폴더 생성: {attach_path}")

            for attachment in attachments:
                if attachment.get("@odata.type") != "#microsoft.graph.fileAttachment":
                    continue                        # 참조·메시지 첨부 등

                if is_logo_like(attachment):
                    logger.debug("본문 로고로 판단, 저장 생략: %s (%s)",
                                attachment.get("name"), attachment.get("contentType"))
                    continue                
                filename = attachment.get('name')
                # filename = truncate_filename(filename, max_bytes=255)  # 파일명 길이 제한
                content = attachment.get('contentBytes')
                
                if content:
                    import base64
                    file_data = base64.b64decode(content)
                    filepath = os.path.join(attach_path, filename)
                    filepath = if_exist_change_filename(filepath)  # 중복 파일명 처리
                    filepath = truncate_filepath(filepath)  
                    with open(filepath, 'wb') as f:
                        f.write(file_data)
                    attach_files.append({
                        'parent_id': None,
                        'email_id': email_id,
                        'save_folder': str(attach_path),
                        'file_name': os.path.basename(filepath)
                    })
        else:
            raise AttachFileFetchError(f"첨부파일 API 호출 실패: {response.status_code} - {response.text}")
        return attach_files
    except Exception as e:
        raise  AttachFileFetchError(f"❌ {e}")
