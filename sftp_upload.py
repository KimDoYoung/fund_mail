import os
import errno
from datetime import datetime, timezone
import paramiko
import sqlite3
from logger import get_logger

logger = get_logger()


def remote_exists(sftp: paramiko.SFTPClient, path: str) -> bool:
    """원격 경로(파일/디렉터리) 존재 여부"""
    try:
        sftp.stat(path)
        return True
    except FileNotFoundError:          # Py3.4+에서 FileNotFoundError
        return False
    except IOError as e:               # 오래된 Paramiko는 IOError(errno=2) 사용
        if e.errno == errno.ENOENT:
            return False
        raise                           # 다른 오류는 그대로 전파

def mkdir_p(sftp: paramiko.SFTPClient, path: str) -> None:
    """`mkdir -p`처럼 상위 디렉터리까지 재귀적으로 생성"""
    parts = []
    while path not in ("", "/"):
        parts.append(path)
        path = os.path.dirname(path)
    for p in reversed(parts):
        if not remote_exists(sftp, p):
            sftp.mkdir(p)
            logger.info(f"SFTP 디렉터리 생성: {p}")


def get_local_attach_file_list(db_path):
    """db_path의 sqlite를 읽어서 첨부파일의 목록을  가져오기"""

    try:
        conn = sqlite3.connect(db_path)
        cur = conn.cursor()
        cur.execute("SELECT save_folder, file_name FROM fund_mail_attach")
        rows = cur.fetchall()
        conn.close()
        
        file_list = [os.path.join(row[0], row[1]) for row in rows]
        return file_list
    except sqlite3.Error as e:
        logger.error(f"❌ DB 첨부파일 목록 조회 오류: {e}")
        return []    



def upload_to_sftp(config, db_path):
    """SFTP 서버에 DB 파일과 첨부파일 업로드"""
    transport = None
    sftp = None
    try:
        transport = paramiko.Transport((config.sftp_host, config.sftp_port))
        transport.connect(username=config.sftp_id, password=config.sftp_pw)
        sftp = paramiko.SFTPClient.from_transport(transport)

        # === 1) DB 파일 업로드 ===
        ymd = datetime.now(timezone.utc).strftime("%Y-%m-%d")
        remote_dir = f"{config.sftp_base_dir}/{ymd}"
        mkdir_p(sftp, remote_dir)  # 디렉터리 생성 (필요 시)

        remote_db_path = f"{remote_dir}/{os.path.basename(db_path)}"
        sftp.put(str(db_path), remote_db_path)
        logger.info(f"✅ DB 파일 SFTP 업로드 완료: {remote_db_path}")

        # === 2) 첨부파일 업로드 ===
        attach_dir = f"{remote_dir}/attach"
        file_list = get_local_attach_file_list(db_path)
        if file_list:                        # 첨부파일이 하나라도 있으면
            mkdir_p(sftp, attach_dir)

        for file_path in file_list:
            if os.path.exists(file_path):
                remote_attach_path = f"{attach_dir}/{os.path.basename(file_path)}"
                sftp.put(file_path, remote_attach_path)
                logger.info(f"첨부파일 SFTP 업로드 완료: {remote_attach_path}")
            else:
                logger.warning(f"첨부파일 경로가 존재하지 않음: {file_path}")

    except Exception as e:
        logger.error(f"❌ SFTP 업로드 오류: {e}")
        raise
    finally:
        # `close()`는 idempotent하므로 중복 호출해도 안전
        if sftp:
            sftp.close()
        if transport:
            transport.close()
        logger.info("SFTP 연결 종료")
