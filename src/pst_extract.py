import argparse
import os
import sys
from pathlib import Path
from pypff import file as pff_file
from datetime import datetime
import sqlite3
from logger import logger
from db_actions import create_db_tables, save_email_data_to_db

def extract_pst_to_sqlite(pst_file_path: str, target_folder: str):
    """PST 파일을 읽어서 SQLite DB와 첨부파일로 변환"""
    
    pst_path = Path(pst_file_path)
    if not pst_path.exists():
        logger.error(f"❌ PST 파일이 존재하지 않습니다: {pst_file_path}")
        return False
        
    target_path = Path(target_folder)
    target_path.mkdir(parents=True, exist_ok=True)
    
    # DB 파일명: PST 파일명과 동일하게 설정
    db_name = pst_path.stem + ".db"
    db_path = target_path / db_name
    
    # 첨부파일 저장 폴더
    attach_folder = target_path / "attach"
    attach_folder.mkdir(parents=True, exist_ok=True)
    
    logger.info(f"📂 PST 파일 처리 시작: {pst_file_path}")
    logger.info(f"📂 DB 저장 경로: {db_path}")
    logger.info(f"📂 첨부파일 폴더: {attach_folder}")
    
    try:
        # PST 파일 열기
        pst_file = pff_file()
        pst_file.open(str(pst_path))
        
        # DB 테이블 생성
        create_db_tables(db_path)
        
        email_data_list = []
        
        # PST 내 폴더 순회
        root_folder = pst_file.get_root_folder()
        process_folder_recursive(root_folder, email_data_list, attach_folder)
        
        # DB에 일괄 저장
        if email_data_list:
            save_email_data_to_db(email_data_list, db_path)
            logger.info(f"✅ 총 {len(email_data_list)}개의 이메일을 처리했습니다.")
        else:
            logger.warning("⚠️ 처리할 이메일이 없습니다.")
            
        pst_file.close()
        return True
        
    except Exception as e:
        logger.error(f"❌ PST 처리 중 오류 발생: {e}")
        return False

def process_folder_recursive(folder, email_data_list, attach_folder):
    """폴더를 재귀적으로 처리하여 모든 이메일 추출"""
    
    try:
        # 현재 폴더의 메시지 처리
        for i in range(folder.get_number_of_messages()):
            message = folder.get_message(i)
            if message:
                email_data = extract_email_data(message, attach_folder)
                if email_data:
                    email_data_list.append(email_data)
        
        # 하위 폴더 재귀 처리
        for i in range(folder.get_number_of_sub_folders()):
            sub_folder = folder.get_sub_folder(i)
            if sub_folder:
                process_folder_recursive(sub_folder, email_data_list, attach_folder)
                
    except Exception as e:
        logger.error(f"❌ 폴더 처리 중 오류: {e}")

def extract_email_data(message, attach_folder):
    """단일 이메일 메시지에서 데이터 추출"""
    
    try:
        # 기본 정보 추출
        email_id = str(message.get_identifier()) if message.get_identifier() else f"pst_{id(message)}"
        subject = message.get_subject() or "제목 없음"
        sender = message.get_sender_name() or "발신자 없음"
        
        # 수신 시간
        delivery_time = message.get_delivery_time()
        if delivery_time:
            email_time = delivery_time.strftime('%Y-%m-%dT%H:%M:%S.%fZ')
            kst_time = delivery_time.strftime('%Y-%m-%dT%H:%M:%S')
        else:
            now = datetime.now()
            email_time = now.strftime('%Y-%m-%dT%H:%M:%S.%fZ')
            kst_time = now.strftime('%Y-%m-%dT%H:%M:%S')
        
        # 수신자 정보
        to_recipients = message.get_display_to() or "수신자 없음"
        cc_recipients = message.get_display_cc() or ""
        
        # 본문 내용
        content = ""
        if message.get_html_body():
            content = message.get_html_body()
        elif message.get_plain_text_body():
            content = message.get_plain_text_body()
        else:
            content = "내용 없음"
        
        # 첨부파일 처리
        attach_files = []
        for i in range(message.get_number_of_attachments()):
            attachment = message.get_attachment(i)
            if attachment:
                attach_file = save_attachment(attachment, attach_folder, email_id)
                if attach_file:
                    attach_files.append(attach_file)
        
        return {
            'email_id': email_id,
            'subject': subject,
            'sender': sender,
            'to_recipients': to_recipients,
            'cc_recipients': cc_recipients,
            'email_time': email_time,
            'kst_time': kst_time,
            'content': content,
            'attach_files': attach_files
        }
        
    except Exception as e:
        logger.error(f"❌ 이메일 데이터 추출 오류: {e}")
        return None

def save_attachment(attachment, attach_folder, email_id):
    """첨부파일 저장"""
    
    try:
        filename = attachment.get_name()
        if not filename:
            filename = f"attachment_{id(attachment)}"
        
        # 안전한 파일명으로 변경
        safe_filename = "".join(c for c in filename if c.isalnum() or c in "._-")
        if not safe_filename:
            safe_filename = f"attachment_{id(attachment)}"
        
        # 이메일 ID별 하위 폴더 생성
        email_attach_folder = attach_folder / email_id
        email_attach_folder.mkdir(parents=True, exist_ok=True)
        
        file_path = email_attach_folder / safe_filename
        
        # 첨부파일 데이터 읽기 및 저장
        data = attachment.read_buffer(attachment.get_size())
        if data:
            with open(file_path, 'wb') as f:
                f.write(data)
            
            logger.info(f"💾 첨부파일 저장: {file_path}")
            return {
                'save_folder': str(email_attach_folder.relative_to(attach_folder.parent)),
                'file_name': safe_filename
            }
        
    except Exception as e:
        logger.error(f"❌ 첨부파일 저장 오류: {e}")
    
    return None

def main():
    parser = argparse.ArgumentParser(description='PST 파일을 SQLite DB로 변환')
    parser.add_argument('pst_file', help='변환할 PST 파일 경로')
    parser.add_argument('target_folder', help='결과물 저장할 폴더 경로')
    
    args = parser.parse_args()
    
    logger.info("=" * 60)
    logger.info("🔄 PST to SQLite 변환 시작")
    logger.info("=" * 60)
    
    success = extract_pst_to_sqlite(args.pst_file, args.target_folder)
    
    if success:
        logger.info("=" * 60)
        logger.info("✅ PST 변환 완료")
        logger.info("=" * 60)
        sys.exit(0)
    else:
        logger.error("=" * 60)
        logger.error("❌ PST 변환 실패")
        logger.error("=" * 60)
        sys.exit(1)

if __name__ == "__main__":
    main()