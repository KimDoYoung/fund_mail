import platform
import os
import hashlib

def truncate_filepath(filepath: str, max_path_length: int = 255) -> str:
    """
    너무 긴 파일 경로의 파일 이름을 잘라서 전체 경로 길이를 줄이는 함수.
    디렉토리 구조는 유지하고, 파일 이름만 축소.
    """
    dir_path, filename = os.path.split(filepath)
    name, ext = os.path.splitext(filename)

    current_path_length = len(filepath.encode('utf-8'))
    if current_path_length <= max_path_length:
        return filepath  # 변경 불필요

    # 경로 길이만큼 여유 길이 계산
    allowed_filename_length = max_path_length - len(dir_path.encode('utf-8')) - 1  # '/' 포함
    hash_part = hashlib.md5(filename.encode('utf-8')).hexdigest()[:8]
    allowed_name_length = allowed_filename_length - len(ext.encode('utf-8')) - len(hash_part.encode('utf-8')) - 1

    # 파일 이름 자르기 (byte 기준)
    name_bytes = name.encode('utf-8')[:allowed_name_length]
    safe_name = name_bytes.decode('utf-8', 'ignore')  # UTF-8 안전 디코딩
    safe_filename = f"{safe_name}_{hash_part}{ext}"

    return os.path.join(dir_path, safe_filename)

def truncate_filename(filename, max_bytes=None, preserve_extension=True, encoding='utf-8'):
    """
    UTF-8 바이트 길이를 고려한 크로스 플랫폼 파일명 길이 제한 함수
    
    Args:
        filename (str): 원본 파일명
        max_bytes (int, optional): 최대 바이트 길이. None이면 플랫폼별 기본값 사용
        preserve_extension (bool): 확장자 보존 여부 (기본값: True)
        encoding (str): 파일명 인코딩 (기본값: 'utf-8')
    
    Returns:
        str: 바이트 길이 제한된 파일명
    """
    if max_bytes is None:
        # 플랫폼별 안전한 최대 바이트 길이 설정
        if platform.system() == "Windows":
            max_bytes = 255  # NTFS 기본 제한 (바이트)
        else:
            max_bytes = 255  # 대부분의 Linux 파일시스템 제한 (바이트)
    
    # 경로 구분자 제거 (파일명만 처리)
    filename = os.path.basename(filename)
    
    # 현재 파일명의 바이트 길이 확인
    filename_bytes = filename.encode(encoding)
    if len(filename_bytes) <= max_bytes:
        return filename
    
    if preserve_extension:
        # 확장자 분리
        name, ext = os.path.splitext(filename)
        
        # 확장자의 바이트 길이 확인
        ext_bytes = ext.encode(encoding)
        ext_byte_len = len(ext_bytes)
        
        # 확장자가 전체 길이의 절반을 넘으면 잘라내기
        if ext_byte_len > max_bytes // 2:
            ext = _truncate_string_by_bytes(ext, max_bytes // 2, encoding)
            ext_byte_len = len(ext.encode(encoding))
        
        # 이름 부분에 할당할 수 있는 바이트 길이 계산
        max_name_bytes = max_bytes - ext_byte_len
        
        if max_name_bytes > 0:
            name = _truncate_string_by_bytes(name, max_name_bytes, encoding)
        else:
            # 확장자가 너무 길어서 이름이 들어갈 공간이 없는 경우
            name = ""
            ext = _truncate_string_by_bytes(ext, max_bytes, encoding)
        
        return name + ext
    else:
        # 확장자 보존하지 않고 바이트 기준으로 자르기
        return _truncate_string_by_bytes(filename, max_bytes, encoding)


def _truncate_string_by_bytes(text, max_bytes, encoding='utf-8'):
    """
    바이트 길이를 기준으로 문자열을 안전하게 자르는 내부 함수
    멀티바이트 문자가 중간에 잘리지 않도록 보장
    
    Args:
        text (str): 자를 문자열
        max_bytes (int): 최대 바이트 길이
        encoding (str): 인코딩 방식
    
    Returns:
        str: 바이트 길이로 잘린 문자열
    """
    if not text:
        return text
    
    text_bytes = text.encode(encoding)
    if len(text_bytes) <= max_bytes:
        return text
    
    # 바이트 단위로 자르되, 유효한 UTF-8 문자 경계에서 자르기
    truncated_bytes = text_bytes[:max_bytes]
    
    # UTF-8 문자가 중간에 잘렸는지 확인하고 보정
    try:
        # 디코딩 시도
        result = truncated_bytes.decode(encoding)
        return result
    except UnicodeDecodeError:
        # 멀티바이트 문자가 중간에 잘린 경우, 안전한 지점까지 뒤로 이동
        for i in range(1, 4):  # UTF-8은 최대 4바이트
            if max_bytes - i < 0:
                break
            try:
                result = text_bytes[:max_bytes - i].decode(encoding)
                return result
            except UnicodeDecodeError:
                continue
        
        # 그래도 실패하면 빈 문자열 반환
        return ""