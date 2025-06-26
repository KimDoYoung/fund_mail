# exceptions.py
class FundMailError(Exception):
    """fund_mail 공통 최상위 예외"""

class TokenError(FundMailError):
    """Graph API 토큰 발급 실패"""

class EmailFetchError(FundMailError):
    """메일 목록/본문 조회 실패"""

class DBWriteError(FundMailError):
    """DB INSERT/UPDATE 실패"""

class SFTPUploadError(FundMailError):
    """SFTP 업로드 실패"""
