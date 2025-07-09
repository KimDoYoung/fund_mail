import requests
from msal import ConfidentialClientApplication

from config import Config
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

token = get_graph_token(Config.load())
cfg = Config.load()    

MAIL_USER = cfg.email_user_id
url = f'https://graph.microsoft.com/v1.0/users/{MAIL_USER}/messages'
params = {
    "$orderby": "receivedDateTime asc",
    "$top": 1,
    "$select": "id,receivedDateTime,subject"
}

headers = {
    "Authorization": f"Bearer {token}"
    , "Content-Type": "application/json"
}

resp = requests.get(url, params=params, headers=headers, timeout=30)
resp.raise_for_status()

first = resp.json()["value"][0]
print("가장 오래된 메일 시각:", first["receivedDateTime"])
print("제목:", first["subject"])