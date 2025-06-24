import os
from dotenv import load_dotenv

load_dotenv()
EMAIL_CONFIG = {
    #"host": "smtp.office365.com",
    "host": "smtp-mail.outlook.com",
    "port": 143,
    "email": os.getenv("EMAIL_ID"),
    "password": os.getenv("EMAIL_PW"),
}
DB_CONFIG = {
    "db_url": "sqlite:///c:/tmp/fund_mail.db"  # 또는 postgresql://user:pass@localhost/dbname
}
