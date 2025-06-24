import os
import datetime
from imapclient import IMAPClient
import pyzmail
from config import EMAIL_CONFIG
from db import save_email

ATTACHMENT_DIR = './attachments'
os.makedirs(ATTACHMENT_DIR, exist_ok=True)

def fetch_today_emails():
    today = datetime.date.today()
    with IMAPClient(EMAIL_CONFIG['host'], port = EMAIL_CONFIG["port"], ssl=False) as client:
        client.starttls()  # STARTTLS 명시 (SSL 안 쓰는 대신 보안 업그레이드)
        client.login(EMAIL_CONFIG['email'], EMAIL_CONFIG['password'])
        client.select_folder('INBOX')
        
        messages = client.search(['ON', today.strftime('%d-%b-%Y')])
        for uid, message_data in client.fetch(messages, ['RFC822']).items():
            msg = pyzmail.PyzMessage.factory(message_data[b'RFC822'])
            subject = msg.get_subject()
            sender = msg.get_addresses('from')[0][1]
            date = msg.get_decoded_header('Date')
            body = msg.text_part.get_payload().decode(msg.text_part.charset) if msg.text_part else ''
            
            attachment_path = None
            for part in msg.mailparts:
                if part.filename:
                    filepath = os.path.join(ATTACHMENT_DIR, part.filename)
                    with open(filepath, 'wb') as f:
                        f.write(part.get_payload())
                    attachment_path = filepath

            save_email(subject, sender, body, datetime.datetime.now(), attachment_path)

if __name__ == '__main__':
    fetch_today_emails()
