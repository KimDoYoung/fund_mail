from sqlalchemy import create_engine, Column, Integer, String, Text, DateTime
from sqlalchemy.orm import declarative_base, sessionmaker
import datetime
from config import DB_CONFIG

Base = declarative_base()

class Email(Base):
    __tablename__ = 'emails'
    id = Column(Integer, primary_key=True)
    subject = Column(String(255))
    sender = Column(String(255))
    body = Column(Text)
    date = Column(DateTime)
    attachment_path = Column(String(500), nullable=True)

engine = create_engine(DB_CONFIG['db_url'], echo=False)
Base.metadata.create_all(engine)
Session = sessionmaker(bind=engine)

def save_email(subject, sender, body, date, attachment_path=None):
    session = Session()
    email = Email(subject=subject, sender=sender, body=body, date=date, attachment_path=attachment_path)
    session.add(email)
    session.commit()
    session.close()
