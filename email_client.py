import imaplib
import email
from email.header import decode_header
import os
import re
from datetime import datetime
import config

class EmailClient:
    def __init__(self):
        self.host = config.EMAIL_HOST
        self.port = config.EMAIL_PORT
        self.user = config.EMAIL_USER
        self.password = config.EMAIL_PASS
        self.conn = None

    def connect(self):
        try:
            self.conn = imaplib.IMAP4_SSL(self.host, self.port)
            self.conn.login(self.user, self.password)
            self.conn.select("INBOX")
            return True
        except Exception as e:
            print(f"邮箱连接失败: {e}")
            return False

    def fetch_mails(self):
        mails = []
        if not self.connect():
            return mails

        try:
            status, messages = self.conn.search(None, 'ALL')
            mail_ids = messages[0].split()
            for mail_id in reversed(mail_ids):
                res, msg_data = self.conn.fetch(mail_id, "(RFC822)")
                for response_part in msg_data:
                    if isinstance(response_part, tuple):
                        msg = email.message_from_bytes(response_part[1])
                        subject = self.decode_subject(msg["Subject"])
                        student_id, name = self.extract_info(subject)
                        attachment_path = self.save_attachment(msg, student_id)
                        mails.append({
                            "student_id": student_id,
                            "name": name,
                            "subject": subject,
                            "attachment_path": attachment_path
                        })
            self.conn.close()
            self.conn.logout()
        except:
            pass
        return mails

    def decode_subject(self, s):
        try:
            return "".join([str(t, c or "utf-8") if isinstance(t, bytes) else t for t, c in decode_header(s)])
        except:
            return str(s)

    def extract_info(self, subject):
        sid = re.search(r"(\d{10,})", subject)
        name = re.search(r"[\u4e00-\u9fa5]{2,4}", subject)
        return (sid.group(1) if sid else "", name.group(0) if name else "")

    def save_attachment(self, msg, student_id):
        for part in msg.walk():
            if part.get_content_disposition() == "attachment":
                filename = part.get_filename()
                if filename and filename.endswith(".docx"):
                    os.makedirs("data", exist_ok=True)
                    path = f"data/{student_id}_{filename}"
                    with open(path, "wb") as f:
                        f.write(part.get_payload(decode=True))
                    return path
        return ""
