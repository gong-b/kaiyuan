import imaplib
import email
from email.header import decode_header
import os
import re
from datetime import datetime

class EmailClient:
    def __init__(self):
        self.host = "imap.zju.edu.cn"
        self.port = 993
        self.user = os.getenv("EMAIL_USER")
        self.password = os.getenv("EMAIL_PASS")
        self.start_date = os.getenv("START_DATE")
        self.end_date = os.getenv("END_DATE")
        self.conn = None

    def connect(self):
        try:
            self.conn = imaplib.IMAP4_SSL(self.host, self.port)
            self.conn.login(self.user, self.password)
            self.conn.select("INBOX")
            return True
        except Exception as e:
            print(f"🔴 邮箱连接失败: {e}")
            return False

    def decode_subject(self, subject):
        try:
            decoded = decode_header(subject)
            return "".join([str(t, c or "utf-8") if isinstance(t, bytes) else str(t) for t, c in decoded])
        except:
            return str(subject)

    def extract_info(self, subject):
        sid = re.search(r"(\d{10})", subject)
        name = re.search(r"([\u4e00-\u9fa5]{2,4})", subject)
        return (sid.group(1) if sid else "", name.group(1) if name else "")

    def save_attachment(self, msg, sid):
        for part in msg.walk():
            if part.get_content_disposition() == "attachment":
                fn = part.get_filename()
                if fn and fn.endswith(".docx"):
                    os.makedirs("data", exist_ok=True)
                    path = f"data/{sid}_{fn}"
                    with open(path, "wb") as f:
                        f.write(part.get_payload(decode=True))
                    return path
        return ""

    def format_date(self, d):
        return datetime.strptime(d, "%Y-%m-%d").strftime("%d-%b-%Y")

    def fetch_mails(self):
        mails = []
        if not self.connect():
            return mails

        try:
            s = self.format_date(self.start_date)
            e = self.format_date(self.end_date)
            criterion = f'(SINCE "{s}" BEFORE "{e}")'
            status, messages = self.conn.search(None, criterion)
            mail_ids = messages[0].split()
            print(f"✅ 筛选 {self.start_date} ~ {self.end_date} 邮件，共找到：{len(mail_ids)} 封")

            for mail_id in reversed(mail_ids):
                try:
                    res, data = self.conn.fetch(mail_id, "(RFC822)")
                    for part in data:
                        if isinstance(part, tuple):
                            msg = email.message_from_bytes(part[1])
                            subject = self.decode_subject(msg["Subject"])
                            sid, name = self.extract_info(subject)
                            if not sid:
                                continue
                            attach = self.save_attachment(msg, sid)
                            mails.append({
                                "student_id": sid,
                                "name": name,
                                "subject": subject,
                                "attachment_path": attach
                            })
                except:
                    continue
            self.conn.close()
            self.conn.logout()
        except Exception as ex:
            print(f"错误：{ex}")
        return mails
