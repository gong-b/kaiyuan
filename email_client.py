import imaplib
import email
from email.header import decode_header
import os
import re
from datetime import datetime, timedelta

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
        try:
            for part in msg.walk():
                if part.get_content_disposition() == "attachment":
                    filename = part.get_filename()
                    if not filename:
                        continue
                    fn, enc = decode_header(filename)[0]
                    if isinstance(fn, bytes):
                        fn = fn.decode(enc or "utf-8")
                    # 只处理 DOCX！！！
                    if fn.lower().endswith(".docx"):
                        os.makedirs("data", exist_ok=True)
                        path = f"data/{sid}_{fn}"
                        with open(path, "wb") as f:
                            f.write(part.get_payload(decode=True))
                        return path
        except:
            pass
        return ""

    def to_imap_date(self, date_str):
        dt = datetime.strptime(date_str, "%Y-%m-%d")
        return dt.strftime("%d-%b-%Y")

    def fetch_mails(self):
        mails = []
        if not self.connect():
            return mails

        try:
            start_dt = datetime.strptime(self.start_date, "%Y-%m-%d")
            end_dt = datetime.strptime(self.end_date, "%Y-%m-%d")
            search_start = self.to_imap_date(self.start_date)
            search_end = self.to_imap_date((end_dt + timedelta(days=1)).strftime("%Y-%m-%d"))

            status, messages = self.conn.search(None, f'SINCE "{search_start}" BEFORE "{search_end}"')
            mail_ids = messages[0].split()
            print(f"✅ 共找到邮件：{len(mail_ids)} 封")

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
        except:
            pass
        return mails
