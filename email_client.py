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

    # ✅ 修复：浙大邮箱必须用英文日期格式（终极修复）
    def format_imap_date(self, date_str):
        dt = datetime.strptime(date_str, "%Y-%m-%d")
        return dt.strftime("%d-%b-%Y")  # 输出 10-Oct-2025

    def fetch_mails(self):
        mails = []
        if not self.connect():
            return mails

        try:
            # ✅ 修复日期格式
            s_date = self.format_imap_date(self.start_date)
            e_date = self.format_imap_date(self.end_date)

            print(f"✅ 搜索日期：{self.start_date} → {s_date}")
            print(f"✅ 搜索日期：{self.end_date} → {e_date}")

            # ✅ 正确的浙大邮箱日期搜索命令
            status, messages = self.conn.search(None, f'SINCE "{s_date}" BEFORE "{e_date}"')
            
            mail_ids = messages[0].split()
            print(f"✅ 找到邮件数量：{len(mail_ids)} 封")

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
                except Exception as e:
                    print(f"⚠️ 跳过一封错误邮件")
                    continue

            self.conn.close()
            self.conn.logout()
        except Exception as ex:
            print(f"错误：{ex}")

        return mails
