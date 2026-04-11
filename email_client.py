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
        self.pwd = os.getenv("EMAIL_PASS")
        self.sd = os.getenv("START_DATE")
        self.ed = os.getenv("END_DATE")

    def connect(self):
        try:
            c = imaplib.IMAP4_SSL(self.host, self.port)
            c.login(self.user, self.pwd)
            c.select("INBOX")
            return c
        except:
            return None

    def decode_str(self, s):
        try:
            return "".join(t.decode(c or "utf-8") if isinstance(t, bytes) else str(t) for t, c in decode_header(s))
        except:
            return str(s)

    def save_attach(self, msg, sid):
        for part in msg.walk():
            if part.get_content_disposition() == "attachment":
                fn = part.get_filename()
                if not fn: continue
                fn, c = decode_header(fn)[0]
                if isinstance(fn, bytes): fn = fn.decode(c or "utf-8")
                if fn.lower().endswith(".docx"):
                    os.makedirs("data", exist_ok=True)
                    p = f"data/{sid}_{fn}"
                    with open(p, "wb") as f:
                        f.write(part.get_payload(decode=True))
                    return p
        return ""

    def fetch_mails(self):
        mails = []
        c = self.connect()
        if not c: return mails

        try:
            s = datetime.strptime(self.sd, "%Y-%m-%d").strftime("%d-%b-%Y")
            e = (datetime.strptime(self.ed, "%Y-%m-%d") + timedelta(1)).strftime("%d-%b-%Y")
            _, ids = c.search(None, f'SINCE "{s}" BEFORE "{e}"')
            # 正序，保证时间最早在前
            for mid in ids[0].split():
                try:
                    _, data = c.fetch(mid, "(RFC822)")
                    msg = email.message_from_bytes(data[0][1])
                    subj = self.decode_str(msg["Subject"])
                    date_str = msg.get("Date", "")
                    sid = re.search(r"\d{10}", subj)
                    name = re.search(r"[\u4e00-\u9fa5]{2,4}", subj)
                    if sid:
                        attach = self.save_attach(msg, sid.group())
                        mails.append({
                            "student_id": sid.group(),
                            "name": name.group() if name else "",
                            "attachment_path": attach,
                            "receive_time": date_str
                        })
                except:
                    continue
        except:
            pass
        c.close()
        c.logout()
        return mails
