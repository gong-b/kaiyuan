import imaplib
import email
from email.header import decode_header
import os
import re
from datetime import datetime, timedelta
from io import BytesIO

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

    # ===================== 修复：内存返回附件，不写磁盘 =====================
    def get_attach_memory(self, msg):
        for part in msg.walk():
            if part.get_content_disposition() == "attachment":
                fn = part.get_filename()
                if not fn:
                    continue
                fn = self.decode_str(fn)
                if fn.lower().endswith(".docx"):
                    data = part.get_payload(decode=True)
                    return BytesIO(data)
        return None

    def fetch_mails(self):
        mails = []
        c = self.connect()
        if not c:
            return mails

        try:
            s = datetime.strptime(self.sd, "%Y-%m-%d").strftime("%d-%b-%Y")
            e = (datetime.strptime(self.ed, "%Y-%m-%d") + timedelta(1)).strftime("%d-%b-%Y")
            _, ids = c.search(None, f'SINCE "{s}" BEFORE "{e}"')
            
            for mid in ids[0].split():
                try:
                    _, data = c.fetch(mid, "(RFC822)")
                    msg = email.message_from_bytes(data[0][1])
                    subj = self.decode_str(msg["Subject"])
                    date_str = msg.get("Date", "")
                    sid_match = re.search(r"\d{10}", subj)
                    name_match = re.search(r"[\u4e00-\u9fa5]{2,4}", subj)

                    if sid_match:
                        sid = sid_match.group()
                        name = name_match.group() if name_match else ""
                        # 内存获取附件
                        attach_io = self.get_attach_memory(msg)
                        mails.append({
                            "student_id": sid,
                            "name": name,
                            "attach_io": attach_io,  # 不再用路径
                            "receive_time": date_str
                        })
                except:
                    continue
        except:
            pass
        c.close()
        c.logout()
        return mails
