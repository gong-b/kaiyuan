import imaplib, ssl, os, base64, logging
from typing import Generator, Tuple
from email import message_from_bytes
from email.message import Message
from config import IMAP_HOST, IMAP_PORT

logger = logging.getLogger(__name__)

def imap_utf7_encode(text):
    """处理中文文件夹名的编码"""
    def _modified_base64(s):
        return base64.b64encode(s.encode('utf-16-be')).decode('ascii').rstrip('=').replace('/', ',')
    res = []
    i = 0
    while i < len(text):
        c = text[i]
        if 0x20 <= ord(c) <= 0x7e:
            res.append('&-') if c == '&' else res.append(c)
            i += 1
        else:
            j = i
            while j < len(text) and not (0x20 <= ord(text[j]) <= 0x7e): j += 1
            res.append('&' + _modified_base64(text[i:j]) + '-')
            i = j
    return "".join(res)

class SecureIMAPClient:
    def __init__(self, user, password, folder="开源课堂"):
        self.user = user
        self.password = password
        self.folder = folder

    def __enter__(self):
        context = ssl.create_default_context()
        self.conn = imaplib.IMAP4_SSL(IMAP_HOST, IMAP_PORT, ssl_context=context)
        self.conn.login(self.user, self.password)
        # 编码文件夹名
        status, _ = self.conn.select(imap_utf7_encode(self.folder))
        if status != "OK":
            raise ValueError(f"无法找到文件夹: {self.folder}")
        return self

    def __exit__(self, *args):
        if self.conn: self.conn.logout()

    def fetch_emails(self, start_date_str):
        status, data = self.conn.uid('SEARCH', 'SINCE', start_date_str)
        if status == 'OK' and data[0]:
            for uid_bytes in data[0].split():
                uid = uid_bytes.decode('utf-8')
                status, full_data = self.conn.uid('FETCH', uid, '(RFC822)')
                yield uid, message_from_bytes(full_data[0][1])
