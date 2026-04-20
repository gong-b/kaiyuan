import imaplib
import ssl
import base64
import logging
from email import message_from_bytes

logger = logging.getLogger(__name__)

def imap_utf7_encode(text):
    try:
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
                while j < len(text) and not (0x20 <= ord(text[j]) <= 0x7e):
                    j += 1
                res.append('&' + _modified_base64(text[i:j]) + '-')
                i = j
        return "".join(res)
    except:
        return text

class SecureIMAPClient:
    def __init__(self, user, password, folder="INBOX"):
        self.user = user
        self.password = password
        self.folder = folder

    def __enter__(self):
        ctx = ssl.create_default_context()
        self.conn = imaplib.IMAP4_SSL("imap.zju.edu.cn", 993, ssl_context=ctx)
        self.conn.login(self.user, self.password)
        encoded = imap_utf7_encode(self.folder)
        self.conn.select(encoded)
        return self

    def __exit__(self, *args):
        self.conn.logout()

    def fetch_emails(self, since_date):
        status, data = self.conn.uid('SEARCH', 'SINCE', since_date)
        if not data[0]:
            return []
        for uid in data[0].split():
            status, d = self.conn.uid('FETCH', uid, '(RFC822)')
            if status == 'OK':
                yield uid.decode(), message_from_bytes(d[0][1])
