import imaplib, ssl, logging, os, base64
from typing import Generator, Tuple
from email import message_from_bytes
from email.message import Message
from config import IMAP_HOST, IMAP_PORT

logger = logging.getLogger(__name__)

def imap_utf7_encode(text):
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
    def __init__(self) -> None:
        self.host = IMAP_HOST
        self.port = IMAP_PORT
        self.user = os.environ.get("EMAIL_USER")
        self.password = os.environ.get("EMAIL_PASSWORD")
        self.mailbox_raw = "开源课堂"
        self.conn = None

    def __enter__(self):
        context = ssl.create_default_context()
        self.conn = imaplib.IMAP4_SSL(self.host, self.port, ssl_context=context)
        self.conn.login(self.user, self.password)
        encoded_mailbox = imap_utf7_encode(self.mailbox_raw)
        status, _ = self.conn.select(encoded_mailbox)
        if status != "OK":
            # 自动降级尝试 INBOX
            self.conn.select("INBOX")
            logger.warning(f"未能找到文件夹 {self.mailbox_raw}，已切换至收件箱")
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        if self.conn:
            self.conn.logout()

    def fetch_emails(self) -> Generator[Tuple[str, Message], None, None]:
        from main import parse_subject, parse_subject_pattern
        start_date = os.environ.get("START_DATE", "01-Mar-2025")
        status, data = self.conn.uid('SEARCH', 'SINCE', start_date)
        if status == 'OK' and data[0]:
            for uid_bytes in data[0].split():
                uid = uid_bytes.decode('utf-8')
                # 预览头信息减少流量
                _, h_data = self.conn.uid('FETCH', uid, '(BODY.PEEK[HEADER.FIELDS (SUBJECT DATE)])')
                header_msg = message_from_bytes(h_data[0][1])
                if parse_subject_pattern(parse_subject(header_msg))[1]:
                    _, full_data = self.conn.uid('FETCH', uid, '(RFC822)')
                    yield uid, message_from_bytes(full_data[0][1])
