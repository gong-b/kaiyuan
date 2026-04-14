import imaplib
import ssl
import logging
import os
import binascii
from typing import Generator, Tuple, Optional
from email import message_from_bytes
from email.message import Message
from config import IMAP_HOST, IMAP_PORT, EMAIL_USER, EMAIL_PASSWORD

logger = logging.getLogger(__name__)

def imap_utf7_encode(text):
    """
    简易版 IMAP Modified UTF-7 编码逻辑
    用于处理中文文件夹名称
    """
    import base64
    def _modified_base64(s):
        s_utf16 = s.encode('utf-16-be')
        return base64.b64encode(s_utf16).decode('ascii').rstrip('=').replace('/', ',')

    res = []
    i = 0
    while i < len(text):
        c = text[i]
        if 0x20 <= ord(c) <= 0x7e:
            if c == '&':
                res.append('&-')
            else:
                res.append(c)
            i += 1
        else:
            j = i
            while j < len(text) and not (0x20 <= ord(text[j]) <= 0x7e):
                j += 1
            res.append('&' + _modified_base64(text[i:j]) + '-')
            i = j
    return "".join(res)

class SecureIMAPClient:
    def __init__(self) -> None:
        self.host = IMAP_HOST
        self.port = IMAP_PORT
        self.user = EMAIL_USER
        self.password = EMAIL_PASSWORD
        # 建议直接在代码里写成已经编码好的字符串，或者使用下面的转换函数
        self.mailbox_raw = "其他文件夹/开源课堂" 
        self.conn = None

    def __enter__(self) -> "SecureIMAPClient":
        context = ssl.create_default_context()
        try:
            self.conn = imaplib.IMAP4_SSL(self.host, self.port, ssl_context=context)
            self.conn.login(self.user, self.user_pwd := self.password)
            
            # 关键修改：对中文路径进行编码处理
            encoded_mailbox = imap_utf7_encode(self.mailbox_raw)
            status, _ = self.conn.select(encoded_mailbox)
            
            if status != "OK":
                logger.error(f"无法进入文件夹: {self.mailbox_raw} (编码后: {encoded_mailbox})")
                raise RuntimeError(f"文件夹 {self.mailbox_raw} 不存在")
            logger.info(f"成功进入文件夹: {self.mailbox_raw}")
        except Exception as e:
            logger.error(f"IMAP连接错误: {e}")
            raise
        return self

    def __exit__(self, exc_type, exc_value, traceback) -> None:
        if self.conn:
            try:
                self.conn.close()
                self.conn.logout()
            except:
                pass

    def fetch_emails(self) -> Generator[Tuple[str, Message], None, None]:
        try:
            start_date = os.environ.get("START_DATE", "01-Mar-2025")
            status, data = self.conn.uid('SEARCH', 'SINCE', start_date)
            
            if status != 'OK' or not data[0]:
                return

            uids = data[0].split()
            for uid_bytes in uids:
                uid = uid_bytes.decode('utf-8')
                status, msg_data = self.conn.uid('FETCH', uid, '(RFC822)')
                if status != 'OK': continue
                msg = message_from_bytes(msg_data[0][1])
                yield uid, msg
        except Exception as e:
            logger.error(f"邮件抓取异常: {e}")
