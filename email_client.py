import asyncio
import aioimaplib
import ssl
import logging
import os
import base64
import re
from typing import AsyncGenerator, Optional
from email import message_from_bytes
from email.message import Message
from config import IMAP_HOST, IMAP_PORT, EMAIL_USER, EMAIL_PASSWORD

logger = logging.getLogger(__name__)

def imap_utf7_encode(text):
    def _modified_base64(s):
        s_utf16 = s.encode('utf-16-be')
        return base64.b64encode(s_utf16).decode('ascii').rstrip('=').replace('/', ',')
    res = []
    i = 0
    while i < len(text):
        c = text[i]
        if 0x20 <= ord(c) <= 0x7e:
            res.append('&-' if c == '&' else c)
            i += 1
        else:
            j = i
            while j < len(text) and not (0x20 <= ord(text[j]) <= 0x7e):
                j += 1
            res.append('&' + _modified_base64(text[i:j]) + '-')
            i = j
    return "".join(res)

class AsyncSecureIMAPClient:
    def __init__(self):
        self.host = IMAP_HOST
        self.port = IMAP_PORT
        self.user = EMAIL_USER
        self.password = EMAIL_PASSWORD
        self.mailbox_raw = "开源课堂"
        self.conn: Optional[aioimaplib.IMAP4_SSL] = None

    async def __aenter__(self):
        self.conn = aioimaplib.IMAP4_SSL(self.host, self.port)
        await self.conn.wait_hello_from_server()
        await self.conn.login(self.user, self.password)
        encoded_box = imap_utf7_encode(self.mailbox_raw)
        await self.conn.select(encoded_box)
        logger.info(f"已进入邮箱文件夹: {self.mailbox_raw}")
        return self

    async def __aexit__(self, *args):
        try:
            await self.conn.close()
            await self.conn.logout()
        except:
            pass

    async def fetch_emails(self) -> AsyncGenerator[tuple[str, Message], None]:
        try:
            from main import parse_subject, parse_subject_pattern
            start_date = os.environ.get("START_DATE", "01-Mar-2025")

            # ✅ 修复：SEARCH 不能用 UID
            status, data = await self.conn.search('SINCE', start_date)
            if status != 'OK' or not data[0]:
                logger.info("未找到邮件")
                return

            msg_ids = data[0].split()
            for msg_id in msg_ids:
                status, data = await self.conn.fetch(msg_id, '(RFC822)')
                if status != 'OK':
                    continue
                msg_bytes = data[1]
                msg = message_from_bytes(msg_bytes)
                yield msg_id, msg
        except Exception as e:
            logger.error(f"抓取邮件异常: {e}")
