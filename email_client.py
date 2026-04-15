import asyncio
import aioimaplib
import ssl
import logging
import os
import base64
import re  # 这里补上缺失的re！
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

class AsyncSecureIMAPClient:
    def __init__(self) -> None:
        self.host = IMAP_HOST
        self.port = IMAP_PORT
        self.user = EMAIL_USER
        self.password = EMAIL_PASSWORD
        self.mailbox_raw = "开源课堂"
        self.conn: Optional[aioimaplib.IMAP4_SSL] = None
        self.ssl_context = ssl.create_default_context()

    async def __aenter__(self) -> "AsyncSecureIMAPClient":
        try:
            self.conn = aioimaplib.IMAP4_SSL(
                host=self.host,
                port=self.port,
                ssl_context=self.ssl_context
            )
            await self.conn.wait_hello_from_server()
            await self.conn.login(self.user, self.password)
            
            encoded_mailbox = imap_utf7_encode(self.mailbox_raw)
            status, _ = await self.conn.select(encoded_mailbox)
            
            if status != "OK":
                logger.error(f"无法进入文件夹: {self.mailbox_raw}")
                raise RuntimeError(f"文件夹不存在")
            
            logger.info(f"成功进入文件夹: {self.mailbox_raw}")
        except Exception as e:
            logger.error(f"IMAP连接错误: {e}")
            raise
        return self

    async def __aexit__(self, exc_type, exc_value, traceback) -> None:
        if self.conn:
            try:
                if self.conn.state == "SELECTED":
                    await self.conn.close()
                await self.conn.logout()
            except:
                pass

    async def fetch_emails(self) -> AsyncGenerator[tuple[str, Message], None]:
        try:
            from main import parse_subject, parse_subject_pattern
            start_date = os.environ.get("START_DATE", "01-Mar-2025")
            
            status, data = await self.conn.uid('SEARCH', 'SINCE', start_date)
            if status != 'OK' or not data[0]:
                logger.info("未找到邮件")
                return

            uids = data[0].split()
            if not uids:
                return

            uid_list = b' '.join(uids).decode()
            status, header_data = await self.conn.uid(
                'FETCH', uid_list, '(BODY.PEEK[HEADER.FIELDS (SUBJECT DATE)])'
            )
            
            if status != 'OK':
                return

            valid_uids = []
            header_lines = header_data[:-1]
            for i in range(0, len(header_lines), 2):
                if i+1 >= len(header_lines):
                    continue
                uid_line = header_lines[i].decode()
                header_bytes = header_lines[i+1]
                
                uid_match = re.search(r'UID (\d+)', uid_line)
                if not uid_match:
                    continue
                uid = uid_match.group(1)
                
                header_msg = message_from_bytes(header_bytes)
                subject = parse_subject(header_msg)
                name, sid = parse_subject_pattern(subject)
                
                if name and sid:
                    valid_uids.append(uid)

            if valid_uids:
                valid_uid_list = ' '.join(valid_uids)
                status, full_data = await self.conn.uid(
                    'FETCH', valid_uid_list, '(RFC822)'
                )
                
                if status == 'OK':
                    full_lines = full_data[:-1]
                    for i in range(0, len(full_lines), 2):
                        if i+1 >= len(full_lines):
                            continue
                        uid_line = full_lines[i].decode()
                        full_bytes = full_lines[i+1]
                        
                        uid_match = re.search(r'UID (\d+)', uid_line)
                        if uid_match:
                            uid = uid_match.group(1)
                            yield uid, message_from_bytes(full_bytes)
                        
        except Exception as e:
            logger.error(f"抓取邮件异常: {e}")
