import imaplib
import ssl
import logging
import os
import base64
from typing import Generator, Tuple, Optional
from email import message_from_bytes
from email.message import Message
from config import IMAP_HOST, IMAP_PORT, EMAIL_USER, EMAIL_PASSWORD
import asyncio
import aioimaplib  # 需安装：pip install aioimaplib

logger = logging.getLogger(__name__)

def imap_utf7_encode(text):
    """原有编码逻辑保留"""
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
        self.conn = None

    async def __aenter__(self) -> "AsyncSecureIMAPClient":
        context = ssl.create_default_context()
        try:
            self.conn = aioimaplib.IMAP4_SSL(host=self.host, port=self.port, ssl_context=context)
            await self.conn.wait_hello_from_server()
            await self.conn.login(self.user, self.password)
            
            encoded_mailbox = imap_utf7_encode(self.mailbox_raw)
            status, _ = await self.conn.select(encoded_mailbox)
            
            if status != "OK":
                logger.error(f"无法进入文件夹: {self.mailbox_raw}")
                typ, folders = await self.conn.list()
                if typ == 'OK':
                    for f in folders:
                        try:
                            logger.info(f"发现文件夹 -> {f.decode('ascii')}")
                        except:
                            logger.info(f"发现文件夹 (原始) -> {f}")
                raise RuntimeError(f"文件夹 {self.mailbox_raw} 不存在")
            
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

    async def fetch_emails_batch(self, batch_size: int = 50) -> Generator[Tuple[str, Message], None, None]:
        """批量抓取邮件，减少IO次数"""
        try:
            start_date = os.environ.get("START_DATE", "01-Mar-2025")
            status, data = await self.conn.uid('SEARCH', 'SINCE', start_date)
            
            if status != 'OK' or not data[0]:
                logger.info("未找到符合日期的邮件")
                return

            uids = data[0].split()
            # 分批处理UID
            for i in range(0, len(uids), batch_size):
                batch_uids = uids[i:i+batch_size]
                # 批量抓取头信息
                uid_str = ','.join([uid.decode('utf-8') for uid in batch_uids])
                status, header_data = await self.conn.uid('FETCH', uid_str, '(BODY.PEEK[HEADER.FIELDS (SUBJECT DATE)])')
                
                if status != 'OK': continue
                # 解析头信息过滤
                from main import parse_subject, parse_subject_pattern
                valid_uids = []
                for idx, uid_bytes in enumerate(batch_uids):
                    uid = uid_bytes.decode('utf-8')
                    # 匹配头信息位置
                    header_part = header_data[idx*2] if len(header_data) > idx*2 else None
                    if not header_part: continue
                    header_msg = message_from_bytes(header_part[1])
                    subject = parse_subject(header_msg)
                    name, sid = parse_subject_pattern(subject)
                    if name and sid:
                        valid_uids.append(uid)
                
                # 批量抓取有效邮件全文
                if valid_uids:
                    valid_uid_str = ','.join(valid_uids)
                    status, full_data = await self.conn.uid('FETCH', valid_uid_str, '(RFC822)')
                    if status == 'OK':
                        for item in full_data:
                            if isinstance(item, tuple) and len(item) >= 2:
                                yield item[0].split()[2].decode('utf-8'), message_from_bytes(item[1])
        except Exception as e:
            logger.error(f"批量抓取邮件异常: {e}")
