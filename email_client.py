import asyncio
import aioimaplib
import ssl
import logging
import os
import base64
from typing import Generator, Tuple, Optional, AsyncGenerator
from email import message_from_bytes
from email.message import Message
from config import IMAP_HOST, IMAP_PORT, EMAIL_USER, EMAIL_PASSWORD

logger = logging.getLogger(__name__)

def imap_utf7_encode(text):
    """修正后的 IMAP Modified UTF-7 编码逻辑"""
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
            
            # 编码并选择文件夹
            encoded_mailbox = imap_utf7_encode(self.mailbox_raw)
            status, _ = await self.conn.select(encoded_mailbox)
            
            if status != "OK":
                logger.error(f"无法进入文件夹: {self.mailbox_raw} (编码后: {encoded_mailbox})")
                # 列出所有文件夹用于诊断
                typ, folders = await self.conn.list()
                if typ == 'OK':
                    for f in folders:
                        try:
                            f_str = f.decode('ascii')
                            logger.info(f"发现文件夹 -> {f_str}")
                        except:
                            logger.info(f"发现文件夹 (原始数据) -> {f}")
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

    async def fetch_emails(self) -> AsyncGenerator[Tuple[str, Message], None]:
        try:
            from main import parse_subject, parse_subject_pattern
            start_date = os.environ.get("START_DATE", "01-Mar-2025")
            
            # 批量搜索邮件UID
            status, data = await self.conn.uid('SEARCH', 'SINCE', start_date)
            if status != 'OK' or not data[0]:
                logger.info("当前文件夹内未找到符合日期要求的邮件")
                return

            uids = data[0].split()
            if not uids:
                return

            # 批量获取邮件头（减少请求次数）
            uid_list = b' '.join(uids).decode()
            status, header_data = await self.conn.uid(
                'FETCH', 
                uid_list, 
                '(BODY.PEEK[HEADER.FIELDS (SUBJECT DATE)])'
            )
            
            if status != 'OK':
                logger.error("批量获取邮件头失败")
                return

            # 解析邮件头，筛选符合条件的UID
            valid_uids = []
            header_lines = header_data[:-1]  # 排除最后的 'OK'
            for i in range(0, len(header_lines), 2):
                if i+1 >= len(header_lines):
                    continue
                uid_line = header_lines[i].decode()
                header_bytes = header_lines[i+1]
                
                # 提取UID
                uid_match = re.search(r'UID (\d+)', uid_line)
                if not uid_match:
                    continue
                uid = uid_match.group(1)
                
                # 解析主题
                header_msg = message_from_bytes(header_bytes)
                subject = parse_subject(header_msg)
                name, sid = parse_subject_pattern(subject)
                
                if name and sid:
                    valid_uids.append(uid)

            # 批量获取符合条件的邮件全文
            if valid_uids:
                valid_uid_list = ' '.join(valid_uids)
                status, full_data = await self.conn.uid(
                    'FETCH', 
                    valid_uid_list, 
                    '(RFC822)'
                )
                
                if status == 'OK':
                    # 解析批量返回的邮件内容
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
