import imaplib
import ssl
import logging
import os
import base64
from typing import Generator, Tuple, Optional
from email import message_from_bytes
from email.message import Message
from config import IMAP_HOST, IMAP_PORT, EMAIL_USER, EMAIL_PASSWORD

logger = logging.getLogger(__name__)

def imap_utf7_encode(text):
    """IMAP Modified UTF-7 编码，用于处理中文文件夹名称"""
    def _modified_base64(s):
        s_utf16 = s.encode('utf-16-be')
        return base64.b64encode(s_utf16).decode('ascii').rstrip('=').replace('/', ',')
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
        self.user = EMAIL_USER
        self.password = EMAIL_PASSWORD
        self.mailbox_raw = "其他文件夹/开源课堂" 
        self.conn = None

    def __enter__(self) -> "SecureIMAPClient":
        context = ssl.create_default_context()
        try:
            self.conn = imaplib.IMAP4_SSL(self.host, self.port, ssl_context=context)
            self.conn.login(self.user, self.password)
            encoded_mailbox = imap_utf7_encode(self.mailbox_raw)
            status, _ = self.conn.select(encoded_mailbox)
            if status != "OK":
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
            except: pass

    def fetch_emails(self) -> Generator[Tuple[str, Message], None, None]:
        """两阶段抓取：先扫标题过滤，再下全文附件"""
        try:
            from main import parse_subject, parse_subject_pattern
            start_date = os.environ.get("START_DATE", "01-Mar-2025")
            status, data = self.conn.uid('SEARCH', 'SINCE', start_date)
            if status != 'OK' or not data[0]: return

            uids = data[0].split()
            logger.info(f"发现 {len(uids)} 封潜在邮件，正在快速扫描标题...")

            for uid_bytes in uids:
                uid = uid_bytes.decode('utf-8')
                # 阶段 1: 只下载邮件头
                status, header_data = self.conn.uid('FETCH', uid, '(BODY.PEEK[HEADER.FIELDS (SUBJECT DATE)])')
                if status != 'OK': continue
                header_msg = message_from_bytes(header_data[0][1])
                subject = parse_subject(header_msg)
                
                # 预筛选：如果标题解析不出任何信息，直接跳过，不下载附件
                name, sid = parse_subject_pattern(subject)
                if not (name and sid): continue

                # 阶段 2: 匹配成功，下载包含附件的全文
                logger.info(f"匹配成功，正在提取详细内容: {subject}")
                status, full_data = self.conn.uid('FETCH', uid, '(RFC822)')
                if status == 'OK':
                    yield uid, message_from_bytes(full_data[0][1])
        except Exception as e:
            logger.error(f"提取异常: {e}")
