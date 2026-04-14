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
    """
    修正后的 IMAP Modified UTF-7 编码逻辑
    用于处理中文文件夹名称
    """
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
        # 建议尝试方案 A: "开源课堂" (浙大邮箱常见情况)
        # 建议尝试方案 B: "其他文件夹/开源课堂"
        self.mailbox_raw = "开源课堂" 
        self.conn = None

    def __enter__(self) -> "SecureIMAPClient":
        context = ssl.create_default_context()
        try:
            self.conn = imaplib.IMAP4_SSL(self.host, self.port, ssl_context=context)
            self.conn.login(self.user, self.password)
            
            # 编码路径
            encoded_mailbox = imap_utf7_encode(self.mailbox_raw)
            status, _ = self.conn.select(encoded_mailbox)
            
            if status != "OK":
                logger.error(f"无法进入文件夹: {self.mailbox_raw} (编码后: {encoded_mailbox})")
                
                # --- 诊断逻辑：列出所有文件夹 ---
                logger.info("正在获取邮箱内所有文件夹列表，请在下方日志中查看正确名称：")
                typ, folders = self.conn.list()
                if typ == 'OK':
                    for f in folders:
                        # 尝试解码文件夹名以便阅读
                        try:
                            f_str = f.decode('ascii')
                            logger.info(f"发现文件夹 -> {f_str}")
                        except:
                            logger.info(f"发现文件夹 (原始数据) -> {f}")
                # ------------------------------
                
                raise RuntimeError(f"文件夹 {self.mailbox_raw} 不存在")
            
            logger.info(f"成功进入文件夹: {self.mailbox_raw}")
        except Exception as e:
            logger.error(f"IMAP连接错误: {e}")
            raise
        return self

    def __exit__(self, exc_type, exc_value, traceback) -> None:
        if self.conn:
            try:
                if self.conn.state == "SELECTED":
                    self.conn.close()
                self.conn.logout()
            except:
                pass

    def fetch_emails(self) -> Generator[Tuple[str, Message], None, None]:
        try:
            from main import parse_subject, parse_subject_pattern
            start_date = os.environ.get("START_DATE", "01-Mar-2025")
            # 搜索日期后的邮件
            status, data = self.conn.uid('SEARCH', 'SINCE', start_date)
            
            if status != 'OK' or not data[0]:
                logger.info("当前文件夹内未找到符合日期要求的邮件")
                return

            uids = data[0].split()
            for uid_bytes in uids:
                uid = uid_bytes.decode('utf-8')
                # 先抓取头信息，减少流量
                status, header_data = self.conn.uid('FETCH', uid, '(BODY.PEEK[HEADER.FIELDS (SUBJECT DATE)])')
                if status != 'OK': continue
                
                header_msg = message_from_bytes(header_data[0][1])
                subject = parse_subject(header_msg)
                
                name, sid = parse_subject_pattern(subject)
                if not (name and sid): continue

                # 匹配成功，抓取全文
                status, full_data = self.conn.uid('FETCH', uid, '(RFC822)')
                if status == 'OK':
                    yield uid, message_from_bytes(full_data[0][1])
                    
        except Exception as e:
            logger.error(f"抓取邮件异常: {e}")
