import imaplib
import ssl
import logging
import os
from typing import Generator, Tuple, Optional
from email import message_from_bytes
from email.message import Message
from config import IMAP_HOST, IMAP_PORT, EMAIL_USER, EMAIL_PASSWORD

logger = logging.getLogger(__name__)

class SecureIMAPClient:
    def __init__(self) -> None:
        self.host = IMAP_HOST
        self.port = IMAP_PORT
        self.user = EMAIL_USER
        self.password = EMAIL_PASSWORD
        self.mailbox = "其他文件夹/开源课堂" 
        self.conn = None

    def __enter__(self) -> "SecureIMAPClient":
        context = ssl.create_default_context()
        try:
            self.conn = imaplib.IMAP4_SSL(self.host, self.port, ssl_context=context)
            self.conn.login(self.user, self.password)
            status, _ = self.conn.select(self.mailbox)
            if status != "OK":
                # 尝试 Modified UTF-7 编码或其他可能的路径写法（如果原始字符串失败）
                logger.error(f"无法进入文件夹: {self.mailbox}")
                raise RuntimeError(f"文件夹 {self.mailbox} 不存在或不可访问")
            logger.info(f"成功进入文件夹: {self.mailbox}")
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
        """获取符合条件的邮件"""
        try:
            start_date = os.environ.get("START_DATE", "01-Mar-2025")
            status, data = self.conn.uid('SEARCH', 'SINCE', start_date)
            
            if status != 'OK' or not data[0]:
                logger.info("未找到规定日期后的邮件")
                return

            uids = data[0].split()
            for uid_bytes in uids:
                uid = uid_bytes.decode('utf-8')
                status, msg_data = self.conn.uid('FETCH', uid, '(RFC822)')
                if status != 'OK': continue

                msg = message_from_bytes(msg_data[0][1])
                
                # 预读主题进行初步过滤
                from main import parse_subject
                subject = parse_subject(msg)
                
                # 放宽条件：只要包含报名或申请关键字就处理
                if any(kw in subject for kw in ["报名", "申请", "班"]):
                    yield uid, msg
                    
        except Exception as e:
            logger.error(f"邮件抓取异常: {e}")
