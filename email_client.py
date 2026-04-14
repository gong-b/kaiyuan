import imaplib
import ssl
import logging
from typing import Generator, Tuple, Optional
from email import message_from_bytes
from email.message import Message
from config import IMAP_HOST, IMAP_PORT, EMAIL_USER, EMAIL_PASSWORD

logger = logging.getLogger(__name__)

class SecureIMAPClient:
    """修复编码问题的IMAP客户端"""
    host: str
    port: int
    user: str
    password: str
    mailbox: str
    conn: Optional[imaplib.IMAP4_SSL]

    def __init__(self) -> None:
        # 直接从config读取（已兼容环境变量），无需global声明
        self.host = IMAP_HOST
        self.port = IMAP_PORT
        self.user = EMAIL_USER
        self.password = EMAIL_PASSWORD
        self.mailbox = "其他文件夹/开源课堂"  # <--- 核心修改点：切换到开源课堂文件夹
        self.conn = None

    def __enter__(self) -> "SecureIMAPClient":
        context: ssl.SSLContext = ssl.create_default_context()
        context.check_hostname = True
        context.verify_mode = ssl.CERT_REQUIRED
        try:
            # 强制使用UTF-8编码
            self.conn = imaplib.IMAP4_SSL(
                self.host, self.port, ssl_context=context
            )
            logger.debug(f"SSL连接建立成功")
            self.conn.login(self.user, self.password)
            status, _ = self.conn.select(self.mailbox)
            if status != "OK":
                logger.error(f"邮箱文件夹选择失败: {self.mailbox}")
                raise RuntimeError("邮箱文件夹不可用")
            else:
                logger.info(f"邮箱文件夹 [{self.mailbox}] 登录选择成功")
        except Exception as e:
            logger.error(f"邮箱连接失败: {str(e).encode('utf-8', errors='replace').decode('utf-8')}")
            raise
        return self

    def __exit__(self, exc_type, exc_value, traceback) -> None:
        if exc_type:
            logger.error(f"上下文管理器中发生异常: {exc_value}", exc_info=(exc_type, exc_value, traceback))
        if self.conn:
            try:
                if self.conn.state == "SELECTED":
                    self.conn.close()
                if self.conn.state == "AUTH":
                    self.conn.logout()
                logger.info("IMAP连接已安全关闭")
            except Exception as e:
                logger.error(f"关闭IMAP连接时发生错误: {e}")
            finally:
                self.conn = None

    def fetch_emails(self) -> Generator[Tuple[str, Message], None, None]:
        """安全获取邮件内容（修复版）"""
        try:
            from typing import Any
            import os

            if self.conn is None:
                logger.error("IMAP连接未初始化")
                return

            # 修复：从环境变量读取日期，兼容原有固定日期
            start_date = os.environ.get("START_DATE", "01-Mar-2025")

            # 使用环境变量的日期搜索邮件
            status: str
            data: list[bytes]
            status, data = self.conn.uid('SEARCH', 'SINCE', start_date)
            if status != 'OK':
                logger.error(f"邮件搜索失败: {data[0].decode('utf-8', errors='replace')}")
                return

            # 拆分UID列表（关键修复）
            uids: list[bytes] = data[0].split()
            if not uids:
                logger.info("未找到匹配邮件")
                return

            for uid_bytes in uids:
                uid: str = uid_bytes.decode('utf-8')
                try:
                    status: str
                    msg_data: list[Tuple[Any, Any]]
                    status, msg_data = self.conn.uid('FETCH', uid, '(RFC822)')
                    if status != 'OK':
                        raw_error = msg_data[0][1] if msg_data and isinstance(msg_data[0][1], bytes) else str(msg_data)
                        if isinstance(raw_error, bytes):
                            raw_error = raw_error.decode('utf-8', errors='replace')
                        logger.error(f"邮件获取失败: {uid} - {raw_error}")
                        continue

                    raw_bytes: list[Tuple[Any, Any]] = msg_data
                    
                    if raw_bytes and isinstance(raw_bytes[0][1], bytes):
                        msg: Message = message_from_bytes(raw_bytes[0][1])
                        # 在Python层过滤：检查主题是否包含"书法班"
                        from main import parse_subject
                        subject = parse_subject(msg)
                        if "书法班" in subject:
                            yield uid, msg
                    else:
                        logger.error(f"邮件原始数据格式异常: {uid} - {raw_bytes}")
                except Exception as e:
                    error_msg: str = str(e).encode('utf-8', errors='replace').decode('utf-8')
                    logger.error(f"邮件处理失败: {uid} - {error_msg}")
        except Exception as e:
            error_msg: str = str(e).encode('utf-8', errors='replace').decode('utf-8')
            logger.error(f"获取邮件失败: {error_msg}")
