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
        self.host = IMAP_HOST
        self.port = IMAP_PORT
        self.user = EMAIL_USER
        self.password = EMAIL_PASSWORD
        self.mailbox = "INBOX"
        self.conn = None

    def __enter__(self) -> "SecureIMAPClient":
        # 适配 Streamlit 环境的 SSL 配置
        context: ssl.SSLContext = ssl.create_default_context()
        context.check_hostname = False  # 兼容部分邮件服务器
        context.verify_mode = ssl.CERT_REQUIRED if self.host != "localhost" else ssl.CERT_NONE
        
        try:
            # 强制使用UTF-8编码
            self.conn = imaplib.IMAP4_SSL(
                self.host, 
                self.port, 
                ssl_context=context
            )
            logger.debug(f"SSL连接建立成功: {self.host}:{self.port}")
            
            # 登录校验（空密码跳过，适配开发环境）
            if self.password:
                self.conn.login(self.user, self.password)
            else:
                logger.warning("邮箱密码未配置，跳过登录")
            
            status, _ = self.conn.select(self.mailbox)
            if status != "OK":
                logger.error("邮箱选择失败")
                raise RuntimeError("邮箱不可用")
            else:
                logger.info("邮箱登录/选择成功")
        except Exception as e:
            error_msg = str(e).encode('utf-8', errors='replace').decode('utf-8')
            logger.error(f"邮箱连接失败: {error_msg}")
            raise
        return self

    def __exit__(self, exc_type, exc_value, traceback) -> None:
        if exc_type:
            logger.error(f"上下文管理器异常: {exc_value}", exc_info=(exc_type, exc_value, traceback))
        
        if self.conn:
            try:
                if self.conn.state == "SELECTED":
                    self.conn.close()
                if self.conn.state == "AUTH":
                    self.conn.logout()
                logger.info("IMAP连接已安全关闭")
            except Exception as e:
                logger.error(f"关闭IMAP连接错误: {str(e)}")
            finally:
                self.conn = None

    def fetch_emails(self) -> Generator[Tuple[str, Message], None, None]:
        """安全获取邮件内容（修复版）"""
        try:
            if self.conn is None:
                logger.error("IMAP连接未初始化")
                return
            
            # 修复搜索条件：兼容不同IMAP服务器的日期格式
            status, data = self.conn.uid('SEARCH', None, 'SINCE', '01-Mar-2025')
            
            if status != 'OK':
                err_msg = data[0].decode('utf-8', errors='replace') if data else "未知错误"
                logger.error(f"邮件搜索失败: {err_msg}")
                return

            uids = data[0].split() if data and data[0] else []
            if not uids:
                logger.info("未找到匹配邮件")
                return

            for uid_bytes in uids:
                uid = uid_bytes.decode('utf-8').strip()
                if not uid:
                    continue
                    
                try:
                    # 获取邮件内容（严格类型处理）
                    status, msg_data = self.conn.uid('FETCH', uid, '(RFC822)')
                    if status != 'OK':
                        logger.error(f"邮件获取失败: {uid} - {str(msg_data)}")
                        continue

                    # 修复邮件原始数据提取逻辑
                    if isinstance(msg_data, list) and len(msg_data) >= 1:
                        raw_bytes = msg_data[0][1] if isinstance(msg_data[0], tuple) else None
                        if isinstance(raw_bytes, bytes):
                            msg = message_from_bytes(raw_bytes)
                            # 过滤包含"书法班"的邮件
                            from main import parse_subject
                            subject = parse_subject(msg)
                            if "书法班" in subject:
                                yield uid, msg
                        else:
                            logger.error(f"邮件数据格式异常: {uid} - 非字节数据")
                    else:
                        logger.error(f"邮件数据结构异常: {uid} - {str(msg_data)}")
                        
                except Exception as e:
                    error_msg = str(e).encode('utf-8', errors='replace').decode('utf-8')
                    logger.error(f"邮件处理失败: {uid} - {error_msg}")
                    
        except Exception as e:
            error_msg = str(e).encode('utf-8', errors='replace').decode('utf-8')
            logger.error(f"获取邮件失败: {error_msg}")
