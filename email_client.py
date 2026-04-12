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
    def __init__(self) -> None:
        self.host = IMAP_HOST
        self.port = IMAP_PORT
        self.user = EMAIL_USER
        self.password = EMAIL_PASSWORD
        self.mailbox = "INBOX"
        self.conn: Optional[imaplib.IMAP4_SSL] = None

    def __enter__(self) -> "SecureIMAPClient":
        # 验证必填参数
        if not all([self.user, self.password, self.host]):
            raise ValueError("IMAP配置不完整（用户/密码/主机不能为空）")
            
        context: ssl.SSLContext = ssl.create_default_context()
        # 兼容部分邮件服务器的证书问题（生产环境建议保留验证）
        context.check_hostname = False
        context.verify_mode = ssl.CERT_NONE
        
        try:
            self.conn = imaplib.IMAP4_SSL(
                self.host, 
                self.port, 
                ssl_context=context
            )
            logger.debug(f"SSL连接建立成功: {self.host}:{self.port}")
            
            # 登录（处理编码问题）
            self.conn.login(self.user.encode('utf-8'), self.password.encode('utf-8'))
            
            # 选择邮箱
            status, _ = self.conn.select(self.mailbox, readonly=True)  # 只读模式避免误操作
            if status != "OK":
                logger.error("邮箱选择失败")
                raise RuntimeError("邮箱文件夹选择失败")
            
            logger.info(f"成功登录邮箱: {self.user}")
        except Exception as e:
            error_msg = str(e).encode('utf-8', errors='replace').decode('utf-8')
            logger.error(f"邮箱连接/登录失败: {error_msg}")
            raise
        return self

    def __exit__(self, exc_type, exc_value, traceback) -> None:
        if exc_type:
            logger.error(f"上下文管理器异常: {exc_value}", exc_info=True)
        
        if self.conn:
            try:
                if self.conn.state == "SELECTED":
                    self.conn.close()
                self.conn.logout()
                logger.info("IMAP连接已安全关闭")
            except Exception as e:
                logger.error(f"关闭IMAP连接失败: {str(e)}")
            finally:
                self.conn = None

    def fetch_emails(self) -> Generator[Tuple[str, Message], None, None]:
        """安全获取邮件内容（修复日期筛选和数据解析）"""
        if self.conn is None:
            logger.error("IMAP连接未初始化")
            return

        try:
            # 读取环境变量中的日期范围
            start_date = os.environ.get("START_DATE", "01-Mar-2025")
            end_date = os.environ.get("END_DATE", datetime.now().strftime("%d-%b-%Y"))
            
            # 构建IMAP搜索条件（SINCE + BEFORE）
            search_criteria = f'(SINCE "{start_date}" BEFORE "{end_date}")'
            logger.info(f"搜索邮件条件: {search_criteria}")
            
            # 执行UID搜索
            status, data = self.conn.uid('SEARCH', None, search_criteria)
            if status != 'OK':
                error_msg = data[0].decode('utf-8', errors='replace') if data else "未知错误"
                logger.error(f"邮件搜索失败: {error_msg}")
                return

            uids = data[0].split() if data and data[0] else []
            if not uids:
                logger.info("未找到符合条件的邮件")
                return

            logger.info(f"找到{len(uids)}封待处理邮件")
            
            for uid_bytes in uids:
                uid = uid_bytes.decode('utf-8').strip()
                if not uid:
                    continue
                    
                try:
                    # 获取邮件原始内容
                    status, msg_data = self.conn.uid('FETCH', uid, '(RFC822)')
                    if status != 'OK':
                        logger.error(f"获取邮件{uid}失败: {msg_data}")
                        continue

                    # 解析邮件数据（处理多种返回格式）
                    raw_email = None
                    for response_part in msg_data:
                        if isinstance(response_part, tuple) and len(response_part) >= 2:
                            raw_email = response_part[1]
                            break
                    
                    if not isinstance(raw_email, bytes):
                        logger.error(f"邮件{uid}数据格式异常: {type(raw_email)}")
                        continue

                    # 解析为Message对象
                    msg = message_from_bytes(raw_email)
                    
                    # 过滤包含"书法班"的邮件（主题/内容）
                    subject = self._get_msg_subject(msg)
                    if "书法班" in subject:
                        yield uid, msg
                    else:
                        logger.debug(f"邮件{uid}非书法班报名，跳过")
                        
                except Exception as e:
                    error_msg = str(e).encode('utf-8', errors='replace').decode('utf-8')
                    logger.error(f"处理邮件{uid}失败: {error_msg}")
                    
        except Exception as e:
            error_msg = str(e).encode('utf-8', errors='replace').decode('utf-8')
            logger.error(f"获取邮件列表失败: {error_msg}")

    @staticmethod
    def _get_msg_subject(msg: Message) -> str:
        """获取邮件主题（简化版，避免循环依赖）"""
        subject = msg.get("Subject", "")
        decoded_parts = []
        for part, charset in decode_header(subject):
            if isinstance(part, bytes):
                try:
                    decoded_parts.append(part.decode(charset or 'utf-8', errors='replace'))
                except:
                    decoded_parts.append(part.decode('gb18030', errors='replace'))
            else:
                decoded_parts.append(str(part))
        return "".join(decoded_parts)
