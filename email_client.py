import imaplib
import ssl
import logging
import os
from datetime import datetime, timezone
from typing import Generator, Tuple, Optional
from email import message_from_bytes
from email.message import Message
from email.header import decode_header
from config import IMAP_HOST, IMAP_PORT, EMAIL_USER, EMAIL_PASSWORD

logger = logging.getLogger(__name__)

class SecureIMAPClient:
    """IMAP客户端（增强异常处理）"""
    def __init__(self) -> None:
        self.host = IMAP_HOST
        self.port = IMAP_PORT
        self.user = EMAIL_USER
        self.password = EMAIL_PASSWORD
        self.mailbox = "INBOX"
        self.conn: Optional[imaplib.IMAP4_SSL] = None

    def __enter__(self) -> "SecureIMAPClient":
        """上下文进入：精准捕获IMAP连接/登录异常"""
        # 前置校验（非空）
        required_config = [self.user, self.password, self.host]
        if not all(required_config):
            raise ValueError(
                "IMAP配置不完整：\n"
                f"- 邮箱账号: {'已配置' if self.user else '缺失'}\n"
                f"- 客户端密码: {'已配置' if self.password else '缺失'}\n"
                f"- IMAP主机: {self.host}"
            )
        
        # 构建SSL上下文
        context: ssl.SSLContext = ssl.create_default_context()
        context.check_hostname = False
        context.verify_mode = ssl.CERT_NONE
        
        try:
            # 连接IMAP服务器（捕获连接异常）
            self.conn = imaplib.IMAP4_SSL(
                self.host, 
                self.port, 
                ssl_context=context,
                timeout=15  # 连接超时15秒
            )
            logger.debug(f"SSL连接建立成功: {self.host}:{self.port}")
            
            # 登录（捕获登录异常）
            try:
                self.conn.login(self.user.encode('utf-8'), self.password.encode('utf-8'))
            except imaplib.IMAP4.error as e:
                error_msg = str(e).strip()
                # 针对性提示
                if "authentication failed" in error_msg.lower():
                    raise RuntimeError(
                        "邮箱登录失败！请检查：\n"
                        "1. 客户端专用密码是否正确（非登录密码）\n"
                        "2. 浙大邮箱是否开启IMAP/SMTP服务\n"
                        "3. 账号是否有IMAP访问权限"
                    ) from e
                else:
                    raise RuntimeError(f"IMAP登录错误: {error_msg}") from e
            
            # 选择邮箱文件夹（捕获文件夹异常）
            try:
                status, _ = self.conn.select(self.mailbox, readonly=True)
                if status != "OK":
                    raise RuntimeError(f"选择邮箱文件夹失败: {self.mailbox}")
            except imaplib.IMAP4.error as e:
                raise RuntimeError(f"访问邮箱文件夹失败: {str(e)}") from e
            
            logger.info(f"成功登录邮箱: {self.user}")
            return self
        
        # 捕获连接类异常
        except TimeoutError:
            raise RuntimeError(f"连接IMAP服务器超时（{self.host}:{self.port}），请检查网络/服务器状态")
        except ConnectionRefusedError:
            raise RuntimeError(f"IMAP服务器拒绝连接（{self.host}:{self.port}），请检查主机/端口配置")
        except Exception as e:
            raise RuntimeError(f"邮箱连接初始化失败: {str(e)}") from e

    def __exit__(self, exc_type, exc_value, traceback) -> None:
        """上下文退出：安全关闭连接（忽略退出时的小错误）"""
        if self.conn:
            try:
                if self.conn.state == "SELECTED":
                    self.conn.close()
                self.conn.logout()
                logger.info("IMAP连接已安全关闭")
            except Exception as e:
                # 退出时的错误不影响主流程，仅日志记录
                logger.warning(f"关闭IMAP连接时警告: {str(e)}")
            finally:
                self.conn = None

    def fetch_emails(self) -> Generator[Tuple[str, Message], None, None]:
        """获取邮件：分步骤捕获异常，出错跳过单封邮件"""
        if self.conn is None:
            logger.error("IMAP连接未初始化，无法获取邮件")
            return

        # 1. 解析日期范围（捕获日期格式异常）
        try:
            start_date_str = os.environ.get("START_DATE", "01-Mar-2025")
            end_date_str = os.environ.get("END_DATE", datetime.now().strftime("%d-%b-%Y"))
            # 验证日期格式（IMAP要求：dd-Mon-yyyy）
            datetime.strptime(start_date_str, "%d-%b-%Y")
            datetime.strptime(end_date_str, "%d-%b-%Y")
        except ValueError as e:
            logger.error(f"日期格式错误（需dd-Mon-yyyy，如01-Mar-2025）: {e}")
            return

        # 2. 搜索邮件（捕获搜索异常）
        try:
            search_criteria = f'(SINCE "{start_date_str}" BEFORE "{end_date_str}")'
            logger.info(f"搜索邮件条件: {search_criteria}")
            status, data = self.conn.uid('SEARCH', None, search_criteria)
            
            if status != 'OK':
                error_msg = data[0].decode('utf-8', errors='replace') if data else "未知错误"
                logger.error(f"邮件搜索失败: {error_msg}")
                return

            uids = data[0].split() if data and data[0] else []
            if not uids:
                logger.info("未找到符合日期条件的邮件")
                return
            logger.info(f"找到{len(uids)}封待处理邮件")
        
        except imaplib.IMAP4.error as e:
            logger.error(f"IMAP搜索邮件异常: {str(e)}", exc_info=True)
            return
        except Exception as e:
            logger.error(f"邮件搜索流程异常: {str(e)}", exc_info=True)
            return

        # 3. 遍历邮件（单封邮件出错不终止，仅跳过）
        for uid_bytes in uids:
            uid = uid_bytes.decode('utf-8').strip()
            if not uid:
                logger.warning("空UID，跳过")
                continue
                
            try:
                # 获取邮件原始内容
                status, msg_data = self.conn.uid('FETCH', uid, '(RFC822)')
                if status != 'OK':
                    logger.error(f"获取邮件{uid}失败: {msg_data}")
                    continue

                # 解析邮件数据
                raw_email = None
                for response_part in msg_data:
                    if isinstance(response_part, tuple) and len(response_part) >= 2:
                        raw_email = response_part[1]
                        break
                
                if not isinstance(raw_email, bytes):
                    logger.error(f"邮件{uid}数据格式异常（非字节类型）: {type(raw_email)}")
                    continue

                # 解析为Message对象
                msg = message_from_bytes(raw_email)
                
                # 过滤书法班邮件
                subject = self._get_msg_subject(msg)
                if "书法班" in subject:
                    yield uid, msg
                else:
                    logger.debug(f"邮件{uid}非书法班报名（主题：{subject[:50]}...），跳过")
                    
            except Exception as e:
                # 单封邮件出错，记录日志后继续处理下一封
                logger.error(f"处理邮件{uid}失败（已跳过）: {str(e)}", exc_info=True)
                continue

    @staticmethod
    def _get_msg_subject(msg: Message) -> str:
        """获取主题：捕获解码异常"""
        subject = msg.get("Subject", "")
        decoded_parts = []
        for part, charset in decode_header(subject):
            try:
                if isinstance(part, bytes):
                    # 优先尝试常见编码
                    for encoding in [charset, 'utf-8', 'gb18030', 'big5', 'gbk']:
                        if not encoding:
                            continue
                        try:
                            decoded = part.decode(encoding)
                            break
                        except:
                            continue
                    else:
                        decoded = part.decode('utf-8', errors='replace')
                else:
                    decoded = str(part)
                decoded_parts.append(decoded)
            except Exception as e:
                logger.warning(f"主题片段解码失败: {e}")
                decoded_parts.append("[解码失败]")
        return "".join(decoded_parts)
