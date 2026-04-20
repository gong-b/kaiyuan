import re
import logging
from datetime import datetime
from email.header import decode_header
from email.message import Message
from email.utils import parsedate_to_datetime
from .config import Config

logger = logging.getLogger(__name__)

class EmailParser:
    """邮件解析核心类"""
    def __init__(self):
        self.subject_re = re.compile(Config.SUBJECT_PATTERN, re.VERBOSE)
    
    def parse_subject(self, msg: Message) -> str:
        """安全解码邮件主题"""
        decoded_parts = []
        for part, charset in decode_header(msg.get("Subject", "")):
            try:
                if charset:
                    decoded = part.decode(charset, errors='replace')
                else:
                    # 自动检测编码
                    for encoding in ['utf-8', 'gb18030', 'big5']:
                        try:
                            decoded = part.decode(encoding)
                            break
                        except:
                            continue
                    else:
                        decoded = part.decode('utf-8', errors='replace')
                decoded_parts.append(decoded)
            except Exception as e:
                logger.warning(f"主题解码异常: {str(e)}")
                decoded_parts.append("[解码失败]")
        return "".join(decoded_parts)
    
    def extract_name_id(self, subject: str) -> tuple[str, str] | tuple[None, None]:
        """从主题提取姓名和学号"""
        clean_subject = re.sub(r"\s+", "", subject)
        match = self.subject_re.search(clean_subject)
        if match:
            return match.group(2), match.group(3)  # 姓名、学号
        return None, None
    
    def check_email_date(self, msg: Message) -> bool:
        """检查邮件日期是否符合要求"""
        try:
            date_str = msg.get("Date")
            if not date_str:
                return False
            recv_date = parsedate_to_datetime(date_str)
            start_date = datetime.strptime(Config.START_DATE, "%Y-%m-%d").replace(tzinfo=recv_date.tzinfo)
            return recv_date >= start_date
        except Exception as e:
            logger.error(f"日期校验失败: {str(e)}")
            return False
    
    def extract_docx_attachments(self, msg: Message, temp_dir: Path) -> list[Path]:
        """提取邮件中的DOCX附件到临时目录"""
        attachments = []
        for part in msg.walk():
            if part.get_content_maintype() == 'multipart':
                continue
            if part.get('Content-Disposition') is None:
                continue
            
            filename = part.get_filename()
            if not filename or not filename.endswith(".docx"):
                continue
            
            # 解码文件名
            decoded_filename = self._decode_header(filename)
            safe_filename = re.sub(r"[^\w\-_.]", "", decoded_filename)
            file_path = temp_dir / safe_filename
            
            # 保存附件
            try:
                payload = part.get_payload(decode=True)
                with open(file_path, "wb") as f:
                    f.write(payload)
                attachments.append(file_path)
            except Exception as e:
                logger.error(f"保存附件失败: {str(e)}")
        return attachments
    
    @staticmethod
    def _decode_header(header: str) -> str:
        """解码邮件头（文件名/主题）"""
        try:
            return "".join(
                part.decode(charset or "utf-8", errors='replace') if isinstance(part, bytes)
                else str(part)
                for part, charset in decode_header(header)
            )
        except Exception as e:
            logger.error(f"头信息解码失败: {str(e)}")
            return str(header)
