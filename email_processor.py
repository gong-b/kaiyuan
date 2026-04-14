import logging
import re
import os
from pathlib import Path
from email.header import decode_header
from email.message import Message
from typing import Optional, List
from config import PDF_DIR

logger = logging.getLogger(__name__)

class EmailProcessor:
    """邮件处理类（仅保留附件提取，彻底删除PDF功能）"""
    def __init__(self) -> None:
        pass

    @staticmethod
    def sanitize_filename(name: str) -> str:
        """清理非法文件名字符（跨平台）"""
        illegal_chars = r'[<>:"/\\|?*]' if os.name == 'nt' else r'[/]'
        return re.sub(illegal_chars, "_", name).strip()

    def save_email_pdf(self, msg: Message, student_id: str, name: str) -> Optional[Path]:
        """彻底关闭PDF生成，直接返回None"""
        logger.info("PDF生成功能已关闭（避免wkhtmltopdf报错）")
        return None

    def save_attachments(self, msg: Message, student_id: str, name: str) -> List[Path]:
        """保存邮件附件（仅保留核心逻辑，增强容错）"""
        attachments = []
        try:
            safe_name = self.sanitize_filename(name)
            save_dir = PDF_DIR / f"{student_id}_{safe_name}_attachments"
            save_dir.mkdir(exist_ok=True, parents=True)

            for part in msg.walk():
                if part.get_content_maintype() == 'multipart':
                    continue
                if part.get('Content-Disposition') is None:
                    continue

                filename = part.get_filename()
                if not filename:
                    continue

                decoded_filename = self._decode_header(filename)
                safe_filename = self.sanitize_filename(decoded_filename)
                filepath = save_dir / safe_filename

                try:
                    payload = part.get_payload(decode=True)
                    if isinstance(payload, bytes):
                        with open(filepath, "wb") as f:
                            f.write(payload)
                        attachments.append(filepath)
                        logger.info(f"附件保存成功: {filepath}")
                    else:
                        logger.warning(f"附件内容非字节类型: {safe_filename}")
                except Exception as e:
                    logger.error(f"保存附件失败: {safe_filename} - {str(e)}")
            return attachments
        except Exception as e:
            logger.error(f"附件保存失败: {student_id}_{name} - {str(e)}")
            return []

    @staticmethod
    def _decode_header(header: str) -> str:
        """安全解码邮件头（增强容错）"""
        try:
            decoded_parts = []
            for part, charset in decode_header(header):
                if isinstance(part, bytes):
                    decode_charset = charset or 'utf-8'
                    decoded_part = part.decode(decode_charset, errors='replace')
                else:
                    decoded_part = str(part)
                decoded_parts.append(decoded_part)
            return "".join(decoded_parts)
        except Exception as e:
            logger.error(f"头信息解码失败: {str(e)}")
            return str(header) if not isinstance(header, bytes) else header.decode('utf-8', errors='replace')
