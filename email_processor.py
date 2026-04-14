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
    def __init__(self) -> None:
        pass

    @staticmethod
    def sanitize_filename(name: str) -> str:
        illegal_chars = r'[<>:"/\\|?*]' if os.name == 'nt' else r'[/]'
        return re.sub(illegal_chars, "_", name).strip()

    def save_attachments(self, msg: Message, student_id: str, name: str) -> List[Path]:
        """提取所有潜在附件"""
        attachments = []
        try:
            safe_name = self.sanitize_filename(name)
            save_dir = PDF_DIR / f"{student_id}_{safe_name}_attachments"
            save_dir.mkdir(exist_ok=True, parents=True)

            for part in msg.walk():
                if part.get_content_maintype() == 'multipart':
                    continue
                
                # 放宽条件：只要有文件名，就认为是附件
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
                        logger.info(f"成功提取附件: {safe_filename}")
                except Exception as e:
                    logger.error(f"保存附件失败: {safe_filename} - {str(e)}")
            return attachments
        except Exception as e:
            logger.error(f"附件处理异常: {student_id}_{name} - {str(e)}")
            return []

    @staticmethod
    def _decode_header(header: str) -> str:
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
        except:
            return str(header)
