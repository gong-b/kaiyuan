import logging
import pdfkit
import re
from pathlib import Path
from email.header import decode_header
from email.message import Message
from typing import Optional, List
from config import PDF_DIR

logger = logging.getLogger(__name__)

class EmailProcessor:
    pdf_config: 'pdfkit.configuration.Configuration'

    def __init__(self) -> None:
        self.pdf_config = pdfkit.configuration(wkhtmltopdf=r"D:\program\wkhtmltopdf\bin\wkhtmltopdf.exe")

    def sanitize_filename(self, name: str) -> str:
        """清理非法文件名字符"""
        return "".join(c for c in name if c.isalnum() or c in (' ', '_', '-', '.')).rstrip()

    def save_email_pdf(self, msg: Message, student_id: str, name: str) -> Optional[Path]:
        """保存邮件内容为PDF"""
        try:
            # 生成PDF文件名
            safe_name: str = re.sub(r"[^\w\-_]", "", name)  # 清理特殊字符
            filename: str = f"{student_id}_{safe_name}"
            pdf_path: Path = PDF_DIR / (self.sanitize_filename(filename) + '.pdf')
            
            # 提取HTML内容
            html_content: List[str] = []
            for part in msg.walk():
                if part.get_content_type() == "text/html":
                    payload = part.get_payload(decode=True)
                    charset: str = part.get_content_charset() or 'utf-8'
                    if isinstance(payload, bytes):
                        html: str = payload.decode(charset, errors='replace')
                    elif isinstance(payload, str):
                        html: str = payload
                    else:
                        html: str = ""
                    html_content.append(html)
            
            if html_content:
                pdfkit.from_string(                 #type: ignore[arg-type]
                    input="\n".join(html_content),  # type: ignore[arg-type]
                    output_path=str(pdf_path),      # type: ignore[arg-type]
                    configuration=self.pdf_config,  # type: ignore[arg-type]
                    options={'encoding': "UTF-8"}   # type: ignore[arg-type]
                )
                logger.info(f"PDF保存成功: {pdf_path}")
            return pdf_path
        except Exception as e:
            logger.error(f"PDF生成失败: {e}")
            return None

    def save_attachments(self, msg: Message, student_id: str, name: str) -> List[Path]:
        """保存邮件附件"""
        attachments: List[Path] = []
        try:
            safe_name: str = re.sub(r"[^\w\-_]", "", name)
            save_dir: Path = PDF_DIR / f"{student_id}_{safe_name}_attachments"
            save_dir.mkdir(exist_ok=True, parents=True)
            
            for part in msg.walk():
                if part.get_content_maintype() == 'multipart':
                    continue
                if part.get('Content-Disposition') is None:
                    continue

                filename: Optional[str] = part.get_filename()
                if filename:
                    decoded_filename: str = self._decode_header(filename)
                    safe_filename: str = self.sanitize_filename(decoded_filename)
                    filepath: Path = save_dir / safe_filename
                    
                    with open(filepath, "wb") as f:
                        payload = part.get_payload(decode=True)
                        if isinstance(payload, bytes):
                            f.write(payload)
                        elif isinstance(payload, str):
                            f.write(payload.encode(part.get_content_charset() or "utf-8", errors="replace"))
                        else:
                            logger.warning(f"附件内容类型未知，未保存: {safe_filename}")
                    attachments.append(filepath)
                    logger.info(f"附件保存成功: {filepath}")
            return attachments
        except Exception as e:
            logger.error(f"附件保存失败: {e}")
            return []

    @staticmethod 
    def _decode_header(header: str) -> str:
        """安全地解码邮件头信息"""
        try:
            return "".join(
                part.decode(charset or "utf-8", errors="replace") if isinstance(part, bytes)
                else str(part)
                for part, charset in decode_header(header)
            )
        except Exception as e:
            logger.error(f"头信息解码失败: {str(e)}")
            # Fallback for unexpected errors
            if isinstance(header, bytes):
                return header.decode("utf-8", errors="replace")
            return str(header)

