import logging
import pdfkit
import re
import os
from pathlib import Path
from email.header import decode_header
from email.message import Message
from typing import Optional, List
from config import PDF_DIR

logger = logging.getLogger(__name__)

class EmailProcessor:
    """邮件处理类（适配跨平台）"""
    pdf_config: pdfkit.configuration.Configuration

    def __init__(self) -> None:
        # 跨平台 wkhtmltopdf 路径配置
        if os.name == 'nt':  # Windows
            wk_path = r"D:\program\wkhtmltopdf\bin\wkhtmltopdf.exe"
        else:  # Linux/Mac (Streamlit Cloud)
            wk_path = "/usr/bin/wkhtmltopdf"
        
        # 兼容无wkhtmltopdf环境
        try:
            self.pdf_config = pdfkit.configuration(wkhtmltopdf=wk_path)
        except:
            self.pdf_config = pdfkit.configuration()
            logger.warning("wkhtmltopdf路径配置失败，使用默认配置")

    @staticmethod
    def sanitize_filename(name: str) -> str:
        """清理非法文件名字符（跨平台）"""
        illegal_chars = r'[<>:"/\\|?*]' if os.name == 'nt' else r'[/]'
        return re.sub(illegal_chars, "_", name).strip()

    def save_email_pdf(self, msg: Message, student_id: str, name: str) -> Optional[Path]:
        """保存邮件为PDF（兼容无HTML内容）"""
        try:
            # 生成安全文件名
            safe_name = self.sanitize_filename(name)
            filename = f"{student_id}_{safe_name}"
            pdf_path = PDF_DIR / f"{filename}.pdf"

            # 提取HTML或纯文本内容
            html_content = []
            text_content = []
            
            for part in msg.walk():
                content_type = part.get_content_type()
                charset = part.get_content_charset() or 'utf-8'
                payload = part.get_payload(decode=True)
                
                if isinstance(payload, bytes):
                    payload = payload.decode(charset, errors='replace')
                
                if content_type == "text/html":
                    html_content.append(payload)
                elif content_type == "text/plain":
                    text_content.append(payload)

            # 优先用HTML，无则用纯文本转HTML
            content = "\n".join(html_content) if html_content else f"<pre>{'\n'.join(text_content)}</pre>"
            
            if content:
                # PDF生成配置（解决中文乱码）
                options = {
                    'encoding': 'UTF-8',
                    'no-images': True,
                    'quiet': ''
                }
                pdfkit.from_string(
                    input=content,
                    output_path=str(pdf_path),
                    configuration=self.pdf_config,
                    options=options
                )
                logger.info(f"PDF保存成功: {pdf_path}")
                return pdf_path
            else:
                logger.warning(f"无邮件内容可保存: {student_id}_{name}")
                return None
                
        except Exception as e:
            logger.error(f"PDF生成失败: {student_id}_{name} - {str(e)}")
            return None

    def save_attachments(self, msg: Message, student_id: str, name: str) -> List[Path]:
        """保存邮件附件（增强容错）"""
        attachments = []
        try:
            safe_name = self.sanitize_filename(name)
            save_dir = PDF_DIR / f"{student_id}_{safe_name}_attachments"
            save_dir.mkdir(exist_ok=True, parents=True)
            
            for part in msg.walk():
                # 跳过多部分邮件容器
                if part.get_content_maintype() == 'multipart':
                    continue
                # 跳过无附件标识的部分
                if part.get('Content-Disposition') is None:
                    continue

                # 解码附件文件名
                filename = part.get_filename()
                if not filename:
                    continue
                decoded_filename = self._decode_header(filename)
                safe_filename = self.sanitize_filename(decoded_filename)
                filepath = save_dir / safe_filename

                # 保存附件
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
