import os
import logging
from email.message import Message
from email.header import decode_header
import pdfkit
from config import PDF_DIR

logger = logging.getLogger(__name__)

class EmailProcessor:
    def __init__(self):
        # 配置pdfkit（根据实际环境调整wkhtmltopdf路径）
        self.pdf_config = pdfkit.configuration(wkhtmltopdf='/usr/local/bin/wkhtmltopdf')  # 示例路径，需适配实际环境

    def save_email_pdf(self, msg: Message, student_id: str, name: str):
        """将邮件内容保存为PDF"""
        try:
            # 解析邮件内容（文本/HTML）
            text_content = []
            html_content = []
            
            # 遍历邮件部分提取内容
            for part in msg.walk():
                content_type = part.get_content_type()
                charset = part.get_content_charset() or 'utf-8'
                
                if content_type == 'text/plain':
                    try:
                        text = part.get_payload(decode=True).decode(charset, errors='replace')
                        text_content.append(text)
                    except Exception as e:
                        logger.error(f"提取文本内容失败: {e}")
                elif content_type == 'text/html':
                    try:
                        html = part.get_payload(decode=True).decode(charset, errors='replace')
                        html_content.append(html)
                    except Exception as e:
                        logger.error(f"提取HTML内容失败: {e}")
            
            # 修复f-string反斜杠问题：将\n定义为变量
            newline = '\n'
            # 重构content赋值逻辑，规避反斜杠
            if html_content:
                content = newline.join(html_content)
            else:
                content = f"<pre>{newline.join(text_content)}</pre>"
            
            # 生成PDF文件路径
            pdf_filename = f"{name}_{student_id}.pdf"
            pdf_path = PDF_DIR / pdf_filename
            
            # 生成PDF
            pdfkit.from_string(content, str(pdf_path), configuration=self.pdf_config)
            logger.info(f"邮件PDF已保存: {pdf_path}")
            
        except Exception as e:
            logger.error(f"保存邮件PDF失败: {str(e).encode('utf-8', errors='replace').decode('utf-8')}")

    def save_attachments(self, msg: Message, student_id: str, name: str) -> list[os.PathLike]:
        """保存邮件附件，返回附件文件路径列表"""
        attachments = []
        try:
            # 创建学生专属附件目录
            attach_dir = PDF_DIR / f"{name}_{student_id}_attachments"
            attach_dir.mkdir(exist_ok=True, parents=True)
            
            for part in msg.walk():
                # 跳过非附件部分
                if part.get_content_disposition() not in ('attachment', 'inline'):
                    continue
                
                # 解码附件文件名
                filename = part.get_filename()
                if filename:
                    decoded_parts = decode_header(filename)
                    filename = ''.join([
                        part.decode(charset or 'utf-8', errors='replace') 
                        for part, charset in decoded_parts
                    ])
                
                # 保存附件
                if filename:
                    file_path = attach_dir / filename
                    with open(file_path, 'wb') as f:
                        f.write(part.get_payload(decode=True))
                    attachments.append(file_path)
                    logger.info(f"附件已保存: {file_path}")
                    
        except Exception as e:
            logger.error(f"保存附件失败: {str(e).encode('utf-8', errors='replace').decode('utf-8')}")
        
        return attachments
