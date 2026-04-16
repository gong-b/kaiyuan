import re, os, logging, shutil
from pathlib import Path
from email.header import decode_header
import pdfkit
from config import PDF_DIR

logger = logging.getLogger(__name__)

class EmailProcessor:
    def __init__(self) -> None:
        # 核心修复：自动寻找系统路径，适配 Linux 和 Windows
        wk_path = shutil.which("wkhtmltopdf")
        if not wk_path:
            wk_path = r"D:\program\wkhtmltopdf\bin\wkhtmltopdf.exe"
        
        try:
            self.pdf_config = pdfkit.configuration(wkhtmltopdf=wk_path)
        except Exception as e:
            logger.warning(f"PDF组件加载失败，将跳过PDF生成: {e}")
            self.pdf_config = None

    def sanitize_filename(self, name: str) -> str:
        return re.sub(r'[<>:"/\\|?*]', "_", str(name)).strip()

    def save_attachments(self, msg, student_id, name):
        attachments = []
        save_dir = PDF_DIR / f"{student_id}_{self.sanitize_filename(name)}"
        save_dir.mkdir(exist_ok=True, parents=True)
        for part in msg.walk():
            if part.get_content_maintype() == 'multipart': continue
            filename = part.get_filename()
            if not filename: continue
            
            decoded = []
            for p, c in decode_header(filename):
                if isinstance(p, bytes):
                    decoded.append(p.decode(c or 'utf-8', errors='replace'))
                else:
                    decoded.append(str(p))
            
            safe_filename = self.sanitize_filename("".join(decoded))
            filepath = save_dir / safe_filename
            payload = part.get_payload(decode=True)
            if payload:
                with open(filepath, "wb") as f: f.write(payload)
                attachments.append(filepath)
        return attachments
