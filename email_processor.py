import re, shutil, pdfkit, logging
from pathlib import Path
from email.header import decode_header
from config import PDF_DIR

class EmailProcessor:
    def __init__(self):
        # 自动探测路径：Linux云端在 /usr/bin/，Windows在 D 盘
        wk_path = shutil.which("wkhtmltopdf") or r"D:\program\wkhtmltopdf\bin\wkhtmltopdf.exe"
        try:
            self.pdf_config = pdfkit.configuration(wkhtmltopdf=wk_path)
        except:
            self.pdf_config = None

    def save_attachments(self, msg, student_id, name):
        attachments = []
        save_dir = PDF_DIR / f"{student_id}_{name}"
        save_dir.mkdir(exist_ok=True, parents=True)
        for part in msg.walk():
            if part.get_content_maintype() == 'multipart': continue
            filename = part.get_filename()
            if not filename: continue
            
            # 解码文件名
            decoded = "".join(p.decode(c or 'utf-8') if isinstance(p, bytes) else p for p, c in decode_header(filename))
            filepath = save_dir / re.sub(r'[\\/:*?"<>|]', "_", decoded)
            
            payload = part.get_payload(decode=True)
            if payload:
                with open(filepath, "wb") as f: f.write(payload)
                attachments.append(filepath)
        return attachments
