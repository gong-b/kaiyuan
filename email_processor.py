import re, os, logging
from pathlib import Path
from email.header import decode_header
from config import PDF_DIR

logger = logging.getLogger(__name__)

class EmailProcessor:
    def sanitize_filename(self, name: str) -> str:
        return re.sub(r'[<>:"/\\|?*]', "_", name).strip()

    def save_attachments(self, msg, student_id, name):
        attachments = []
        save_dir = PDF_DIR / f"{student_id}_{self.sanitize_filename(name)}"
        save_dir.mkdir(exist_ok=True, parents=True)
        for part in msg.walk():
            if part.get_content_maintype() == 'multipart': continue
            filename = part.get_filename()
            if not filename: continue
            
            # 安全解码文件名
            decoded = []
            for p, c in decode_header(filename):
                decoded.append(p.decode(c or 'utf-8', errors='replace') if isinstance(p, bytes) else str(p))
            safe_filename = self.sanitize_filename("".join(decoded))
            
            filepath = save_dir / safe_filename
            payload = part.get_payload(decode=True)
            if isinstance(payload, bytes):
                with open(filepath, "wb") as f: f.write(payload)
                attachments.append(filepath)
        return attachments
