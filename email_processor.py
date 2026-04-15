import re, os, logging
from pathlib import Path
from email.header import decode_header
from config import PDF_DIR
import aiofiles  # 需安装：pip install aiofiles

logger = logging.getLogger(__name__)

class EmailProcessor:
    def sanitize_filename(self, name: str) -> str:
        return re.sub(r'[<>:"/\\|?*]', "_", name).strip()

    async def save_attachments(self, msg, student_id, name):
        """异步保存附件，批量写入"""
        attachments = []
        save_dir = PDF_DIR / f"{student_id}_{self.sanitize_filename(name)}"
        save_dir.mkdir(exist_ok=True, parents=True)
        
        # 先收集所有附件数据，再批量写入
        attachment_data = []
        for part in msg.walk():
            if part.get_content_maintype() == 'multipart': continue
            filename = part.get_filename()
            if not filename: continue
            
            decoded = []
            for p, c in decode_header(filename):
                decoded.append(p.decode(c or 'utf-8', errors='replace') if isinstance(p, bytes) else str(p))
            safe_filename = self.sanitize_filename("".join(decoded))
            filepath = save_dir / safe_filename
            payload = part.get_payload(decode=True)
            if isinstance(payload, bytes):
                attachment_data.append((filepath, payload))
                attachments.append(filepath)
        
        # 异步批量写入文件
        for filepath, payload in attachment_data:
            async with aiofiles.open(filepath, "wb") as f:
                await f.write(payload)
        
        return attachments
