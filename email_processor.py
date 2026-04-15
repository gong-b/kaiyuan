import os
import aiofiles
from pathlib import Path
from email.message import Message

class EmailProcessor:
    def __init__(self):
        self.attachments_dir = Path("data/attachments")
        self.attachments_dir.mkdir(exist_ok=True, parents=True)

    def save_attachments(self, msg: Message, student_id: str, name: str) -> list[Path]:
        saved_files = []
        
        for part in msg.walk():
            if part.get_content_maintype() == 'multipart':
                continue
            if part.get('Content-Disposition') is None:
                continue

            filename = part.get_filename()
            if not filename:
                continue

            suffix = Path(filename).suffix.lower()
            save_name = f"{student_id}_{name}{suffix}"
            save_path = self.attachments_dir / save_name

            try:
                with open(save_path, 'wb') as f:
                    f.write(part.get_payload(decode=True))
                saved_files.append(save_path)
            except:
                continue

        return saved_files
