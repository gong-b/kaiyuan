import re
from email.header import decode_header
from email.utils import parsedate_to_datetime
from pathlib import Path
from .config import Config

class EmailParser:
    def __init__(self):
        self.pattern = re.compile(Config.SUBJECT_PATTERN, re.VERBOSE)

    def parse_subject(self, msg):
        parts = []
        for p, charset in decode_header(msg.get("Subject", "")):
            try:
                if isinstance(p, bytes):
                    p = p.decode(charset or "utf-8", errors="replace")
                parts.append(str(p))
            except:
                parts.append("?")
        return "".join(parts)

    def extract_name_id(self, subject):
        clean = re.sub(r"\s+", "", subject)
        m = self.pattern.search(clean)
        if not m:
            return None, None
        name = m.group(2).strip()
        sid = m.group(3).strip()
        if not sid.isdigit() or len(sid) < 8:
            return None, None
        return name, sid

    def extract_docx_attachments(self, msg, tmp_dir):
        res = []
        for part in msg.walk():
            if part.get_content_maintype() == "multipart":
                continue
            fn = part.get_filename()
            if not fn or not fn.lower().endswith(".docx"):
                continue
            safe = Path(fn).name
            p = tmp_dir / safe
            with open(p, "wb") as f:
                f.write(part.get_payload(decode=True))
            res.append(p)
        return res
