from docx import Document
import re

class DocxParser:
    def __init__(self, path):
        self.path = path
        self.ok = False
        self.cnt = 0
        self.parse()

    def parse(self):
        try:
            doc = Document(self.path)
            txt = []
            for t in doc.tables:
                for r in t.rows:
                    for c in r.cells:
                        txt.append(c.text)
            full = "".join(txt)
            self.ok = "是否为学生资助对象" in full and "是" in full
            m = re.split(r"申请理由.*100.*：", full)
            if len(m) > 1:
                self.cnt = len(re.sub(r"\s", "", m[1]))
        except:
            self.ok = False
            self.cnt = 0

    def is_subsidy(self):
        return self.ok

    def get_reason_length(self):
        return self.cnt
