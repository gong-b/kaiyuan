from docx import Document
import re

class DocxParser:
    @staticmethod
    def parse(path):
        res = {"is_supported": False, "reason_length": 0}
        try:
            doc = Document(path)
            if not doc.tables:
                return res
            for row in doc.tables[0].rows:
                cells = row.cells
                for i, c in enumerate(cells):
                    t = c.text.strip()
                    if "是否为学生资助对象" in t and i+1 < len(cells):
                        v = cells[i+1].text.strip()
                        res["is_supported"] = "是" in v and "不是" not in v
                    if "申请理由" in t:
                        s = re.sub(r"\s+", "", c.text.strip())
                        res["reason_length"] = len(s)
        except:
            pass
        return res
