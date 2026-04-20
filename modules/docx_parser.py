from docx import Document
import re

# docx_parser.py

class DocxParser:
    @staticmethod
    def parse(path):
        res = {"is_supported": False, "reason_length": 0}
        try:
            doc = Document(path)
            if not doc.tables:
                return res
            table = doc.tables[0]
            for row in table.rows:
                cells = row.cells
                for i, c in enumerate(cells):
                    t = c.text.strip()
                    # 资助对象判断
                    if "是否为学生资助对象" in t and i + 1 < len(cells):
                        v = cells[i + 1].text.strip()
                        res["is_supported"] = ("是" in v) and ("不是" not in v)
                    
                    # 关键修改：获取“申请理由”右侧单元格的内容字数
                    if "申请理由" in t and i + 1 < len(cells):
                        content = cells[i + 1].text.strip()
                        # 去除所有空白字符后统计长度
                        clean_content = re.sub(r"\s+", "", content)
                        res["reason_length"] = len(clean_content)
        except Exception:
            pass
        return res
