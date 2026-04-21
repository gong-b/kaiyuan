import re
import pdfplumber
from docx import Document

class FileParser:
    @staticmethod
    def parse(path):
        ext = str(path).lower()
        if ext.endswith('.pdf'):
            return FileParser._parse_pdf(path)
        return FileParser._parse_docx(path)

    @staticmethod
    def _parse_docx(path):
        res = {"is_supported": False, "reason_length": 0, "name": "未知", "sid": None, "apply_class": ""}
        try:
            doc = Document(path)
            # 1. 提取班级：从标题行精准抓取（如：多媒体软件班）
            for para in doc.paragraphs:
                text = para.text.strip()
                match = re.search(r"“?(.+?班)”?", text)
                if match and "报名申请表" in text:
                    res["apply_class"] = match.group(1)
                    break
            
            # 2. 提取表格数据
            if doc.tables:
                table = doc.tables[0]
                full_cells = [cell.text.strip() for row in table.rows for cell in row.cells]
                for i, t in enumerate(full_cells):
                    if t == "姓名" and i + 1 < len(full_cells):
                        res["name"] = full_cells[i+1]
                    if t == "学号" and i + 1 < len(full_cells):
                        res["sid"] = "".join(filter(str.isdigit, full_cells[i+1]))
                    if "资助对象" in t:
                        # 检查当前格或后一格是否有“是”
                        context = "".join(full_cells[max(0, i-1):i+2])
                        res["is_supported"] = "是" in context and "不是" not in context
                    if "申请理由" in t:
                        content = full_cells[i] if len(full_cells[i]) > 20 else (full_cells[i+1] if i+1 < len(full_cells) else "")
                        res["reason_length"] = len(re.sub(r"\s+", "", content))
        except Exception: pass
        return res

    @staticmethod
    def _parse_pdf(path):
        res = {"is_supported": False, "reason_length": 0, "name": "未知", "sid": None, "apply_class": ""}
        try:
            with pdfplumber.open(path) as pdf:
                page = pdf.pages[0]
                text = page.extract_text()
                
                # 1. 提取标题中的班级
                if text:
                    lines = text.split('\n')
                    for line in lines:
                        match = re.search(r"“?(.+?班)”?", line)
                        if match and "报名申请表" in line:
                            res["apply_class"] = match.group(1)
                            break
                
                # 2. 提取表格（针对 PDF 布局优化）
                table = page.extract_table()
                if table:
                    flat = [str(item).replace('\n', '').strip() if item else "" for row in table for item in row]
                    for i, val in enumerate(flat):
                        if val == "姓名" and i + 1 < len(flat):
                            res["name"] = flat[i+1]
                        if val == "学号" and i + 1 < len(flat):
                            res["sid"] = "".join(filter(str.isdigit, flat[i+1]))
                        if "资助对象" in val or "是否为学生" in val:
                            # 针对 PDF 换行情况，检查附近格
                            context = "".join(flat[max(0, i-2):i+3])
                            res["is_supported"] = "是" in context and "不是" not in context
                        if "申请理由" in val:
                            # 理由可能在同一格也可能在后面
                            content = flat[i+1] if i+1 < len(flat) and len(flat[i+1]) > 10 else flat[i]
                            res["reason_length"] = len(re.sub(r"\s+", "", content))
        except Exception: pass
        return res
