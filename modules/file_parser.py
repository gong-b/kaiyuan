import re
import pdfplumber
from docx import Document

class FileParser:
    @staticmethod
    def parse(path):
        ext = str(path).lower()
        if ext.endswith('.pdf'):
            return FileParser._parse_pdf(path)
        return FileParser._parse_docx(path) # 默认为 Word 解析

    @staticmethod
    def _parse_docx(path):
        """保留并优化你原有的 Docx 逻辑"""
        res = {"is_supported": False, "reason_length": 0, "name": "未知", "sid": None, "apply_class": ""}
        try:
            doc = Document(path)
            # 优先从标题行提取班级
            for para in doc.paragraphs:
                text = para.text.strip()
                match = re.search(r"([^\s]+?班)报名申请表", text)
                if match:
                    res["apply_class"] = match.group(1)
                    break
            
            if doc.tables:
                table = doc.tables[0]
                for row in table.rows:
                    cells = row.cells
                    for i, c in enumerate(cells):
                        t = c.text.strip()
                        if t == "姓名" and i + 1 < len(cells):
                            res["name"] = cells[i+1].text.strip()
                        if t == "学号" and i + 1 < len(cells):
                            res["sid"] = "".join(filter(str.isdigit, cells[i+1].text.strip()))
                        if "是否为学生资助对象" in t and i + 1 < len(cells):
                            v = cells[i+1].text.strip()
                            res["is_supported"] = ("是" in v) and ("不是" not in v)
                        if "申请理由" in t:
                            content = re.sub(r"申请理由.*?[:：]", "", c.text).strip()
                            res["reason_length"] = len(re.sub(r"\s+", "", content))
        except Exception: pass
        return res

    @staticmethod
    def _parse_pdf(path):
        """新增 PDF 解析逻辑（针对 Word 转 PDF 的电子版）"""
        res = {"is_supported": False, "reason_length": 0, "name": "未知", "sid": None, "apply_class": ""}
        try:
            with pdfplumber.open(path) as pdf:
                page = pdf.pages[0]
                # 1. 从纯文本行中找标题班级
                text = page.extract_text()
                if text:
                    for line in text.split('\n'):
                        match = re.search(r"([^\s]+?班)报名申请表", line)
                        if match:
                            res["apply_class"] = match.group(1)
                            break
                
                # 2. 提取表格数据
                table = page.extract_table()
                if table:
                    # 展平表格数据方便定位 
                    flat = [str(item).strip() if item else "" for row in table for item in row]
                    for i, val in enumerate(flat):
                        if val == "姓名" and i + 1 < len(flat):
                            res["name"] = flat[i+1]
                        if val == "学号" and i + 1 < len(flat):
                            res["sid"] = "".join(filter(str.isdigit, flat[i+1]))
                        if "是否为学生资助对象" in val:
                            context = "".join(flat[i:i+2])
                            res["is_supported"] = "是" in context and "不是" not in context
                        if "申请理由" in val:
                            # 理由可能在当前格也可能在下一格
                            content = flat[i+1] if i+1 < len(flat) and len(flat[i+1]) > 10 else flat[i]
                            res["reason_length"] = len(re.sub(r"\s+", "", content))
        except Exception: pass
        return res
