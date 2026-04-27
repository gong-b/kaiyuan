import os
import re
import tempfile
from docx import Document
from pdf2docx import Converter

class FileParser:
    @staticmethod
    def parse(path):
        ext = str(path).lower()
        if ext.endswith('.pdf'):
            with tempfile.NamedNamedTemporaryFile(suffix='.docx', delete=False) as tmp_docx:
                tmp_docx_path = tmp_docx.name
            
            try:
                cv = Converter(path)
                cv.convert(tmp_docx_path, start=0, end=1)
                cv.close()
                res = FileParser._parse_docx(tmp_docx_path)
                if os.path.exists(tmp_docx_path):
                    os.remove(tmp_docx_path)
                return res
            except Exception as e:
                print(f"PDF 转换失败: {e}")
                return {"is_supported": False, "reason_length": 0, "name": "转换失败", "sid": None, "apply_class": "PDF解析异常", "contact": ""}
                
        return FileParser._parse_docx(path)

    @staticmethod
    def _parse_docx(path):
        res = {
            "is_supported": False,
            "reason_length": 0,
            "name": "未知",
            "sid": None,
            "apply_class": "",
            "contact": ""
        }
        try:
            doc = Document(path)
            
            # 提取班级
            for para in doc.paragraphs:
                text = para.text.replace(" ", "")
                match = re.search(r"(.+?班)报名申请表", text)
                if match:
                    res["apply_class"] = match.group(1)
                    break

            # 按单元格逐行解析（精准匹配你的标准）
            if doc.tables:
                table = doc.tables[0]
                for row in table.rows:
                    for cell in row.cells:
                        cell_text = cell.text.strip()

                        # 姓名
                        if cell_text == "姓名":
                            next_cell = FileParser._get_next_cell(row, cell)
                            if next_cell:
                                res["name"] = next_cell.strip()

                        # 学号
                        elif cell_text == "学号":
                            next_cell = FileParser._get_next_cell(row, cell)
                            if next_cell:
                                res["sid"] = "".join(filter(str.isdigit, next_cell.strip()))

                        # 联系方式
                        elif any(key in cell_text for key in ["联系方式", "电话", "手机", "联系电话"]):
                            next_cell = FileParser._get_next_cell(row, cell)
                            if next_cell:
                                contact = re.sub(r"[^\d\- ]", "", next_cell.strip())
                                res["contact"] = contact

                        # 资助对象
                        elif "资助对象" in cell_text:
                            res["is_supported"] = "是" in cell_text

                        # ====================== 你指定的申请理由解析规则 ======================
                        elif "申请理由" in cell_text:
                            # 完全使用你提供的逻辑
                            reason_paragraphs: list[str] = [
                                p.text.strip()
                                for p in cell.paragraphs
                                if p.text.strip()
                            ]
                            reason_text: str = "\n".join(reason_paragraphs)
                            reason_text = re.sub(r"\s+", "", reason_text)
                            res["reason_length"] = len(reason_text.replace(" ", ""))
                        # ====================================================================

        except Exception as e:
            print(f"解析异常: {e}")
            return res
        return res

    @staticmethod
    def _get_next_cell(row, current_cell):
        """获取同一行的下一个单元格内容"""
        try:
            idx = row.cells.index(current_cell)
            if idx + 1 < len(row.cells):
                return row.cells[idx + 1].text.strip()
        except:
            return None
        return None
