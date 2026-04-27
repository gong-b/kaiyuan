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
            with tempfile.NamedTemporaryFile(suffix='.docx', delete=False) as tmp_docx:
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

            if doc.tables:
                table = doc.tables[0]
                all_cells = []
                for row in table.rows:
                    for cell in row.cells:
                        all_cells.append(cell)

                for i, cell in enumerate(all_cells):
                    cell_text = cell.text.strip()

                    # 姓名
                    if cell_text == "姓名" and i + 1 < len(all_cells):
                        res["name"] = all_cells[i+1].text.strip()

                    # 学号
                    elif cell_text == "学号" and i + 1 < len(all_cells):
                        sid_text = all_cells[i+1].text.strip()
                        res["sid"] = "".join(filter(str.isdigit, sid_text))

                    # 联系方式
                    elif any(key in cell_text for key in ["联系方式", "电话", "手机", "联系电话"]):
                        if i + 1 < len(all_cells):
                            contact_text = all_cells[i+1].text.strip()
                            res["contact"] = re.sub(r"[^\d\- ]", "", contact_text).strip()

                    # 资助对象
                    elif "资助对象" in cell_text:
                        res["is_supported"] = "是" in cell_text

                    # ====================== 终极稳定版：申请理由统计 ======================
                    # 规则：找到“申请理由” → 取它后面所有内容 → 直接统计字数
                    elif "申请理由" in cell_text:
                        # 取出整个单元格的所有文字（无视换行、无视段落）
                        full_text = cell.text

                        # 从“申请理由”这四个字后面开始，截取所有内容
                        if "申请理由" in full_text:
                            reason_part = full_text.split("申请理由", 1)[-1]  # 只保留后面所有内容
                            
                            # 清洗：去掉所有空白、换行、标点符号
                            reason_part = re.sub(r"\s+", "", reason_part)
                            reason_part = re.sub(r"[：:]", "", reason_part)
                            
                            # 统计字数
                            res["reason_length"] = len(reason_part)
                    # ====================================================================

        except Exception as e:
            print(f"解析异常: {e}")
        return res
