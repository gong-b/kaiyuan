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
                text = para.text.strip()
                match = re.search(r"(.+?班)报名申请表", text)
                if match:
                    res["apply_class"] = match.group(1)
                    break

            # 先把整个表格所有文字全部提取出来（专治合并单元格）
            full_table_text = ""
            reason_candidates = []

            if doc.tables:
                table = doc.tables[0]
                all_cells = []
                for row in table.rows:
                    for cell in row.cells:
                        txt = cell.text.strip()
                        all_cells.append(txt)
                        full_table_text += " " + txt

                        # 收集所有包含“申请理由”的段落
                        if "申请理由" in txt:
                            reason_candidates.append(txt)
                final_reason = ""
                # 从所有包含申请理由的文本里找最长的一段
                for cand in reason_candidates:
                    if len(cand) > len(final_reason):
                        final_reason = cand

                # 截取“申请理由”后面所有内容
                if "申请理由" in final_reason:
                    final_reason = final_reason.split("申请理由", 1)[-1]

                # 清洗空白、标点
                final_reason = re.sub(r"\s+", "", final_reason)
                final_reason = re.sub(r"[：:]", "", final_reason)
                res["reason_length"] = len(final_reason)
                # ==========================================================================

                # 正常解析其他字段
                for i, cell_text in enumerate(all_cells):
                    # 姓名
                    if cell_text == "姓名" and i + 1 < len(all_cells):
                        res["name"] = all_cells[i+1].strip()

                    # 学号
                    elif cell_text == "学号" and i + 1 < len(all_cells):
                        sid = all_cells[i+1].strip()
                        res["sid"] = "".join(filter(str.isdigit, sid))

                    # 联系方式
                    elif any(k in cell_text for k in ["联系方式"]):
                        if i + 1 < len(all_cells):
                            contact = all_cells[i+1].strip()
                            res["contact"] = re.sub(r"[^\d\- ]", "", contact).strip()

                    # 资助对象
                    elif "资助对象" in cell_text:
                        res["is_supported"] = "是" in cell_text

        except Exception as e:
            print("解析错误:", e)

        return res
