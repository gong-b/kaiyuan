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
                return {"name": "转换失败", "sid": None, "apply_class": "PDF解析异常", "contact": ""}
                
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
                full_text_list = [cell.text.strip() for row in table.rows for cell in row.cells]
                
                for i, text in enumerate(full_text_list):
                    # 姓名
                    if text == "姓名" and i + 1 < len(full_text_list):
                        res["name"] = full_text_list[i+1]
                    # 学号
                    if text == "学号" and i + 1 < len(full_text_list):
                        res["sid"] = "".join(filter(str.isdigit, full_text_list[i+1]))
                    # 联系方式（手机/电话）
                    if any(key in text for key in ["联系方式", "电话", "手机", "联系电话"]):
                        if i + 1 < len(full_text_list):
                            contact = re.sub(r"[^\d\- ]", "", full_text_list[i+1])
                            res["contact"] = contact.strip()
                    # 资助对象
                    if "资助对象" in text:
                        context = "".join(full_text_list[i:i+3])
                        res["is_supported"] = "是" in context and "不是" not in context

                    # ====================== 核心修复 ======================
                    # 申请理由：优先读当前格，再读下一格，兼容“标题+内容同格”
                    if "申请理由" in text:
                        # 先读当前单元格内容
                        current_content = full_text_list[i]
                        # 再读下一单元格内容
                        next_content = full_text_list[i+1] if (i+1 < len(full_text_list)) else ""
                        # 合并
                        total_content = current_content + next_content
                        # 清洗掉标题文字
                        total_content = re.sub(r"申请理由.*?[:：]", "", total_content).strip()
                        # 去空白统计真实长度
                        res["reason_length"] = len(re.sub(r"\s+", "", total_content))
                    # ======================================================

        except Exception:
            pass
        return res
