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

            # 全局全文文本（解决合并单元格读不到的问题）
            full_text = ""

            # 先把整个表格所有内容读出来
            all_cells_text = []
            if doc.tables:
                table = doc.tables[0]
                for row in table.rows:
                    for cell in row.cells:
                        t = cell.text.strip()
                        all_cells_text.append(t)
                        full_text += " " + t

            # ====================== 1. 修复资助对象判断（合并单元格兼容） ======================
            # 只要全文任意位置出现“是” + 附近有“资助对象”，就判定为是
            if "资助对象" in full_text and "是" in full_text:
                res["is_supported"] = True
            else:
                res["is_supported"] = False

            # ====================== 2. 修复申请理由字数（合并单元格兼容） ======================
            reason_text = ""
            for txt in all_cells_text:
                if "申请理由" in txt:
                    reason_text = txt
                    break

            # 截取“申请理由”之后所有内容
            if "申请理由" in reason_text:
                reason_text = reason_text.split("申请理由", 1)[1]

            # 清洗并统计字数（保留标点，去掉空白）
            reason_text = re.sub(r"\s+", "", reason_text)
            res["reason_length"] = len(reason_text)

            # ====================== 3. 提取姓名、学号、联系方式 ======================
            for i, txt in enumerate(all_cells_text):
                # 姓名
                if txt == "姓名" and i + 1 < len(all_cells_text):
                    res["name"] = all_cells_text[i+1].strip()

                # 学号
                if txt == "学号" and i + 1 < len(all_cells_text):
                    sid_val = all_cells_text[i+1].strip()
                    res["sid"] = "".join(filter(str.isdigit, sid_val))

                # 联系方式
                if any(key in txt for key in ["联系方式", "电话", "手机"]):
                    if i + 1 < len(all_cells_text):
                        contact_val = all_cells_text[i+1].strip()
                        res["contact"] = re.sub(r"[^\d\s\-]", "", contact_val).strip()

        except Exception as e:
            print("解析异常:", e)

        return res
