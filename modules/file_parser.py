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
            # 创建临时 docx 文件路径
            with tempfile.NamedTemporaryFile(suffix='.docx', delete=False) as tmp_docx:
                tmp_docx_path = tmp_docx.name
            
            try:
                # 核心步骤：将 PDF 转换为 Docx
                cv = Converter(path)
                cv.convert(tmp_docx_path, start=0, end=1)  # 只转第一页，节省资源
                cv.close()
                
                # 调用 docx 解析逻辑
                res = FileParser._parse_docx(tmp_docx_path)
                
                # 清理临时文件
                if os.path.exists(tmp_docx_path):
                    os.remove(tmp_docx_path)
                return res
            except Exception as e:
                print(f"PDF 转换失败: {e}")
                return {"name": "转换失败", "sid": None, "apply_class": "PDF解析异常"}
                
        return FileParser._parse_docx(path)

    @staticmethod
    def _parse_docx(path):
        """稳健的 Docx 解析逻辑，提取报名关键信息"""
        res = {
            "is_supported": False, 
            "reason_length": 0, 
            "name": "未知", 
            "sid": None, 
            "apply_class": ""
        }
        try:
            doc = Document(path)
            
            # 1. 提取班级（匹配「XXX班报名申请表」格式）
            for para in doc.paragraphs:
                text = para.text.replace(" ", "")
                match = re.search(r"(.+?班)报名申请表", text)
                if match:
                    res["apply_class"] = match.group(1)
                    break

            if doc.tables:
                table = doc.tables[0]
                # 将表格内容展平处理，增加容错性
                full_text_list = [cell.text.strip() for row in table.rows for cell in row.cells]
                
                for i, text in enumerate(full_text_list):
                    # 提取姓名
                    if text == "姓名" and i + 1 < len(full_text_list):
                        res["name"] = full_text_list[i+1].strip()
                    # 提取学号（仅保留数字）
                    if text == "学号" and i + 1 < len(full_text_list):
                        res["sid"] = "".join(filter(str.isdigit, full_text_list[i+1]))
                    # 判断是否为资助对象
                    if "资助对象" in text:
                        context = "".join(full_text_list[i:i+3])
                        res["is_supported"] = "是" in context and "不是" not in context
                    # 提取申请理由并计算长度
                    if "申请理由" in text:
                        content = full_text_list[i] if len(full_text_list[i]) > 30 else (full_text_list[i+1] if i+1 < len(full_text_list) else "")
                        content = re.sub(r"申请理由.*?[:：]", "", content).strip()
                        res["reason_length"] = len(re.sub(r"\s+", "", content))
        except Exception as e:
            print(f"Docx 解析失败: {e}")
        return res
