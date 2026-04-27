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
                cv.convert(tmp_docx_path, start=0, end=1) # 只转第一页，节省资源
                cv.close()
                
                # 调用 docx 解析逻辑
                res = FileParser._parse_docx(tmp_docx_path)
                
                # 清理临时文件
                if os.path.exists(tmp_docx_path):
                    os.remove(tmp_docx_path)
                return res
            except Exception as e:
                print(f"PDF 转换失败: {e}")
                return {"name": "转换失败", "sid": None, "apply_class": "PDF解析异常", "phone": ""}
                
        return FileParser._parse_docx(path)

    @staticmethod
def _parse_docx(path):
    res = {
        "is_supported": False,
        "reason_length": 0,
        "name": "未知",
        "sid": None,
        "apply_class": "",
        "phone": ""
    }
    try:
        doc = Document(path)

        # 提取班级
        for para in doc.paragraphs:
            text = para.text.strip()
            match = re.search(r"(.+班)报名申请表", text)
            if match:
                res["apply_class"] = match.group(1)
                break

        if not doc.tables:
            return res

        table = doc.tables[0]

        # 逐行读取（精准匹配你的表格）
        for row in table.rows:
            cells = [c.text.strip() for c in row.cells]

            for i, text in enumerate(cells):
                # ==============================================
                # 姓名
                if "姓名" in text and i + 1 < len(cells):
                    res["name"] = cells[i+1]

                # 学号
                if "学号" in text and i + 1 < len(cells):
                    res["sid"] = "".join(filter(str.isdigit, cells[i+1]))

                # 联系方式（你要的：后一格直接提取）
                if "联系方式" in text and i + 1 < len(cells):
                    phone = cells[i+1]
                    res["phone"] = "".join(filter(str.isdigit, phone))

                # ==============================================
                # ✅ 资助对象：修复版！整行找“是”
                if "是否为学生资助对象" in text:
                    # 把这一行后面所有内容拼起来
                    row_text = "".join(cells[i:])
                    # 只要出现“是” → 就是资助对象
                    if "是" in row_text:
                        res["is_supported"] = True
                    else:
                        res["is_supported"] = False

                # ==============================================
                # 申请理由长度
                if "申请理由" in text:
                    content = cells[i] if len(cells[i]) > 30 else (cells[i+1] if i+1 < len(cells) else "")
                    content = re.sub(r"申请理由[^：]*[:：]", "", content).strip()
                    res["reason_length"] = len(re.sub(r"\s+", "", content))

    except Exception:
        pass

    return res
