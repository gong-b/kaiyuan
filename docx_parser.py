from docx import Document
import re

class DocxParser:
    def __init__(self, file_path):
        self.file_path = file_path
        self.text = ""
        self.read_text()

    def read_text(self):
        try:
            doc = Document(self.file_path)
            text = []
            # 读取段落
            for para in doc.paragraphs:
                text.append(para.text.strip())
            # 读取表格
            for table in doc.tables:
                for row in table.rows:
                    for cell in row.cells:
                        text.append(cell.text.strip())
            self.text = "\n".join(text)
        except:
            self.text = ""

    def is_subsidy(self):
        """判断是否为资助对象"""
        keywords = ["是", "资助", "困难", "助学金", "贫困"]
        return any(key in self.text for key in keywords)

    def get_reason_length(self):
        """统计申请理由有效字数（去除空白）"""
        # 去除所有空白字符
        clean_text = re.sub(r"\s+", "", self.text)
        # 返回纯中文/数字/字母长度
        return len(clean_text)
