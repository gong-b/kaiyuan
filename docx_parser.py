from docx import Document
import re

class DocxParser:
    def __init__(self, file_path):
        self.file_path = file_path
        self.full_text = ""
        self.is_subsidy_flag = False
        self.reason_count = 0
        self.parse()

    def parse(self):
        try:
            doc = Document(self.file_path)
            lines = []

            # 读取所有表格内容
            for table in doc.tables:
                for row in table.rows:
                    for cell in row.cells:
                        txt = cell.text.strip()
                        if txt:
                            lines.append(txt)

            self.full_text = "\n".join(lines)
            self.check_subsidy()
            self.extract_reason()

        except Exception as e:
            self.full_text = ""
            self.is_subsidy_flag = False
            self.reason_count = 0

    def check_subsidy(self):
        # 精准匹配你的申请表：是否为学生资助对象 → 是
        if "是否为学生资助对象" in self.full_text and "是" in self.full_text:
            self.is_subsidy_flag = True

    def extract_reason(self):
        # 精准提取“申请理由（不少于100字）：”后面的内容
        match = re.split(r"申请理由.*不少于100字.*：", self.full_text)
        if len(match) > 1:
            reason = match[1].strip()
            reason = re.sub(r"\s+", "", reason)
            self.reason_count = len(reason)
        else:
            self.reason_count = 0

    def is_subsidy(self):
        return self.is_subsidy_flag

    def get_reason_length(self):
        return self.reason_count
