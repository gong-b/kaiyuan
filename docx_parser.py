from docx import Document
import re

class DocxParser:
    def __init__(self, file_path):
        self.file_path = file_path
        self.full_text = ""
        self.is_subsidy_flag = False
        self.reason_text = ""
        self.reason_count = 0
        self.parse()

    def parse(self):
        try:
            doc = Document(self.file_path)
            # 只读取表格（你的申请表全在表格里）
            text_parts = []
            for table in doc.tables:
                for row in table.rows:
                    for cell in row.cells:
                        text_parts.append(cell.text.strip())
            self.full_text = "".join(text_parts)

            # 1. 判断资助对象（精准匹配你的表）
            if "是否为学生资助对象" in self.full_text and "是" in self.full_text:
                self.is_subsidy_flag = True

            # 2. 精准提取【申请理由】正文（只算理由，不算标题）
            # 匹配：申请理由（不少于100字）：xxxx
            split_list = re.split(
                r"申请理由\s*[（\(].*?[）\)]\s*[：:]",
                self.full_text,
                flags=re.I
            )
            if len(split_list) >= 2:
                # 只取后面的正文
                raw = split_list[1].strip()
                # 去掉所有空格、换行、制表符
                clean = re.sub(r"\s+", "", raw)
                self.reason_text = clean
                self.reason_count = len(clean)
            else:
                self.reason_count = 0

        except Exception:
            self.full_text = ""
            self.is_subsidy_flag = False
            self.reason_count = 0

    def is_subsidy(self):
        return self.is_subsidy_flag

    def get_reason_length(self):
        return self.reason_count
