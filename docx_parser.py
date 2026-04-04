from docx import Document
import re

class DocxParser:
    def __init__(self, filepath):
        self.filepath = filepath
        self.is_supported_flag = False
        self.reason_len = 0
        self.parse()

    def parse(self):
        try:
            doc = Document(self.filepath)
            text = []
            for table in doc.tables:
                for row in table.rows:
                    for cell in row.cells:
                        text.append(cell.text.strip())
            full_text = "".join(text)

            # 判断是否资助对象
            if "是" in full_text and ("资助" in full_text or "困难" in full_text):
                self.is_supported_flag = True

            # 理由字数
            reason_text = re.sub(r"\s+", "", full_text)
            self.reason_len = len(reason_text)
        except:
            self.is_supported_flag = False
            self.reason_len = 0

    def is_supported(self):
        return self.is_supported_flag

    def get_reason_length(self):
        return self.reason_len
