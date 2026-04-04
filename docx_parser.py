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
            for paragraph in doc.paragraphs:
                text.append(paragraph.text)
            for table in doc.tables:
                for row in table.rows:
                    for cell in row.cells:
                        text.append(cell.text)
            self.text = "\n".join(text)
        except:
            self.text = ""

    def is_subsidy(self):
        return "是" in self.text and ("资助" in self.text or "困难" in self.text or "助学金" in self.text)

    def count_reason(self):
        match = re.search(r"[\u4e00-\u9fa5]{10,}", self.text)
        if match:
            return len(match.group(0))
        return 0
