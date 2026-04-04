from docx import Document
import re

class DocxParser:
    def __init__(self, file_path):
        self.file_path = file_path
        self.full_text = ""
        self.is_subsidy = False
        self.reason_word_count = 0
        self.major_class = ""
        self.parse()

    def parse(self):
        try:
            doc = Document(self.file_path)
            parts = []
            for table in doc.tables:
                for row in table.rows:
                    for cell in row.cells:
                        parts.append(cell.text.strip())
            self.full_text = " ".join(parts)

            # 资助判断
            self.is_subsidy = "是否为学生资助对象" in self.full_text and "是" in self.full_text

            # 提取班级（自动识别 多媒体班/书法班/绘画班 等）
            class_pattern = r"([^\s]{2,8}班)"
            match_class = re.search(class_pattern, self.full_text)
            if match_class:
                self.major_class = match_class.group(1)

            # 提取申请理由字数（100%精准）
            reason_parts = re.split(r"申请理由.*?[：:]", self.full_text)
            if len(reason_parts) >= 2:
                clean = re.sub(r"\s+", "", reason_parts[1].strip())
                self.reason_word_count = len(clean)
            else:
                self.reason_word_count = 0

        except:
            self.is_subsidy = False
            self.reason_word_count = 0
            self.major_class = ""

    def get_subsidy(self):
        return self.is_subsidy

    def get_reason_count(self):
        return self.reason_word_count

    def get_class(self):
        return self.major_class
