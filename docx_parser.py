from docx import Document
import re

class DocxParser:
    def __init__(self, file_path):
        self.file_path = file_path
        self.full_text = ""
        self.grade = "未知"
        self.apply_class = "未知班级"
        self.subsidy_flag = False
        self.reason_count = 0
        self.parse_safely()

    def parse_safely(self):
        try:
            doc = Document(self.file_path)
            all_text = []
            for para in doc.paragraphs:
                if para.text.strip():
                    all_text.append(para.text.strip())
            for table in doc.tables:
                for row in table.rows:
                    for cell in row.cells:
                        if cell.text.strip():
                            all_text.append(cell.text.strip())
            self.full_text = " ".join(all_text)

            # 提取班级（100% 匹配你的报名表）
            match = re.search(r"([^\s]+班)", self.full_text)
            if match:
                self.apply_class = match.group(1).strip()

            # 提取年级
            g_match = re.search(r"(\d+级)", self.full_text)
            if g_match:
                self.grade = g_match.group(1).strip()

            # 资助对象
            self.subsidy_flag = "是" in self.full_text and "学生资助对象" in self.full_text

            # 申请理由字数
            r_match = re.search(r"申请理由.*?：(.+)", self.full_text)
            if r_match:
                raw = r_match.group(1).strip()
                clean = re.sub(r"\s+", "", raw)
                self.reason_count = len(clean)

        except:
            self.grade = "未知"
            self.apply_class = "未知班级"
            self.subsidy_flag = False
            self.reason_count = 0

    def get_grade(self):
        return self.grade

    def get_apply_class(self):
        return self.apply_class

    def get_subsidy(self):
        return self.subsidy_flag

    def get_reason_count(self):
        return self.reason_count
