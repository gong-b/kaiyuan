from docx import Document
import re

class DocxParser:
    def __init__(self, file_path):
        self.file_path = file_path
        self.full_text = ""
        self.grade = ""
        self.apply_class = ""
        self.is_subsidy = False
        self.reason_count = 0
        self.parse()

    def parse(self):
        try:
            doc = Document(self.file_path)
            lines = []
            for para in doc.paragraphs:
                lines.append(para.text.strip())
            for table in doc.tables:
                for row in table.rows:
                    for cell in row.cells:
                        lines.append(cell.text.strip())
            self.full_text = "\n".join(lines)

            # ======================
            # 1. 提取班级（精准匹配：开源课堂 → 小提琴班）
            # ======================
            pattern_class = r"“开源课堂”\s*\n?(.+?)报名申请表"
            match_class = re.search(pattern_class, self.full_text, re.DOTALL | re.I)
            self.apply_class = match_class.group(1).strip() if match_class else "未知班级"

            # ======================
            # 2. 提取年级（精准匹配：年级 23级）
            # ======================
            pattern_grade = r"年级\s*([^\s]+)"
            match_grade = re.search(pattern_grade, self.full_text)
            self.grade = match_grade.group(1).strip() if match_grade else "未知"

            # ======================
            # 3. 是否资助对象
            # ======================
            pattern_subsidy = r"是否为学生资助对象\s*(\w+)"
            match_subsidy = re.search(pattern_subsidy, self.full_text)
            self.is_subsidy = (match_subsidy.group(1) == "是") if match_subsidy else False

            # ======================
            # 4. 申请理由字数（100%精准）
            # ======================
            pattern_reason = r"申请理由.*?：(.+)"
            match_reason = re.search(pattern_reason, self.full_text, re.DOTALL | re.I)
            if match_reason:
                reason_text = match_reason.group(1).strip()
                clean_text = re.sub(r"\s+", "", reason_text)
                self.reason_count = len(clean_text)
            else:
                self.reason_count = 0

        except Exception:
            self.grade = "解析失败"
            self.apply_class = "解析失败"
            self.is_subsidy = False
            self.reason_count = 0

    def get_grade(self):
        return self.grade

    def get_apply_class(self):
        return self.apply_class

    def is_subsidy(self):
        return self.is_subsidy

    def get_reason_count(self):
        return self.reason_count
