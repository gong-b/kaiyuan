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
            all_text = []
            for table in doc.tables:
                for row in table.rows:
                    for cell in row.cells:
                        all_text.append(cell.text.strip())
            self.full_text = " ".join(all_text)

            # ==========================================
            # 1. 提取年级（精准匹配你的表格）
            # ==========================================
            grade_match = re.search(r'年级\s*[:：]\s*(\w+)', self.full_text)
            if grade_match:
                self.grade = grade_match.group(1).strip()
            else:
                self.grade = "未知"

            # ==========================================
            # 2. 提取班级：从【开源课堂】XXX 提取
            # ==========================================
            class_match = re.search(r'【开源课堂】\s*([^\s】]+)', self.full_text)
            if class_match:
                self.apply_class = class_match.group(1).strip()
            else:
                self.apply_class = "未知班级"

            # ==========================================
            # 3. 是否资助对象
            # ==========================================
            self.is_subsidy = "是否为学生资助对象" in self.full_text and "是" in self.full_text

            # ==========================================
            # 4. 申请理由字数（精准）
            # ==========================================
            reason_parts = re.split(r'申请理由.*?100.*?[：:]', self.full_text)
            if len(reason_parts) >= 2:
                clean = re.sub(r'\s+', '', reason_parts[1])
                self.reason_count = len(clean)
            else:
                self.reason_count = 0

        except:
            self.grade = "解析失败"
            self.apply_class = "解析失败"
            self.is_subsidy = False
            self.reason_count = 0

    # 对外接口
    def get_grade(self):
        return self.grade

    def get_apply_class(self):
        return self.apply_class

    def is_subsidy(self):
        return self.is_subsidy

    def get_reason_count(self):
        return self.reason_count
