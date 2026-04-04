from docx import Document
import re

class DocxParser:
    def __init__(self, file_path):
        self.file_path = file_path
        self.full_text = ""
        self.grade = ""          # 从表格“年级”字段获取
        self.is_subsidy = False
        self.reason_word_count = 0
        self.apply_class = ""    # 从标题获取：多媒体软件班
        self.parse()

    def parse(self):
        try:
            doc = Document(self.file_path)
            # 读取所有表格内容
            text_parts = []
            for table in doc.tables:
                for row in table.rows:
                    for cell in row.cells:
                        text_parts.append(cell.text.strip())
            self.full_text = " || ".join(text_parts)

            # ----------------------
            # 1. 提取年级（从表格“年级”）
            # ----------------------
            grade_pattern = r"年级\s*\|?\s*([^\s|]+)"
            g_match = re.search(grade_pattern, self.full_text)
            if g_match:
                self.grade = g_match.group(1).strip()
            else:
                self.grade = "未知"

            # ----------------------
            # 2. 提取班级（从标题：多媒体软件班）
            # ----------------------
            if "多媒体软件班" in self.full_text:
                self.apply_class = "多媒体软件班"
            elif "书法班" in self.full_text:
                self.apply_class = "书法班"
            else:
                self.apply_class = "未知班级"

            # ----------------------
            # 3. 是否资助对象
            # ----------------------
            if "是否为学生资助对象" in self.full_text and "是" in self.full_text:
                self.is_subsidy = True

            # ----------------------
            # 4. 申请理由字数（100%精准）
            # ----------------------
            reason_split = re.split(r"申请理由.*?不少于100字.*?[：:]", self.full_text)
            if len(reason_split) >= 2:
                reason_raw = reason_split[1].strip()
                clean_reason = re.sub(r"\s+", "", reason_raw)
                self.reason_word_count = len(clean_reason)
            else:
                self.reason_word_count = 0

        except Exception:
            self.grade = "解析失败"
            self.is_subsidy = False
            self.reason_word_count = 0
            self.apply_class = "解析失败"

    # 对外接口
    def get_grade(self):
        return self.grade

    def get_subsidy(self):
        return self.is_subsidy

    def get_reason_count(self):
        return self.reason_word_count

    def get_apply_class(self):
        return self.apply_class
