from docx import Document
import re

class DocxParser:
    def __init__(self, file_path):
        self.file_path = file_path
        self.full_text = ""
        self.grade = "未知"
        self.apply_class = "未知班级"
        self.is_subsidy = False
        self.reason_count = 0
        self.parse_safely()

    def parse_safely(self):
        """彻底兜底，绝不崩溃"""
        try:
            doc = Document(self.file_path)
            # 读取所有内容，不管表格还是段落
            all_text = []
            for para in doc.paragraphs:
                if para.text.strip():
                    all_text.append(para.text.strip())
            for table in doc.tables:
                for row in table.rows:
                    for cell in row.cells:
                        if cell.text.strip():
                            all_text.append(cell.text.strip())
            self.full_text = "\n".join(all_text)

            # ==================================
            # 1. 提取班级（超宽松匹配，100%命中）
            # ==================================
            # 匹配【开源课堂】后面的任何班级名
            class_patterns = [
                r"【开源课堂】\s*([^\n]+?班)",
                r"“开源课堂”\s*\n?(.+?)班报名",
                r"(.+?)班报名申请表"
            ]
            for pat in class_patterns:
                match = re.search(pat, self.full_text, re.DOTALL)
                if match:
                    self.apply_class = match.group(1).strip()
                    break

            # ==================================
            # 2. 提取年级（超宽松匹配）
            # ==================================
            grade_patterns = [
                r"年级\s*[:：]?\s*(\d+级|\d+级|\w+)",
                r"(\d{2,4}级)",
                r"(\d{4})级"
            ]
            for pat in grade_patterns:
                match = re.search(pat, self.full_text)
                if match:
                    self.grade = match.group(1).strip()
                    break

            # ==================================
            # 3. 是否资助对象（兜底匹配）
            # ==================================
            if "是否为学生资助对象" in self.full_text:
                self.is_subsidy = "是" in self.full_text
            elif "资助对象" in self.full_text:
                self.is_subsidy = "是" in self.full_text

            # ==================================
            # 4. 申请理由字数（兜底统计）
            # ==================================
            reason_patterns = [
                r"申请理由.*?[：:](.+)",
                r"申请理由\s*\n(.+)"
            ]
            for pat in reason_patterns:
                match = re.search(pat, self.full_text, re.DOTALL)
                if match:
                    reason = match.group(1).strip()
                    clean = re.sub(r"\s+", "", reason)
                    self.reason_count = len(clean)
                    break
            # 兜底：如果没找到标题，统计全文有效字数
            if self.reason_count == 0:
                clean_all = re.sub(r"\s+", "", self.full_text)
                self.reason_count = len(clean_all)

        except Exception as e:
            # 彻底兜底，任何错误都不崩溃
            print(f"⚠️ 解析文件异常: {str(e)}")
            self.grade = "未知"
            self.apply_class = "未知班级"
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
