from docx import Document
import re

class DocxParser:
    def __init__(self, file_path):
        self.file_path = file_path
        self.full_text = ""
        self.grade = "未知"
        self.apply_class = "未知班级"
        self.subsidy_flag = False  # 改名！彻底避免冲突
        self.reason_count = 0
        self.parse_safely()

    def parse_safely(self):
        """彻底兜底，绝不崩溃"""
        try:
            doc = Document(self.file_path)
            all_text = []
            # 读取所有段落
            for para in doc.paragraphs:
                if para.text.strip():
                    all_text.append(para.text.strip())
            # 读取所有表格
            for table in doc.tables:
                for row in table.rows:
                    for cell in row.cells:
                        if cell.text.strip():
                            all_text.append(cell.text.strip())
            self.full_text = "\n".join(all_text)

            # 1. 提取班级（超宽松匹配，兼容所有格式）
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

            # 2. 提取年级（超宽松匹配）
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

            # 3. 是否资助对象（彻底修复命名冲突）
            if "是否为学生资助对象" in self.full_text:
                self.subsidy_flag = "是" in self.full_text
            elif "资助对象" in self.full_text:
                self.subsidy_flag = "是" in self.full_text

            # 4. 申请理由字数（兜底统计）
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
            # 兜底：统计全文有效字数
            if self.reason_count == 0:
                clean_all = re.sub(r"\s+", "", self.full_text)
                self.reason_count = len(clean_all)

        except Exception as e:
            print(f"⚠️ 解析文件异常: {str(e)}")
            self.grade = "未知"
            self.apply_class = "未知班级"
            self.subsidy_flag = False
            self.reason_count = 0

    # 对外接口（方法名和属性名彻底区分）
    def get_grade(self):
        return self.grade

    def get_apply_class(self):
        return self.apply_class

    def get_subsidy(self):  # 改名！彻底避免冲突
        return self.subsidy_flag

    def get_reason_count(self):
        return self.reason_count
