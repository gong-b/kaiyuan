from docx import Document
import re
import logging

logger = logging.getLogger(__name__)

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
        """彻底兜底，绝不崩溃"""
        try:
            doc = Document(self.file_path)
            all_text = []
            # 读取所有段落（用于提取班级、年级）
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

            # ==============================================
            # 1. 提取班级（超宽松匹配，兼容所有格式）
            # ==============================================
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

            # ==============================================
            # 2. 提取年级（超宽松匹配）
            # ==============================================
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

            # ==============================================
            # 3. 使用你提供的【精准方法】解析：资助 + 申请理由字数
            # ==============================================
            table = doc.tables[0] if doc.tables else None
            if table:
                reason_text = ""
                # 遍历表格所有单元格
                for row in table.rows:
                    cells = row.cells
                    for cell_index, cell in enumerate(cells):
                        cell_text = cell.text.strip()
                        
                        # --- 精准判断资助对象 ---
                        if "是否为学生资助对象" in cell_text:
                            next_text = cells[cell_index+1].text.strip() if (cell_index + 1 < len(cells)) else ""
                            self.subsidy_flag = (
                                ("是" in next_text or "为" in next_text) 
                                and "不是" not in next_text
                            )
                        
                        # --- 精准提取申请理由字数 ---
                        if "申请理由" in cell_text:
                            reason_paragraphs = [
                                p.text.strip() for p in cell.paragraphs if p.text.strip()
                            ]
                            reason_text = "\n".join(reason_paragraphs)
                            reason_clean = re.sub(r"\s+", "", reason_text)
                            self.reason_count = len(reason_clean)

            # 兜底保护
            if self.reason_count == 0:
                clean_all = re.sub(r"\s+", "", self.full_text)
                self.reason_count = len(clean_all)

        except Exception as e:
            logger.error(f"文档解析异常: {str(e)}")
            self.grade = "未知"
            self.apply_class = "未知班级"
            self.subsidy_flag = False
            self.reason_count = 0

    # 对外接口
    def get_grade(self):
        return self.grade

    def get_apply_class(self):
        return self.apply_class

    def get_subsidy(self):
        return self.subsidy_flag

    def get_reason_count(self):
        return self.reason_count
