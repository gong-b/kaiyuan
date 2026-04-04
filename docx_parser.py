from docx import Document
import re

class DocxParser:
    def __init__(self, file_path):
        self.file_path = file_path
        self.full_text = ""
        self.is_subsidy = False
        self.real_reason = ""
        self.reason_count = 0

        # 一次性解析完所有内容
        self.parse_all()

    def parse_all(self):
        try:
            doc = Document(self.file_path)
            text_list = []

            # 读取所有表格内容（你的申请表全在表格里）
            for table in doc.tables:
                for row in table.rows:
                    for cell in row.cells:
                        cell_txt = cell.text.strip()
                        text_list.append(cell_txt)

            # 合并所有文本
            self.full_text = "\n".join(text_list)

            # 1. 判断是否资助对象
            self.check_subsidy()

            # 2. 提取真实申请理由（只提取理由正文）
            self.extract_reason()

        except Exception as e:
            self.full_text = ""
            self.is_subsidy = False
            self.real_reason = ""
            self.reason_count = 0

    def check_subsidy(self):
        # 匹配你的表格：是否为学生资助对象 | 是
        if "是否为学生资助对象" in self.full_text:
            # 后面跟着“是”才判定为资助对象
            self.is_subsidy = "是" in self.full_text

    def extract_reason(self):
        # 按你的表格结构：从“申请理由（不少于100字）”后面开始提取正文
        reason_split = re.split(
            r"申请理由\s*[（\(]\s*不少于\s*100\s*字\s*[）\)]\s*[：:]",
            self.full_text,
            flags=re.I
        )

        if len(reason_split) >= 2:
            reason = reason_split[1].strip()

            # 去掉多余空行、空格、换行
            reason = re.sub(r"\s+", " ", reason)
            reason = reason.strip()

            self.real_reason = reason
            self.reason_count = len(reason)
        else:
            self.real_reason = ""
            self.reason_count = 0

    # 给外部调用的接口
    def is_subsidy(self):
        return self.is_subsidy

    def get_reason_length(self):
        return self.reason_count
