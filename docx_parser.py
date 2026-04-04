from docx import Document
import re
import os
import subprocess

class DocxParser:
    def __init__(self, file_path):
        self.file_path = file_path
        self.full_text = ""
        self.is_subsidy_flag = False
        self.real_reason = ""
        self.reason_count = 0

        self.parse_safely()

    def parse_safely(self):
        try:
            if not os.path.exists(self.file_path):
                return

            # ==========================================
            # 🔥 核心：自动支持 .doc 和 .docx
            # ==========================================
            ext = self.file_path.lower().split('.')[-1]

            if ext == "docx":
                doc = Document(self.file_path)
                text_list = []
                for para in doc.paragraphs:
                    t = para.text.strip()
                    if t:
                        text_list.append(t)
                for table in doc.tables:
                    for row in table.rows:
                        for cell in row.cells:
                            t = cell.text.strip()
                            if t:
                                text_list.append(t)
                self.full_text = "\n".join(text_list)

            elif ext == "doc":
                # 旧版doc 文本提取（兼容Linux/Windows/Streamlit）
                try:
                    result = subprocess.run(
                        ['catdoc', self.file_path],
                        capture_output=True,
                        text=True,
                        encoding='utf-8',
                        errors='ignore'
                    )
                    self.full_text = result.stdout
                except:
                    self.full_text = ""

            else:
                self.full_text = ""

            # 判断资助
            self.check_subsidy()
            # 提取理由
            self.extract_reason()

        except Exception as e:
            self.full_text = ""
            self.is_subsidy_flag = False
            self.reason_count = 0

    def check_subsidy(self):
        text = self.full_text
        if "是否为学生资助对象" in text and "是" in text:
            self.is_subsidy_flag = True
        elif "是" in text and any(k in text for k in ["资助", "困难", "贫困", "助学金"]):
            self.is_subsidy_flag = True
        else:
            self.is_subsidy_flag = False

    def extract_reason(self):
        text = self.full_text
        patterns = [
            r"申请理由\s*[（\(]\s*不少于\s*100\s*字\s*[）\)]\s*[：:]",
            r"申请理由\s*[：:]",
        ]
        for p in patterns:
            parts = re.split(p, text, flags=re.I)
            if len(parts) >= 2:
                reason = parts[1].strip()
                reason = re.sub(r"\s+", " ", reason)
                self.real_reason = reason
                self.reason_count = len(reason)
                return

        clean = re.sub(r"\s+", "", text)
        self.real_reason = clean
        self.reason_count = len(clean)

    def is_subsidy(self):
        return self.is_subsidy_flag

    def get_reason_length(self):
        return self.reason_count
