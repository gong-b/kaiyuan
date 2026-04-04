from docx import Document
import re
import os

class DocxParser:
    def __init__(self, file_path):
        self.file_path = file_path
        self.full_text = ""
        self.is_subsidy_flag = False
        self.real_reason = ""
        self.reason_count = 0

        # 一次性解析，全程兜底
        self.parse_safely()

    def parse_safely(self):
        """安全解析，彻底杜绝崩溃"""
        try:
            # 1. 先检查文件是否存在、格式是否正确
            if not os.path.exists(self.file_path):
                print(f"⚠️ 文件不存在: {self.file_path}")
                return

            # 2. 区分.doc和.docx，只处理docx
            if self.file_path.lower().endswith(".doc"):
                print(f"⚠️ 不支持.doc格式，请转成.docx: {self.file_path}")
                return

            # 3. 正常解析docx
            doc = Document(self.file_path)
            text_list = []

            # 读取所有段落
            for para in doc.paragraphs:
                if para.text.strip():
                    text_list.append(para.text.strip())

            # 读取所有表格（兼容所有表格结构）
            for table in doc.tables:
                for row in table.rows:
                    for cell in row.cells:
                        cell_txt = cell.text.strip()
                        if cell_txt:
                            text_list.append(cell_txt)

            # 合并所有文本
            self.full_text = "\n".join(text_list)

            # 4. 判断资助对象（增强版，兼容多种表述）
            self.check_subsidy()

            # 5. 提取申请理由（增强版，兼容多种标题格式）
            self.extract_reason()

        except Exception as e:
            # 彻底兜底，任何错误都不崩溃
            print(f"⚠️ 解析文件 {os.path.basename(self.file_path)} 失败: {str(e)}")
            self.full_text = ""
            self.is_subsidy_flag = False
            self.real_reason = ""
            self.reason_count = 0

    def check_subsidy(self):
        """增强版资助判断，兼容多种表格表述"""
        # 匹配所有可能的资助相关表述
        subsidy_keywords = [
            "是否为学生资助对象",
            "是否家庭经济困难",
            "是否为资助对象",
            "是否困难生",
            "是否享受助学金"
        ]
        
        # 只要表格里有相关问题，且答案是"是"，就判定为资助对象
        for keyword in subsidy_keywords:
            if keyword in self.full_text:
                if "是" in self.full_text:
                    self.is_subsidy_flag = True
                    return
        # 兜底：直接匹配"是" + "资助/困难"
        if "是" in self.full_text and any(k in self.full_text for k in ["资助", "困难", "助学金", "贫困"]):
            self.is_subsidy_flag = True

    def extract_reason(self):
        """增强版理由提取，兼容多种标题格式"""
        # 匹配所有可能的理由标题
        reason_patterns = [
            r"申请理由\s*[（\(]\s*不少于\s*100\s*字\s*[）\)]\s*[：:]",
            r"申请理由\s*[：:]",
            r"申请陈述\s*[：:]",
            r"申请原因\s*[：:]"
        ]

        # 按优先级匹配
        for pattern in reason_patterns:
            split_result = re.split(pattern, self.full_text, flags=re.I)
            if len(split_result) >= 2:
                # 提取标题后的内容
                reason = split_result[1].strip()
                # 去除多余空格、换行、制表符
                reason = re.sub(r"\s+", " ", reason)
                # 去除表格末尾的空内容
                reason = re.sub(r"\s*$", "", reason)
                
                self.real_reason = reason
                self.reason_count = len(reason)
                return

        # 兜底：如果没找到标题，取全文有效字数
        clean_text = re.sub(r"\s+", "", self.full_text)
        self.real_reason = clean_text
        self.reason_count = len(clean_text)

    # 外部调用接口
    def is_subsidy(self):
        return self.is_subsidy_flag

    def get_reason_length(self):
        return self.reason_count
