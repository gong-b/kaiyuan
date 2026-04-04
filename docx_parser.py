from docx import Document
import re
import os
import zipfile
import xml.etree.ElementTree as ET

class DocxParser:
    def __init__(self, file_path):
        self.file_path = file_path
        self.full_text = ""
        self.is_subsidy_flag = False
        self.real_reason = ""
        self.reason_count = 0
        self.parse_safely()

    def parse_safely(self):
        """纯Python解析，无系统依赖，兼容.doc/.docx"""
        try:
            if not os.path.exists(self.file_path):
                print(f"⚠️ 文件不存在: {self.file_path}")
                return

            ext = self.file_path.lower().split('.')[-1]

            # 1. 解析.docx（原生支持）
            if ext == "docx":
                self._parse_docx()
            # 2. 解析.doc（纯Python兼容方案，Streamlit云可用）
            elif ext == "doc":
                self._parse_doc_fallback()
            else:
                print(f"⚠️ 不支持的格式: {ext}")
                self.full_text = ""

            # 3. 业务逻辑：判断资助+提取理由
            self.check_subsidy()
            self.extract_reason()

        except Exception as e:
            print(f"⚠️ 解析失败: {str(e)}")
            self.full_text = ""
            self.is_subsidy_flag = False
            self.reason_count = 0

    def _parse_docx(self):
        """解析.docx文件（原生python-docx）"""
        doc = Document(self.file_path)
        text_list = []
        # 读取段落
        for para in doc.paragraphs:
            t = para.text.strip()
            if t:
                text_list.append(t)
        # 读取表格
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    t = cell.text.strip()
                    if t:
                        text_list.append(t)
        self.full_text = "\n".join(text_list)

    def _parse_doc_fallback(self):
        """.doc文件纯Python兼容方案（Streamlit云可用）"""
        try:
            # 尝试用zipfile解析（部分.doc文件是zip格式）
            with zipfile.ZipFile(self.file_path, 'r') as zf:
                # 提取word/document.xml
                if 'word/document.xml' in zf.namelist():
                    with zf.open('word/document.xml') as f:
                        tree = ET.parse(f)
                        root = tree.getroot()
                        # 提取所有文本节点
                        ns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
                        texts = [node.text for node in root.findall('.//w:t', ns) if node.text]
                        self.full_text = "\n".join(texts)
                        return
        except:
            # 兜底：直接读取二进制文本（兼容老版.doc）
            with open(self.file_path, 'rb') as f:
                raw = f.read()
                # 提取可打印字符
                text = ''.join([chr(c) for c in raw if 32 <= c <= 126 or c in (10, 13)])
                self.full_text = text

    def check_subsidy(self):
        """精准判断资助对象（兼容你的申请表格式）"""
        text = self.full_text
        # 优先匹配表格字段
        if "是否为学生资助对象" in text and "是" in text:
            self.is_subsidy_flag = True
        # 兜底匹配
        elif "是" in text and any(k in text for k in ["资助", "困难", "贫困", "助学金"]):
            self.is_subsidy_flag = True
        else:
            self.is_subsidy_flag = False

    def extract_reason(self):
        """精准提取申请理由，只算正文"""
        text = self.full_text
        # 匹配多种理由标题格式
        patterns = [
            r"申请理由\s*[（\(]\s*不少于\s*100\s*字\s*[）\)]\s*[：:]",
            r"申请理由\s*[：:]",
            r"申请陈述\s*[：:]"
        ]
        for p in patterns:
            parts = re.split(p, text, flags=re.I)
            if len(parts) >= 2:
                reason = parts[1].strip()
                reason = re.sub(r"\s+", " ", reason)
                self.real_reason = reason
                self.reason_count = len(reason)
                return
        # 兜底：统计全文有效字数
        clean = re.sub(r"\s+", "", text)
        self.real_reason = clean
        self.reason_count = len(clean)

    # 外部调用接口
    def is_subsidy(self):
        return self.is_subsidy_flag

    def get_reason_length(self):
        return self.reason_count
