# -*- coding: utf-8 -*-
# Word解析器：负责读取申请表，提取资助对象、申请理由字数
import re
from docx import Document

class DocxParser:
    def __init__(self, min_reason_length):
        self.min_reason_length = min_reason_length

    # 解析docx文件，提取关键信息
    def parse_application(self, file_path):
        result = {
            "is_supported": False,
            "reason_length": 0,
            "is_valid": False
        }
        try:
            doc = Document(file_path)
            full_text = ""
            for para in doc.paragraphs:
                full_text += para.text.strip()

            # 提取是否为资助对象
            if "是否为学生资助对象" in full_text:
                result["is_supported"] = "是" in full_text and "不是" not in full_text
            
            # 提取申请理由，统计字数
            if "申请理由" in full_text:
                reason = re.sub(r"\s+", "", full_text.split("申请理由")[-1])
                result["reason_length"] = len(reason)
            
            # 校验是否符合要求
            result["is_valid"] = (
                result["is_supported"] 
                and result["reason_length"] >= self.min_reason_length
            )
        except Exception as e:
            pass
        return result
