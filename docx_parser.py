import zipfile
import re
import logging
import xml.etree.ElementTree as ET
from typing import Dict, Any

logger = logging.getLogger(__name__)

def parse_docx(filepath: str) -> Dict[str, Any]:
    """无需 python-docx 库，直接解析 docx 的 XML 内容"""
    result = {"is_supported": False, "reason_length": 0}
    try:
        with zipfile.ZipFile(filepath) as zf:
            # docx 的文字内容存储在 word/document.xml 中
            xml_content = zf.read('word/document.xml')
            tree = ET.fromstring(xml_content)
            
            # 定义 Office Open XML 的命名空间
            ns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
            
            # 提取所有文本节点的内容
            all_text_elements = tree.findall('.//w:t', ns)
            full_text = "".join([t.text for t in all_text_elements if t.text])

            # 1. 判定资助对象
            # 逻辑：查找关键字，并观察其后 10 个字符内是否有“是”
            if "是否为学生资助对象" in full_text:
                target_area = full_text.split("是否为学生资助对象")[1][:15]
                if "是" in target_area or "为" in target_area:
                    if "不是" not in target_area:
                        result["is_supported"] = True

            # 2. 判定申请理由字数
            # 逻辑：提取“申请理由”和“导师意见”或“学院意见”之间的部分
            if "申请理由" in full_text:
                try:
                    reason_part = full_text.split("申请理由")[1]
                    # 截断到下一个板块
                    for end_kw in ["导师", "学院", "日期"]:
                        if end_kw in reason_part:
                            reason_part = reason_part.split(end_kw)[0]
                            break
                    # 只统计中文字数
                    chinese_chars = re.findall(r'[\u4e00-\u9fa5]', reason_part)
                    result["reason_length"] = len(chinese_chars)
                except:
                    result["reason_length"] = 0
                    
        return result
    except Exception as e:
        logger.error(f"解析 docx 失败: {e}")
        return result
