from docx import Document
import re
from typing import Dict, Any
import logging

logger = logging.getLogger(__name__)

def parse_docx(filepath: str) -> Dict[str, Any]:
    """解析DOCX表格，提取资助对象和申请理由"""
    result: Dict[str, Any] = {
        "is_supported": False,
        "reason_length": 0
    }
    try:
        doc = Document(filepath)
        if not doc.tables:
            logger.warning(f"文档无表格: {filepath}")
            return result

        # 遍历所有表格（不只第一个）
        for table in doc.tables:
            for row in table.rows:
                cells = row.cells
                # 修复单元格索引越界问题
                for cell_index, cell in enumerate(cells):
                    if cell_index + 1 >= len(cells):
                        continue
                    cell_text = cell.text.strip()
                    
                    # 匹配资助对象字段
                    if "是否为学生资助对象" in cell_text:
                        next_cell_text = cells[cell_index+1].text.strip()
                        result["is_supported"] = "是" in next_cell_text and "不是" not in next_cell_text
                    
                    # 匹配申请理由字段
                    if "申请理由" in cell_text:
                        # 提取纯文本并去重空格
                        reason_text = re.sub(r"\s+", "", cell.text.strip())
                        # 移除特殊字符，只保留中文
                        reason_text = re.sub(r"[^\u4e00-\u9fa5]", "", reason_text)
                        result["reason_length"] = len(reason_text)
        
        logger.info(f"DOCX解析完成: {filepath} - 资助对象:{result['is_supported']}, 理由长度:{result['reason_length']}")
        return result
    except Exception as e:
        logger.error(f"文档解析失败: {filepath} - {str(e)}")
        return result
