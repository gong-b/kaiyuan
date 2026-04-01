from docx import Document
import re
from typing import Dict, Any
import logging
logger = logging.getLogger(__name__)

def parse_docx(filepath: str) -> Dict[str, Any]:
    doc = Document(filepath)
    result: Dict[str, Any] = {
        "is_supported": False,
        "reason_length": 0
    }

    try:
        # 获取文档中第一个表格
        table = doc.tables[0] if doc.tables else None
        if not table:
            return result

        # 遍历表格所有单元格
        for row in table.rows:
            cells = row.cells
            for cell_index, cell in enumerate(cells):
                cell_text: str = cell.text.strip()
                if "是否为学生资助对象" in cell_text:
                    is_supported: bool = (
                        ("是" in cells[cell_index+1].text.strip())
                        or ("为" in cells[cell_index+1].text.strip())
                    ) and ("不是" not in cells[cell_index+1].text.strip())
                    result["is_supported"] = is_supported
                if "申请理由" in cell_text:
                    # 提取该单元格所有段落
                    reason_paragraphs: list[str] = [
                        p.text.strip()
                        for p in cell.paragraphs
                        if p.text.strip()
                    ]
                    reason_text: str = "\n".join(reason_paragraphs)
                    reason_text = re.sub(r"\s+", "", reason_text)
                    result["reason_length"] = len(reason_text.replace(" ", ""))
        return result

    except Exception as e:
        logger.error(f"文档解析失败: {e}")
        return result