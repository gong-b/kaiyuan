from docx import Document
import re
from typing import Dict, Any
import logging
from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path

logger = logging.getLogger(__name__)

def parse_single_docx(filepath: str) -> Dict[str, Any]:
    """解析单个DOCX文件"""
    result: Dict[str, Any] = {
        "is_supported": False,
        "reason_length": 0,
        "filepath": filepath
    }
    try:
        doc = Document(filepath)
        if not doc.tables:
            logger.warning(f"文档无表格: {filepath}")
            return result

        for table in doc.tables:
            for row in table.rows:
                cells = row.cells
                for cell_index, cell in enumerate(cells):
                    if cell_index + 1 >= len(cells):
                        continue
                    cell_text = cell.text.strip()
                    
                    if "是否为学生资助对象" in cell_text:
                        next_cell_text = cells[cell_index+1].text.strip()
                        result["is_supported"] = "是" in next_cell_text and "不是" not in next_cell_text
                    
                    if "申请理由" in cell_text:
                        reason_text = re.sub(r"\s+", "", cell.text.strip())
                        reason_text = re.sub(r"[^\u4e00-\u9fa5]", "", reason_text)
                        result["reason_length"] = len(reason_text)
        
        logger.info(f"DOCX解析完成: {filepath} - 资助对象:{result['is_supported']}, 理由长度:{result['reason_length']}")
        return result
    except Exception as e:
        logger.error(f"文档解析失败: {filepath} - {str(e)}")
        return result

def parse_docx_batch(filepaths: list[str], max_workers: int = 4) -> Dict[str, Dict[str, Any]]:
    """批量解析DOCX文件（多线程）"""
    results = {}
    with ThreadPoolExecutor(max_workers=max_workers) as executor:
        future_to_file = {executor.submit(parse_single_docx, fp): fp for fp in filepaths}
        for future in as_completed(future_to_file):
            fp = future_to_file[future]
            try:
                results[fp] = future.result()
            except Exception as e:
                logger.error(f"批量解析失败: {fp} - {e}")
                results[fp] = {"is_supported": False, "reason_length": 0, "filepath": fp}
    return results
