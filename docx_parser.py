from docx import Document
import re
from typing import Dict, Any
import logging
from pathlib import Path

logger = logging.getLogger(__name__)

def parse_docx(filepath: str | Path) -> Dict[str, Any]:
    """解析DOCX表格，提取资助对象和申请理由（增强容错）"""
    result: Dict[str, Any] = {
        "is_supported": False,
        "reason_length": 0
    }

    try:
        filepath = Path(filepath)
        if not filepath.exists():
            logger.warning(f"DOCX文件不存在: {filepath}")
            return result
        
        # 尝试打开文档（处理损坏文件）
        try:
            doc = Document(filepath)
        except Exception as e:
            logger.error(f"打开DOCX文件失败: {filepath} - {e}")
            return result

        if not doc.tables:
            logger.warning(f"DOCX文件无表格: {filepath}")
            return result

        # 遍历所有表格和行（增强匹配）
        support_flag = False
        reason_text = ""
        
        for table in doc.tables:
            for row in table.rows:
                try:
                    # 遍历单元格对
                    for i in range(len(row.cells)):
                        cell_text = row.cells[i].text.strip()
                        if not cell_text:
                            continue
                            
                        # 匹配资助对象
                        if any(keyword in cell_text for keyword in ["是否为学生资助对象", "资助对象", "贫困生"]):
                            # 查找下一个单元格
                            if i + 1 < len(row.cells):
                                next_cell = row.cells[i+1].text.strip()
                                support_flag = any(yes in next_cell for yes in ["是", "确认", "√", "对"])
                                support_flag = support_flag and not any(no in next_cell for no in ["否", "不是", "×"])
                        
                        # 匹配申请理由
                        if any(keyword in cell_text for keyword in ["申请理由", "申请原因", "申请说明"]):
                            # 提取当前单元格或后续单元格的理由
                            reason_parts = []
                            # 读取当前单元格剩余内容
                            reason_parts.append(cell_text)
                            # 读取后续单元格
                            for j in range(i+1, len(row.cells)):
                                reason_parts.append(row.cells[j].text.strip())
                            # 读取下一行（如果有）
                            try:
                                next_row = table.rows[table.rows.index(row) + 1]
                                reason_parts.extend([c.text.strip() for c in next_row.cells])
                            except:
                                pass
                            
                            reason_text = "".join(reason_parts)
                except Exception as e:
                    logger.debug(f"解析表格行失败: {e}")
                    continue

        # 清理理由文本
        reason_text = re.sub(r"\s+", "", reason_text)  # 移除所有空格
        reason_text = re.sub(r"[^\u4e00-\u9fa5]", "", reason_text)  # 只保留中文
        reason_length = len(reason_text)

        # 更新结果
        result["is_supported"] = support_flag
        result["reason_length"] = reason_length
        
        logger.info(f"DOCX解析完成: {filepath} - 资助对象:{support_flag}, 理由长度:{reason_length}")
        return result

    except Exception as e:
        logger.error(f"文档解析失败: {filepath} - {str(e)}")
        return result
