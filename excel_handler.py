import pandas as pd
from typing import Any, Set
from pathlib import Path
import logging

logger = logging.getLogger(__name__)

def read_student_list(file_path: str) -> Set[str]:
    """读取学生名单，增强对空格、大小写和列名的容错"""
    try:
        file_path = Path(file_path)
        if not file_path.exists():
            return set()

        # 读取所有 Sheet 检查内容
        all_sheets = pd.read_excel(file_path, sheet_name=None, engine="openpyxl")
        for sheet_name, df in all_sheets.items():
            # 1. 清理列名：强制转字符串、去空格、统一转小写
            df.columns = [str(c).strip() for c in df.columns]
            
            # 2. 查找包含“学号”关键字的列（不区分大小写）
            id_cols = [col for col in df.columns if "学号" in col or "ID" in col.upper()]
            
            if id_cols:
                # 3. 提取学号列，转为字符串并去重
                student_ids = df[id_cols[0]].astype(str).str.strip()
                # 过滤掉无效数据
                valid_ids = {sid for sid in student_ids if sid.lower() not in ["nan", "none", "", "学号"] and any(c.isdigit() for c in sid)}
                
                logger.info(f"成功从 {file_path.name} ({sheet_name}) 匹配到列 [{id_cols[0]}]，提取 {len(valid_ids)} 个学号")
                return valid_ids

        logger.warning(f"警告：在文件 {file_path.name} 的所有 Sheet 中均未找到包含“学号”的列。")
        return set()
    except Exception as e:
        logger.error(f"读取 Excel 失败: {file_path.name} - {str(e)}")
        return set()

def save_results(admitted: list[Any], rejected: list[Any]):
    """保存录取/拒绝结果（修复后的写入逻辑）"""
    try:
        from config import ADMITTED_FILE, REJECTED_FILE

        # 录取名单
        if admitted:
            df_admitted = pd.DataFrame(admitted)
            # 确保列顺序
            cols = ["学号", "姓名", "备注"] if "备注" in df_admitted.columns else ["学号", "姓名"]
            df_admitted = df_admitted[cols]
            df_admitted.to_excel(ADMITTED_FILE, sheet_name="录取名单", index=False, engine="openpyxl")

        # 拒绝名单
        if rejected:
            df_rejected = pd.DataFrame(rejected)
            # 确保列顺序
            cols = ["学号", "姓名", "原主题", "原因"] if "原主题" in df_rejected.columns else ["学号", "姓名", "原因"]
            df_rejected = df_rejected[cols]
            df_rejected.to_excel(REJECTED_FILE, sheet_name="拒绝名单", index=False, engine="openpyxl")

        logger.info(f"结果保存成功 - 录取:{len(admitted)}人, 拒绝:{len(rejected)}人")
    except Exception as e:
        logger.error(f"保存结果失败: {str(e)}")
        raise
