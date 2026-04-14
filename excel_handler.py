import pandas as pd
from typing import Any, Set
from pathlib import Path
import logging

logger = logging.getLogger(__name__)

def read_student_list(file_path: str) -> Set[str]:
    """读取学生名单，返回学号集合（增强容错）"""
    try:
        file_path = Path(file_path)
        if not file_path.exists():
            logger.warning(f"文件不存在: {file_path}")
            return set()

        # 兼容不同sheet名称
        df = pd.read_excel(
            file_path,
            sheet_name=None,  # 读取所有sheet
            engine="openpyxl"
        )
        
        # 优先取第一个sheet
        sheet_name = next(iter(df.keys())) if df else "Sheet1"
        df = pd.read_excel(file_path, sheet_name=sheet_name, engine="openpyxl")

        # 查找包含"学号"的列（兼容不同列名）
        id_cols = [col for col in df.columns if "学号" in str(col)]
        if not id_cols:
            logger.warning(f"文件无学号列: {file_path}")
            return set()

        # 提取学号并转为字符串（去重 + 空值过滤）
        student_ids = df[id_cols[0]].astype(str).dropna().str.strip()
        student_ids = set(student_ids[student_ids != "nan"])
        
        logger.info(f"读取学生名单成功: {file_path} - 共{len(student_ids)}个学号")
        return student_ids
    except Exception as e:
        logger.error(f"读取学生名单失败: {file_path} - {str(e)}")
        return set()


def save_results(admitted: list[Any], rejected: list[Any]):
    """保存录取/拒绝结果（修复Excel写入）"""
    try:
        from config import ADMITTED_FILE, REJECTED_FILE

        # 录取名单
        if admitted:
            df_admitted = pd.DataFrame(admitted)
            # 确保列顺序
            df_admitted = df_admitted[["学号", "姓名", "备注"]] if "备注" in df_admitted.columns else df_admitted[["学号", "姓名"]]
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
