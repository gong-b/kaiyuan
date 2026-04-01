import pandas as pd
from typing import Any
from config import ADMITTED_FILE, REJECTED_FILE

def read_student_list(file_path: str) -> set[str]:
    """修复学号列匹配问题"""
    try:
        # 动态匹配包含"学号"的列名
        df: pd.DataFrame = pd.read_excel(file_path, sheet_name="sheet1")# type:ignore
        id_col: list[str] = [col for col in df.columns.tolist() if "学号" in col]

        return set(df[id_col[0]].astype(str).tolist())
    except Exception as e:
        print(f"读取失败: {file_path} - 错误信息: {e}")
        return set()
def save_results(admitted: list[Any], rejected: list[Any]):
    """保存录取结果"""
    pd.DataFrame(admitted).to_excel(ADMITTED_FILE, sheet_name="Sheet1", index=False) # type: ignore
    pd.DataFrame(rejected).to_excel(REJECTED_FILE, sheet_name="Sheet1", index=False) # type: ignore