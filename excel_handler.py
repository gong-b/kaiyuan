import pandas as pd
from pathlib import Path
import logging

logger = logging.getLogger(__name__)

def read_student_list(file_path: str):
    try:
        if not Path(file_path).exists(): return set()
        df_dict = pd.read_excel(file_path, sheet_name=None, engine="openpyxl")
        for df in df_dict.values():
            df.columns = [str(c).strip() for c in df.columns]
            id_cols = [c for c in df.columns if "学号" in c or "ID" in c.upper()]
            if id_cols:
                ids = df[id_cols[0]].astype(str).str.strip()
                return {sid for sid in ids if sid.lower() not in ["nan", "none", ""]}
        return set()
    except Exception as e:
        logger.error(f"读取失败: {e}")
        return set()

def save_results(admitted, rejected):
    from config import ADMITTED_FILE, REJECTED_FILE
    if admitted: pd.DataFrame(admitted).to_excel(ADMITTED_FILE, index=False)
    if rejected: pd.DataFrame(rejected).to_excel(REJECTED_FILE, index=False)
    logger.info("结果已保存")
