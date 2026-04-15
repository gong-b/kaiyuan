import pandas as pd
from pathlib import Path
import logging
from functools import lru_cache

logger = logging.getLogger(__name__)

@lru_cache(maxsize=4)
def read_student_list(file_path):
    try:
        path = Path(file_path)
        if not path.exists():
            return set()

        df = pd.read_excel(path, engine="openpyxl", sheet_name=None)
        all_ids = set()

        for sheet_name, sheet_data in df.items():
            sheet_data.columns = [str(c).strip() for c in sheet_data.columns]
            id_cols = [c for c in sheet_data.columns if "学号" in c or "ID" in c.upper()]
            if id_cols:
                ids = sheet_data[id_cols[0]].astype(str).str.strip()
                all_ids.update(i for i in ids if i and i.lower() not in ["nan", "none", ""])
        return all_ids
    except Exception as e:
        logger.error(f"读取失败: {e}")
        return set()

def save_results(admitted, rejected):
    from config import ADMITTED_FILE, REJECTED_FILE
    try:
        if admitted:
            pd.DataFrame(admitted).to_excel(ADMITTED_FILE, index=False, engine="openpyxl")
        if rejected:
            pd.DataFrame(rejected).to_excel(REJECTED_FILE, index=False, engine="openpyxl")
        logger.info("结果已保存")
    except Exception as e:
        logger.error(f"保存失败: {e}")
