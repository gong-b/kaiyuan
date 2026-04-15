import pandas as pd
from pathlib import Path
import logging
from functools import lru_cache

logger = logging.getLogger(__name__)

@lru_cache(maxsize=4)
def read_student_list(file_path: str):
    """缓存读取学生名单，避免重复IO"""
    try:
        path = Path(file_path)
        if not path.exists():
            return set()
        
        # 流式读取Excel，减少内存占用
        df_dict = pd.read_excel(
            path, 
            sheet_name=None, 
            engine="openpyxl",
            chunksize=1000  # 分块读取
        )
        
        all_ids = set()
        for sheet_name, chunk in df_dict.items():
            chunk.columns = [str(c).strip() for c in chunk.columns]
            id_cols = [c for c in chunk.columns if "学号" in c or "ID" in c.upper()]
            if id_cols:
                ids = chunk[id_cols[0]].astype(str).str.strip()
                all_ids.update({sid for sid in ids if sid.lower() not in ["nan", "none", ""]})
        
        return all_ids
    except Exception as e:
        logger.error(f"读取失败: {e}")
        return set()

def save_results(admitted, rejected):
    from config import ADMITTED_FILE, REJECTED_FILE
    
    # 流式保存Excel，避免内存溢出
    def save_df(data, path):
        if not data:
            return
        # 分块保存
        chunk_size = 1000
        chunks = [data[i:i+chunk_size] for i in range(0, len(data), chunk_size)]
        with pd.ExcelWriter(path, engine="openpyxl") as writer:
            for i, chunk in enumerate(chunks):
                df = pd.DataFrame(chunk)
                if i == 0:
                    df.to_excel(writer, index=False)
                else:
                    df.to_excel(writer, index=False, header=False, startrow=writer.sheets['Sheet1'].max_row)
    
    save_df(admitted, ADMITTED_FILE)
    save_df(rejected, REJECTED_FILE)
    logger.info("结果已保存")
