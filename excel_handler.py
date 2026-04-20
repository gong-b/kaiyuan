import pandas as pd
import logging
from typing import Set, List, Dict
from io import BytesIO

logger = logging.getLogger(__name__)

class ExcelHandler:
    """Excel文件读写处理类"""
    
    @staticmethod
    def read_student_list(uploaded_file) -> Set[str]:
        """从Streamlit上传的Excel读取学号列表"""
        try:
            df = pd.read_excel(uploaded_file, sheet_name="sheet1")
            # 匹配包含"学号"的列
            id_cols = [col for col in df.columns if "学号" in col]
            if not id_cols:
                logger.warning("Excel中未找到'学号'列")
                return set()
            # 转为字符串去重
            return set(df[id_cols[0]].astype(str).tolist())
        except Exception as e:
            logger.error(f"读取Excel失败: {str(e)}")
            return set()
    
    @staticmethod
    def to_csv_bytes(data: List[Dict]) -> bytes:
        """将数据转为CSV字节流（用于Streamlit下载）"""
        if not data:
            return b""
        df = pd.DataFrame(data)
        # 导出为UTF-8带BOM的CSV（兼容Excel）
        return df.to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig")
