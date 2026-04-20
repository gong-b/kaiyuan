import pandas as pd
from io import BytesIO

class ExcelHandler:
    @staticmethod
    def read_student_list(uploaded):
        df = pd.read_excel(uploaded, sheet_name=0)
        cols = [c for c in df.columns if "学号" in str(c)]
        if not cols:
            return set()
        # ========== 优化：学号清洗（去空、去小数点、去空格） ==========
        sid_series = df[cols[0]].astype(str).str.strip()
        # 移除数字学号后的.0（如12345678.0 → 12345678）
        sid_series = sid_series.str.replace(".0", "", regex=False)
        # 过滤空值和nan
        sid_series = sid_series[~sid_series.isin(["", "nan"])]
        return set(sid_series.tolist())

    @staticmethod
    def to_excel_bytes(data):
        out = BytesIO()
        df = pd.DataFrame(data)
        with pd.ExcelWriter(out, engine="openpyxl") as w:
            df.to_excel(w, index=False)
        return out.getvalue()
