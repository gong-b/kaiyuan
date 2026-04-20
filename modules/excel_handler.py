import pandas as pd
from io import BytesIO

class ExcelHandler:
    @staticmethod
    def read_student_list(uploaded):
        df = pd.read_excel(uploaded, sheet_name=0)
        cols = [c for c in df.columns if "学号" in str(c)]
        if not cols:
            return set()
        return set(df[cols[0]].astype(str).str.strip().tolist())

    @staticmethod
    def to_excel_bytes(data):
        out = BytesIO()
        df = pd.DataFrame(data)
        with pd.ExcelWriter(out, engine="openpyxl") as w:
            df.to_excel(w, index=False)
        return out.getvalue()
