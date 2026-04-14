# -*- coding: utf-8 -*-
# Excel处理器：负责读取名单、导出结果
import pandas as pd
from io import BytesIO

class ExcelHandler:
    # 读取名单Excel，提取学号集合
    @staticmethod
    def load_id_list(file):
        if not file:
            return set()
        df = pd.read_excel(file)
        # 自动识别学号列
        col = next((c for c in df.columns if "学号" in str(c)), df.columns[0])
        return set(df[col].astype(str).str.strip().tolist())

    # 导出结果到Excel
    @staticmethod
    def export_to_excel(data, file_name):
        df = pd.DataFrame(data)
        output = BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            df.to_excel(writer, index=False)
        return output.getvalue(), file_name
