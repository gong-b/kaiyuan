import pandas as pd
import config

class ExcelHandler:
    def read_student_ids(self, file_path):
        try:
            df = pd.read_excel(file_path)
            for col in df.columns:
                if "学号" in str(col):
                    return [str(x).strip() for x in df[col] if pd.notna(x)]
        except:
            return []
        return []

    def write_accept(self, data):
        df = pd.DataFrame(data, columns=["学号", "姓名", "状态"])
        df.to_excel(config.ADMITTED_FILE, index=False)

    def write_reject(self, data):
        df = pd.DataFrame(data, columns=["学号", "姓名", "拒绝原因"])
        df.to_excel(config.REJECTED_FILE, index=False)
