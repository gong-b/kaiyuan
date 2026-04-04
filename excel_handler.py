import pandas as pd

class ExcelHandler:
    def read_student_ids(self, path):
        try:
            df = pd.read_excel(path, dtype=str)
            return set(df.iloc[:, 0].dropna().astype(str).str.strip())
        except:
            return set()

    def write_accept(self, data):
        pd.DataFrame(data, columns=["学号", "姓名", "状态"]).to_excel("录取名单.xlsx", index=False)

    def write_reject(self, data):
        pd.DataFrame(data, columns=["学号", "姓名", "原因"]).to_excel("拒绝名单.xlsx", index=False)
