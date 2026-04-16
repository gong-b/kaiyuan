from pathlib import Path
import os

# 基础目录
DATA_DIR = Path("data")
DATA_DIR.mkdir(exist_ok=True, parents=True)

# 临时文件路径（由 app.py 写入，main.py 读取）
NEW_HONGJI_FILE = DATA_DIR / "new_hongji.xlsx"
LAST_YEAR_FILE = DATA_DIR / "last_year.xlsx"
ADMITTED_FILE = DATA_DIR / "admitted_students.xlsx"
REJECTED_FILE = DATA_DIR / "rejected_students.xlsx"

# 附件存储
PDF_DIR = DATA_DIR / "pdfs"
PDF_DIR.mkdir(exist_ok=True, parents=True)

# 邮箱配置
IMAP_HOST = "imap.zju.edu.cn"
IMAP_PORT = 993
