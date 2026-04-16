from pathlib import Path
import os

# 动态获取环境变量，兼容本地测试
IMAP_HOST = os.environ.get("IMAP_HOST", "imap.zju.edu.cn")
IMAP_PORT = int(os.environ.get("IMAP_PORT", 993))
EMAIL_USER = os.environ.get("EMAIL_USER", "zzbgs@zju.edu.cn")
EMAIL_PASSWORD = os.environ.get("EMAIL_PASSWORD", "") # 通过 app.py 设置

DATA_DIR = Path("data")
DATA_DIR.mkdir(exist_ok=True, parents=True)

NEW_HONGJI_FILE = DATA_DIR / "new_hongji.xlsx"
LAST_YEAR_FILE = DATA_DIR / "last_year.xlsx"
ADMITTED_FILE = DATA_DIR / "admitted_students.xlsx"
REJECTED_FILE = DATA_DIR / "rejected_students.xlsx"
PDF_DIR = DATA_DIR / "pdfs"
PDF_DIR.mkdir(exist_ok=True, parents=True)
