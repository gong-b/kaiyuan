from pathlib import Path
import os

DATA_DIR = Path("data")
DATA_DIR.mkdir(exist_ok=True)

NEW_HONGJI_FILE = DATA_DIR / "2024-2025学年秋冬学期新鸿基推荐学生名单.xlsx"
LAST_YEAR_FILE = DATA_DIR / "24秋冬学期开源课堂人员名单.xlsx"

ADMITTED_FILE = DATA_DIR / "admitted_students.xlsx"
REJECTED_FILE = DATA_DIR / "rejected_students.xlsx"

IMAP_HOST = "imap.zju.edu.cn"
IMAP_PORT = 993

# 从环境变量读取，不写死
EMAIL_USER = os.getenv("EMAIL_USER", "")
EMAIL_PASSWORD = os.getenv("EMAIL_PASSWORD", "")
