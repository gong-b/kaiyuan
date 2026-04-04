import os

# 邮箱配置
EMAIL_HOST = "imap.zju.edu.cn"
EMAIL_PORT = 993
EMAIL_USER = os.getenv("EMAIL_USER")
EMAIL_PASS = os.getenv("EMAIL_PASS")

# 文件路径配置（全部补齐，无任何缺失）
DATA_DIR = "data"
PDF_DIR = "data/pdfs"
ADMITTED_FILE = "录取名单.xlsx"
REJECTED_FILE = "拒绝名单.xlsx"

# 规则配置
MAX_ACCEPT = 25
REASON_MIN_WORDS = 95

# 自动创建所有文件夹
os.makedirs(DATA_DIR, exist_ok=True)
os.makedirs(PDF_DIR, exist_ok=True)
