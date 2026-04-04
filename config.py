import os

# 邮箱配置
EMAIL_HOST = "imap.zju.edu.cn"
EMAIL_PORT = 993
EMAIL_USER = os.getenv("EMAIL_USER")
EMAIL_PASS = os.getenv("EMAIL_PASS")

# 文件路径配置（修复缺失的关键项）
DATA_DIR = "data"
ADMITTED_FILE = "录取名单.xlsx"
REJECTED_FILE = "拒绝名单.xlsx"

# 规则配置
MAX_ACCEPT = 25
REASON_MIN_WORDS = 95

# 自动创建目录
os.makedirs(DATA_DIR, exist_ok=True)
