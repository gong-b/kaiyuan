import os

# 邮箱配置
EMAIL_HOST = "imap.zju.edu.cn"
EMAIL_PORT = 993
EMAIL_USER = os.getenv("EMAIL_USER")
EMAIL_PASS = os.getenv("EMAIL_PASS")

# 路径
DATA_DIR = "data"
os.makedirs(DATA_DIR, exist_ok=True)

# 规则
MAX_ACCEPT = 25
REASON_MIN_WORDS = 95
