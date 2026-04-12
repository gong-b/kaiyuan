from pathlib import Path
import os

# 邮箱配置（优先读取环境变量，兼容动态传递）
IMAP_HOST = os.environ.get("IMAP_HOST", "imap.zju.edu.cn")
IMAP_PORT = int(os.environ.get("IMAP_PORT", 993))
EMAIL_USER = os.environ.get("EMAIL_USER", "")  # 从Streamlit传递，不硬编码
EMAIL_PASSWORD = os.environ.get("EMAIL_PASSWORD", "")

# 基础目录配置（基于脚本位置的绝对路径，避免相对路径错误）
BASE_DIR = Path(__file__).parent
DATA_DIR = BASE_DIR / "data"
DATA_DIR.mkdir(exist_ok=True, parents=True)  # 自动创建data目录

# ---------------------- 核心：适配你的3个Excel文件名 ----------------------
NEW_HONGJI_FILE = DATA_DIR / "新鸿基名单.xlsx"       # 你的新鸿基推荐名单
LAST_YEAR_FILE = DATA_DIR / "副本去年报名名单.xlsx"  # 你的去年报名名单
BLACKLIST_FILE = DATA_DIR / "黑名单.xlsx"           # 你的黑名单

# 筛选结果输出路径
ADMITTED_FILE = DATA_DIR / "admitted_students.xlsx"  # 录取名单
REJECTED_FILE = DATA_DIR / "rejected_students.xlsx"  # 拒绝名单

# 业务配置（可通过环境变量调整）
ADMISSION_QUOTA = int(os.environ.get("ADMISSION_QUOTA", 25))  # 总录取名额
MIN_REASON_LENGTH = 95  # 申请理由最低字数要求
