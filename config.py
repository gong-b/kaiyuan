from pathlib import Path
import os

# 邮箱配置（优先读取环境变量，兼容原有硬编码）
IMAP_HOST = os.environ.get("IMAP_HOST", "imap.zju.edu.cn")
IMAP_PORT = int(os.environ.get("IMAP_PORT", 993))
EMAIL_USER = os.environ.get("EMAIL_USER", "zzbgs@zju.edu.cn")  # 环境变量优先
EMAIL_PASSWORD = os.environ.get("EMAIL_PASSWORD", "ww9SVyi9ixmhnqu3")  # 环境变量优先

# 文件路径配置
DATA_DIR = Path("data")
DATA_DIR.mkdir(exist_ok=True, parents=True)

# Excel文件路径
NEW_HONGJI_FILE = DATA_DIR / "2024-2025学年秋冬学期新鸿基推荐学生名单.xlsx"
LAST_YEAR_FILE = DATA_DIR / "24秋冬学期开源课堂人员名单.xlsx"

# 输出文件
ADMITTED_FILE = DATA_DIR / "admitted_students.xlsx"
REJECTED_FILE = DATA_DIR / "rejected_students.xlsx"

# PDF存储路径
PDF_DIR = DATA_DIR / "pdfs"
PDF_DIR.mkdir(exist_ok=True, parents=True)
