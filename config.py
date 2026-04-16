from pathlib import Path
import os

# 基础目录配置
DATA_DIR = Path("data")
DATA_DIR.mkdir(exist_ok=True, parents=True)

PDF_DIR = DATA_DIR / "pdfs"
PDF_DIR.mkdir(exist_ok=True, parents=True)

# 输出文件路径
ADMITTED_FILE = DATA_DIR / "admitted_students.xlsx"
REJECTED_FILE = DATA_DIR / "rejected_students.xlsx"

# 邮箱服务器配置
IMAP_HOST = "imap.zju.edu.cn"
IMAP_PORT = 993
