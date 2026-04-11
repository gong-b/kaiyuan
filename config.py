from pathlib import Path
import os

# ==================== 邮箱配置（建议通过 Streamlit Secrets 配置）====================
IMAP_HOST = os.getenv("IMAP_HOST", "imap.zju.edu.cn")
IMAP_PORT = int(os.getenv("IMAP_PORT", 993))
EMAIL_USER = os.getenv("EMAIL_USER", "zzbgs@zju.edu.cn")  # 敏感信息用环境变量
EMAIL_PASSWORD = os.getenv("EMAIL_PASSWORD", "")  # Streamlit Secrets 中配置

# ==================== 文件路径配置（适配 Streamlit 运行路径）====================
# 获取项目根目录（兼容本地/Streamlit 环境）
BASE_DIR = Path(__file__).parent
DATA_DIR = BASE_DIR / "data"
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

# ==================== Streamlit 配置 ====================
STREAMLIT_PAGE_TITLE = "书法班报名审核系统"
STREAMLIT_PAGE_ICON = "📝"
