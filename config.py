from pathlib import Path

class Config:
    """配置类（统一管理所有配置）"""
    # 邮箱基础配置（Streamlit中改为手动输入/上传）
    IMAP_HOST = "imap.zju.edu.cn"
    IMAP_PORT = 993
    
    # 文件路径配置（适配Streamlit临时文件）
    DATA_DIR = Path("data")
    PDF_DIR = DATA_DIR / "pdfs"
    
    # 录取规则配置
    MIN_REASON_LENGTH = 100  # 申请理由最低字数
    # 正则表达式（主题匹配）
    SUBJECT_PATTERN = r"""
        ^\s*
        ([()（）\[\]【】\{\}｛｝])?
        ([\u4e00-\u9fa5]{2,})
        \s*
        [+＋-]?
        (\d+)
        \s*
        [+＋-]?
        书法班报名申请
        ([)）\]\】\}\｝])?
        \s*$
    """
    
    @classmethod
    def init_dirs(cls):
        """初始化目录（本地运行时）"""
        cls.DATA_DIR.mkdir(exist_ok=True, parents=True)
        cls.PDF_DIR.mkdir(exist_ok=True, parents=True)
