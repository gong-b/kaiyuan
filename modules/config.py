class Config:
    IMAP_HOST = "imap.zju.edu.cn"
    IMAP_PORT = 993
    MIN_REASON_LENGTH = 95

    SUBJECT_PATTERN = r"""
        ^\s*
        ([()（）\[\]【】\{\}｛｝])?
        ([\u4e00-\u9fa5]{2,})
        \s*
        (\d+)
        \s*
        书法班报名申请
        ([)）\]\】\}\｝])?
        \s*$
    """
