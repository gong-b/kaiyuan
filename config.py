# -*- coding: utf-8 -*-
# 全局配置文件

# 邮箱配置
IMAP_SERVER = "imap.zju.edu.cn"
IMAP_PORT = 993

# 文件夹配置（浙大邮箱路径，已适配子文件夹）
MAIL_FOLDER = "其他文件夹/开源课堂"  # 你的真实文件夹路径
FALLBACK_FOLDER = "INBOX"  # 备用：收件箱

# 时间配置
DEFAULT_START_DATE = "2026-03-01"
DEFAULT_END_DATE = "2026-04-10"

# 审核规则配置
MIN_REASON_LENGTH = 95  # 申请理由最低字数

# 过滤关键词
MAIL_KEYWORDS = ["报名", "开源课堂"]  # 只保留包含这些关键词的邮件
