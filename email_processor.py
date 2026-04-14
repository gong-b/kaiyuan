# -*- coding: utf-8 -*-
# 邮件处理器：负责过滤、解析、提取附件
import re
import os
from email.header import decode_header

class EmailProcessor:
    def __init__(self, keywords):
        self.keywords = keywords

    # 过滤邮件：只保留包含关键词的
    def filter_mails(self, mails):
        filtered = []
        for mail in mails:
            subject_clean = mail["subject"].replace(" ", "").replace("　", "")
            if any(k in subject_clean for k in self.keywords):
                # 从主题提取姓名和学号
                name, sid = self._parse_name_sid(mail["subject"])
                mail["name"] = name
                mail["sid"] = sid
                filtered.append(mail)
        return filtered

    # 从主题提取姓名+学号（正则匹配）
    def _parse_name_sid(self, subject):
        s = re.sub(r"\s+", "", subject)
        match = re.search(r"([\u4e00-\u9fa5]{2,}).*?(\d{8,12})", s)
        if match:
            return match.group(1), match.group(2)
        return None, None

    # 提取邮件中的docx附件
    def extract_attachments(self, mail, save_dir):
        attachments = []
        msg = mail["msg_obj"]
        os.makedirs(save_dir, exist_ok=True)

        for part in msg.walk():
            if part.get_content_maintype() == "multipart":
                continue
            filename = part.get_filename()
            if not filename:
                continue
            
            # 解码文件名
            filename, encoding = decode_header(filename)[0]
            if isinstance(filename, bytes):
                filename = filename.decode(encoding or "utf-8", "replace")
            
            # 只保留docx
            if filename.lower().endswith(".docx"):
                file_path = os.path.join(save_dir, filename)
                with open(file_path, "wb") as f:
                    f.write(part.get_payload(decode=True))
                attachments.append(file_path)
        
        return attachments
