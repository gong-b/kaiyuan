import logging
import re
import os
from datetime import datetime
from email.header import decode_header
from email.message import Message
from email.utils import parsedate_to_datetime
from config import *
from email_client import SecureIMAPClient
from email_processor import EmailProcessor
from docx_parser import parse_docx
from excel_handler import read_student_list, save_results

logging.basicConfig(level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s")
logger = logging.getLogger(__name__)

def parse_subject_pattern(subject: str) -> tuple[str, str] | tuple[None, None]:
    """放宽的主题解析：包含2-4位中文作为姓名，及连续数字作为学号"""
    clean_subject = re.sub(r"\s+", "", subject)
    id_match = re.search(r"(\d{6,12})", clean_subject)
    name_match = re.search(r"([\u4e00-\u9fa5]{2,4})", clean_subject)
    is_valid = any(kw in clean_subject for kw in ["报名", "申请", "班", "书法"])
    if id_match and name_match and is_valid:
        return name_match.group(1), id_match.group(1)
    return None, None

def parse_subject(msg: Message) -> str:
    """安全解码主题"""
    decoded_parts = []
    subject_raw = msg.get("Subject", "")
    for part, charset in decode_header(subject_raw):
        if isinstance(part, bytes):
            decoded_parts.append(part.decode(charset or 'utf-8', errors='replace'))
        else:
            decoded_parts.append(str(part))
    return "".join(decoded_parts)

def main():
    email_processor = EmailProcessor()
    new_hongji = read_student_list(str(NEW_HONGJI_FILE))
    last_year = read_student_list(str(LAST_YEAR_FILE))
    admitted, rejected, candidates = [], [], []

    try:
        with SecureIMAPClient() as client:
            for uid, msg in client.fetch_emails():
                recv_date = parsedate_to_datetime(msg.get("Date"))
                subject = parse_subject(msg)
                name, student_id = parse_subject_pattern(subject)

                if not student_id:
                    rejected.append({"学号": "未知", "姓名": "未知", "原主题": subject, "原因": "格式错误"})
                    continue

                # 1. 提取所有附件（解决特定人群不被跳过的问题）
                attachments = email_processor.save_attachments(msg, student_id, name)
                
                # 2. 去年已录取判定
                if student_id in last_year:
                    rejected.append({"学号": student_id, "姓名": name, "原因": "去年已录取"})
                    continue

                # 3. 新鸿基判定（直接录取，不再往下解析DOCX）
                if student_id in new_hongji:
                    admitted.append({"学号": student_id, "姓名": name, "备注": "新鸿基直接录取"})
                    continue

                # 4. 普通人：校验 DOCX 附件
                docx_files = [a for a in attachments if a.suffix.lower() == ".docx"]
                if not docx_files:
                    rejected.append({"学号": student_id, "姓名": name, "原因": "缺少DOCX附件"})
                    continue

                try:
                    info = parse_docx(str(docx_files[0]))
                    if not info["is_supported"]:
                        rejected.append({"学号": student_id, "姓名": name, "原因": "非资助对象"})
                    elif info["reason_length"] < 95:
                        rejected.append({"学号": student_id, "姓名": name, "原因": f"理由不足({info['reason_length']}字)"})
                    else:
                        candidates.append((student_id, name, recv_date))
                except:
                    rejected.append({"学号": student_id, "姓名": name, "原因": "文档解析失败"})

        # 处理名额逻辑
        remaining = 25 - len(admitted)
        candidates.sort(key=lambda x: x[2])
        for sid, n, _ in candidates[:remaining]: admitted.append({"学号": sid, "姓名": n, "备注": "普通录取"})
        for sid, n, _ in candidates[remaining:]: rejected.append({"学号": sid, "姓名": n, "原因": "名额已满"})

        save_results(admitted, rejected)
    except Exception as e: logger.error(f"运行失败: {e}")

if __name__ == "__main__": main()
