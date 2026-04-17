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

logging.basicConfig(level=logging.INFO, format="%(asctime)s - %(message)s")
logger = logging.getLogger(__name__)

def parse_subject(msg: Message) -> str:
    subject_raw = msg.get("Subject", "")
    decoded = []
    for part, charset in decode_header(subject_raw):
        if isinstance(part, bytes):
            decoded.append(part.decode(charset or 'utf-8', errors='ignore'))
        else:
            decoded.append(str(part))
    return "".join(decoded)

def parse_subject_pattern(subject: str):
    clean_s = re.sub(r"\s+", "", subject)
    id_match = re.search(r"(\d{6,12})", clean_s)
    name_match = re.search(r"([\u4e00-\u9fa5]{2,4})", clean_s)
    if id_match and name_match:
        return name_match.group(1), id_match.group(1)
    return None, None

def run_task():
    # 获取环境变量参数
    start_dt = datetime.strptime(os.environ.get("START_DATE"), "%Y-%m-%d")
    end_dt = datetime.strptime(os.environ.get("END_DATE"), "%Y-%m-%d")
    
    # 加载名单
    new_hongji = read_student_list(str(NEW_HONGJI_FILE))
    last_year = read_student_list(str(LAST_YEAR_FILE))
    
    # 动态加载黑名单（如果不存在则为空集）
    blacklist_path = DATA_DIR / "blacklist.xlsx"
    blacklist = read_student_list(str(blacklist_path)) if blacklist_path.exists() else set()

    admitted, rejected, candidates = [], [], []
    processor = EmailProcessor()

    with SecureIMAPClient() as client:
        for uid, msg in client.fetch_emails():
            # 1. 邮件时间过滤
            try:
                raw_date = msg.get("Date")
                msg_date = parsedate_to_datetime(raw_date).replace(tzinfo=None)
                if not (start_dt <= msg_date <= end_dt):
                    continue
            except: continue

            # 2. 主题解析
            subject = parse_subject(msg)
            name, sid = parse_subject_pattern(subject)
            if not sid: continue

            # 3. 黑名单过滤 (最高优先级)
            if sid in blacklist:
                rejected.append({"学号": sid, "姓名": name, "原因": "黑名单人员"})
                continue

            # 4. 往年录取过滤
            if sid in last_year:
                rejected.append({"学号": sid, "姓名": name, "原因": "往年已录取"})
                continue
            
            # 5. 新鸿基直录
            if sid in new_hongji:
                admitted.append({"学号": sid, "姓名": name, "备注": "新鸿基直录"})
                continue

            # 6. 普通附件审核
            attachments = processor.save_attachments(msg, sid, name)
            docx_files = [a for a in attachments if a.suffix.lower() == ".docx"]
            
            if not docx_files:
                rejected.append({"学号": sid, "姓名": name, "原因": "缺少申请表"})
                continue

            res = parse_docx(str(docx_files[0]))
            if not res["is_supported"]:
                rejected.append({"学号": sid, "姓名": name, "原因": "非资助对象"})
            elif res["reason_length"] < 95:
                rejected.append({"学号": sid, "姓名": name, "原因": f"理由不足({res['reason_length']}字)"})
            else:
                candidates.append((sid, name, msg_date))

    # 7. 名额排序录取 (总名额25)
    remaining_slots = 25 - len(admitted)
    candidates.sort(key=lambda x: x[2]) # 时间优先原则
    
    for sid, n, _ in candidates[:remaining_slots]:
        admitted.append({"学号": sid, "姓名": n, "备注": "普通录取(时间优先)"})
    for sid, n, _ in candidates[remaining_slots:]:
        rejected.append({"学号": sid, "姓名": n, "原因": "名额已满"})

    save_results(admitted, rejected)

if __name__ == "__main__":
    run_task()
