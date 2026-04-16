import os, logging
from email.utils import parsedate_to_datetime
from config import *
from email_client import SecureIMAPClient
from email_processor import EmailProcessor
from docx_parser import parse_docx
from excel_handler import read_student_list, save_results

def run_task():
    # 从环境变量获取 UI 传递的参数
    user = os.environ.get("EMAIL_USER")
    pwd = os.environ.get("EMAIL_PASSWORD")
    start_date = os.environ.get("START_DATE") 
    folder = os.environ.get("TARGET_FOLDER", "开源课堂")

    new_hongji = read_student_list(NEW_HONGJI_FILE)
    last_year = read_student_list(LAST_YEAR_FILE)
    
    admitted, rejected, candidates = [], [], []
    processor = EmailProcessor()

    with SecureIMAPClient(user, pwd, folder) as client:
        for uid, msg in client.fetch_emails(start_date):
            from main_logic_helper import parse_subject_pattern, parse_subject # 假设你现有的正则
            subject = parse_subject(msg)
            name, student_id = parse_subject_pattern(subject)
            
            if not (name and student_id): continue
            
            # 逻辑：去年录取过 -> 拒绝
            if student_id in last_year:
                rejected.append({"学号": student_id, "姓名": name, "原因": "去年已录取"})
                continue
            
            # 逻辑：新鸿基 -> 直接录取
            if student_id in new_hongji:
                admitted.append({"学号": student_id, "姓名": name, "备注": "新鸿基直接录取"})
                continue

            # 普通学生：检查附件
            attachments = processor.save_attachments(msg, student_id, name)
            docx_files = [a for a in attachments if a.suffix.lower() == ".docx"]
            
            if not docx_files:
                rejected.append({"学号": student_id, "姓名": name, "原因": "缺申请表"})
                continue

            try:
                info = parse_docx(str(docx_files[0]))
                if not info["is_supported"]:
                    rejected.append({"学号": student_id, "姓名": name, "原因": "非资助对象"})
                elif info["reason_length"] < 95:
                    rejected.append({"学号": student_id, "姓名": name, "原因": f"理由不足({info['reason_length']}字)"})
                else:
                    recv_date = parsedate_to_datetime(msg.get("Date"))
                    candidates.append((student_id, name, recv_date))
            except:
                rejected.append({"学号": student_id, "姓名": name, "原因": "文件解析错误"})

    # 排序名额
    remaining = 25 - len(admitted)
    candidates.sort(key=lambda x: x[2]) # 按时间先后
    for sid, n, _ in candidates[:remaining]: admitted.append({"学号": sid, "姓名": n, "备注": "普通录取"})
    for sid, n, _ in candidates[remaining:]: rejected.append({"学号": sid, "姓名": n, "原因": "名额已满"})

    save_results(admitted, rejected)

if __name__ == "__main__":
    run_task()
