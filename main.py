import os, logging, re
from email.utils import parsedate_to_datetime
from config import *
from email_client import SecureIMAPClient
from email_processor import EmailProcessor
from docx_parser import parse_docx
from excel_handler import read_student_list, save_results

logging.basicConfig(level=logging.INFO, format="%(asctime)s - %(message)s")
logger = logging.getLogger(__name__)

def parse_subject_pattern(subject: str):
    """从主题提取姓名和学号"""
    clean_s = re.sub(r"\s+", "", subject)
    id_match = re.search(r"(\d{6,12})", clean_s)
    name_match = re.search(r"([\u4e00-\u9fa5]{2,4})", clean_s)
    if id_match and name_match:
        return name_match.group(1), id_match.group(1)
    return None, None

def run_task():
    # 获取环境变量（由app.py传入）
    user = os.environ.get("EMAIL_USER")
    pwd = os.environ.get("EMAIL_PASSWORD")
    start_date = os.environ.get("START_DATE")
    
    logger.info("开始读取本地Excel名单...")
    new_hongji = read_student_list(str(NEW_HONGJI_FILE))
    last_year = read_student_list(str(LAST_YEAR_FILE))
    
    admitted, rejected, candidates = [], [], []
    processor = EmailProcessor()

    with SecureIMAPClient() as client: # 注意：确保email_client已适配环境变量读取
        logger.info(f"正在从 {start_date} 起搜索邮件...")
        for uid, msg in client.fetch_emails():
            from main_logic_helper import parse_subject # 使用你原有的解码函数
            subject = parse_subject(msg)
            name, sid = parse_subject_pattern(subject)
            
            if not (name and sid): continue
            
            # 1. 去年录取判定
            if sid in last_year:
                rejected.append({"学号": sid, "姓名": name, "原因": "去年已录取过"})
                continue
            
            # 2. 新鸿基判定
            if sid in new_hongji:
                admitted.append({"学号": sid, "姓名": name, "备注": "新鸿基推荐直接录取"})
                continue

            # 3. 普通审核：下载并解析DOCX
            attachments = processor.save_attachments(msg, sid, name)
            docx_files = [a for a in attachments if a.suffix.lower() == ".docx"]
            
            if not docx_files:
                rejected.append({"学号": sid, "姓名": name, "原因": "未发现DOCX申请表"})
                continue

            try:
                res = parse_docx(str(docx_files[0]))
                if not res["is_supported"]:
                    rejected.append({"学号": sid, "姓名": name, "原因": "非资助对象"})
                elif res["reason_length"] < 95:
                    rejected.append({"学号": sid, "姓名": name, "原因": f"理由字数不足({res['reason_length']})"})
                else:
                    recv_date = parsedate_to_datetime(msg.get("Date"))
                    candidates.append((sid, name, recv_date))
            except Exception as e:
                rejected.append({"学号": sid, "姓名": name, "原因": f"文件解析异常: {str(e)}"})

    # 4. 排序并录取剩余名额 (总额25)
    remaining_slots = 25 - len(admitted)
    candidates.sort(key=lambda x: x[2]) # 按时间顺序
    
    for sid, n, _ in candidates[:remaining_slots]:
        admitted.append({"学号": sid, "姓名": n, "备注": "普通候补录取"})
    for sid, n, _ in candidates[remaining_slots:]:
        rejected.append({"学号": sid, "姓名": n, "原因": "名额已满"})

    save_results(admitted, rejected)
    logger.info("审核完成，名单已生成。")

if __name__ == "__main__":
    run_task()
