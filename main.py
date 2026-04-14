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

# 日志配置
class SafeLogFilter(logging.Filter):
    def filter(self, record: logging.LogRecord):
        try:
            record.msg = str(record.msg).encode('utf-8', errors='replace').decode('utf-8')
        except:
            pass
        return True

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s",
    handlers=[
        logging.FileHandler(DATA_DIR / "processing.log", encoding="utf-8"),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)
logger.addFilter(SafeLogFilter())

def parse_subject_pattern(subject: str) -> tuple[str, str] | tuple[None, None]:
    """
    放宽版主题解析：
    1. 允许空格和任意干扰字符
    2. 提取连续2-4位的中文作为姓名
    3. 提取连续的数字作为学号
    4. 只要包含 '报名' 或 '申请' 即可
    """
    # 移除所有空格
    clean_subject = re.sub(r"\s+", "", subject)
    
    # 提取学号 (一般为8-10位数字)
    id_match = re.search(r"(\d{6,12})", clean_subject)
    # 提取姓名 (2-4位中文)
    name_match = re.search(r"([\u4e00-\u9fa5]{2,4})", clean_subject)
    
    # 检查是否包含核心关键字（容错：只要有报名或申请字样）
    is_valid_type = any(kw in clean_subject for kw in ["报名", "申请", "书法", "班"])

    if id_match and name_match and is_valid_type:
        return name_match.group(1), id_match.group(1)
    
    return None, None

def parse_subject(msg: Message) -> str:
    """安全解码邮件主题"""
    decoded_parts: list[str] = []
    for part, charset in decode_header(msg.get("Subject", "")):
        try:
            if isinstance(part, bytes):
                if charset:
                    decoded = part.decode(charset, errors='replace')
                else:
                    decoded = part.decode('utf-8', errors='replace')
            else:
                decoded = str(part)
            decoded_parts.append(decoded)
        except Exception:
            decoded_parts.append("[解码失败]")
    return "".join(decoded_parts)

def main():
    email_processor = EmailProcessor()
    new_hongji = read_student_list(str(NEW_HONGJI_FILE))
    last_year = read_student_list(str(LAST_YEAR_FILE))

    admitted: list[dict[str, str]] = []
    rejected: list[dict[str, str]] = []
    candidates: list[tuple[str, str, datetime]] = []

    try:
        start_date_str = os.environ.get("START_DATE", "01-Mar-2025")
        start_date = datetime.strptime(start_date_str, "%d-%b-%Y")
    except:
        start_date = datetime(2025, 3, 1)

    try:
        with SecureIMAPClient() as client:
            for uid, msg in client.fetch_emails():
                # 1. 时间校验
                try:
                    date_str = msg.get("Date")
                    recv_date = parsedate_to_datetime(date_str)
                except:
                    continue

                if recv_date is None or recv_date < start_date.replace(tzinfo=recv_date.tzinfo):
                    continue

                # 2. 主题解析
                subject = parse_subject(msg)
                name, student_id = parse_subject_pattern(subject)

                if not student_id or not name:
                    rejected.append({
                        "学号": "未知", "姓名": "未知", "原主题": subject,
                        "原因": "主题格式无法识别（需包含姓名、学号及报名意图）"
                    })
                    continue

                # 3. 附件提取 (针对所有匹配成功的邮件，包括新鸿基人群)
                # 这样做解决了特定人群被跳过附件提取的问题
                attachments = email_processor.save_attachments(msg, student_id, name)
                
                # 4. 判定逻辑
                # A. 去年已录取
                if student_id in last_year:
                    rejected.append({"学号": student_id, "姓名": name, "原因": "去年已录取"})
                    continue

                # B. 新鸿基人群 (直接录取，但已在上面保存了附件)
                if student_id in new_hongji:
                    admitted.append({"学号": student_id, "姓名": name, "备注": "新鸿基直接录取"})
                    continue

                # C. 普通候选人：检查DOCX附件
                docx_files = [a for a in attachments if a.suffix.lower() == ".docx"]
                if not docx_files:
                    rejected.append({"学号": student_id, "姓名": name, "原因": "缺少DOCX报名表附件"})
                    continue

                # 解析DOCX内容
                try:
                    docx_info = parse_docx(str(docx_files[0]))
                    if not docx_info["is_supported"]:
                        rejected.append({"学号": student_id, "姓名": name, "原因": "非资助对象"})
                    elif docx_info["reason_length"] < 95:
                        rejected.append({"学号": student_id, "姓名": name, "原因": f"申请理由字数不足({docx_info['reason_length']})"})
                    else:
                        candidates.append((student_id, name, recv_date))
                except:
                    rejected.append({"学号": student_id, "姓名": name, "原因": "附件解析失败"})

        # 处理名额限制
        remaining = 25 - len(admitted)
        if remaining > 0:
            candidates.sort(key=lambda x: x[2]) # 按时间先后排序
            for student_id, name, _ in candidates[:remaining]:
                admitted.append({"学号": student_id, "姓名": name, "备注": "择优录取"})
            for student_id, name, _ in candidates[remaining:]:
                rejected.append({"学号": student_id, "姓名": name, "原因": "名额已满"})

        save_results(admitted, rejected)
        logger.info(f"处理完成，录取{len(admitted)}人，拒绝{len(rejected)}人")

    except Exception as e:
        logger.error(f"严重错误: {e}", exc_info=True)

if __name__ == "__main__":
    main()
