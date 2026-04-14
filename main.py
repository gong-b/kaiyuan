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

# 正则表达式：匹配可能被括号包裹的姓名+学号+书法班报名申请
SUBJECT_PATTERN = re.compile(
    r"^\s*"
    r"([()（）\[\]【】\{\}｛｝])?"
    r"([\u4e00-\u9fa5]{2,})"
    r"\s*"
    r"[+＋-]?"
    r"(\d+)"
    r"\s*"
    r"[+＋-]?"
    r"书法班报名申请"
    r"([)）\]\】\}\｝])?"
    r"\s*$"
)

logger = logging.getLogger(__name__)
logger.addFilter(SafeLogFilter())

def parse_subject_pattern(subject: str) -> tuple[str,str]|tuple[None,None]:
    """强化版主题解析，匹配指定的格式"""
    clean_subject = re.sub(r"\s+", "", subject)
    match = SUBJECT_PATTERN.search(clean_subject)
    if match:
        return (match.group(2), match.group(3))
    return (None, None)

def parse_subject(msg: Message) -> str:
    """安全解码邮件主题"""
    decoded_parts: list[str] = []
    for part, charset in decode_header(msg.get("Subject", "")):
        try:
            if charset:
                decoded = part.decode(charset, errors='replace')
            else:
                for encoding in ['utf-8', 'gb18030', 'big5']:
                    try:
                        decoded = part.decode(encoding)
                        break
                    except:
                        continue
                else:
                    decoded = part.decode('utf-8', errors='replace')
            decoded_parts.append(decoded)
        except Exception as e:
            logger.warning(f"主题解码异常: {str(e).encode('utf-8', errors='replace').decode('utf-8')}")
            decoded_parts.append("[解码失败]")
    return "".join(decoded_parts)

def main():
    """主处理流程"""
    email_processor = EmailProcessor()

    # 读取基础数据
    new_hongji = read_student_list(str(NEW_HONGJI_FILE))
    last_year = read_student_list(str(LAST_YEAR_FILE))

    admitted: list[dict[str, str]] = []
    rejected: list[dict[str, str]] = []
    candidates: list[tuple[str, str, datetime]] = []

    # 修复：从环境变量读取日期
    try:
        start_date_str = os.environ.get("START_DATE", "01-Mar-2025")
        start_date = datetime.strptime(start_date_str, "%d-%b-%Y")
    except:
        start_date = datetime(2025, 3, 1)

    try:
        with SecureIMAPClient() as client:
            for uid, msg in client.fetch_emails():
                # 获取邮件接收时间
                try:
                    date_str = msg.get("Date")
                    recv_date = parsedate_to_datetime(date_str)
                except Exception as e:
                    logger.error(f"日期解析失败: {e}")
                    continue

                if recv_date is None or recv_date < start_date.replace(tzinfo=recv_date.tzinfo):
                    logger.warning(f"邮件{uid}时间不符合要求：{recv_date}")
                    continue

                subject = ""
                try:
                    subject = parse_subject(msg)
                    name, student_id = parse_subject_pattern(subject)
                except Exception as e:
                    logger.error(f"主题解析失败: {str(e).encode('utf-8', errors='replace').decode('utf-8')}")
                    student_id, name = None, None

                if not student_id or not name:
                    rejected.append({
                        "学号": "未知",
                        "姓名": "未知",
                        "原主题": f"{subject}",
                        "原因": "主题格式错误（正确示例：薛孜324011234书法班报名申请或者薛孜+3240101517+书法班报名申请）"
                    })
                    continue

                # 新鸿基直接录取
                if student_id in new_hongji:
                    admitted.append({"学号": student_id, "姓名": name, "备注": "新鸿基"})
                    email_processor.save_email_pdf(msg, student_id, name)
                    continue

                # 去年已录取
                if student_id in last_year:
                    rejected.append({
                        "学号": student_id,
                        "姓名": name,
                        "原因": "去年已录取"
                    })
                    continue

                # 处理附件
                attachments = email_processor.save_attachments(msg, student_id, name)
                docx_files = [a for a in attachments if a.suffix == ".docx"]

                if not docx_files:
                    rejected.append({
                        "学号": student_id,
                        "姓名": name,
                        "原因": "缺少DOCX附件"
                    })
                    continue

                # 解析DOCX
                try:
                    docx_info = parse_docx(str(docx_files[0]))
                    if not docx_info["is_supported"]:
                        rejected.append({
                            "学号": student_id,
                            "姓名": name,
                            "原因": "非资助对象"
                        })
                    elif docx_info["reason_length"] < 95:
                        rejected.append({
                            "学号": student_id,
                            "姓名": name,
                            "原因": f"申请理由不足（{docx_info['reason_length']}字）"
                        })
                    else:
                        candidates.append((student_id, name, recv_date))
                except Exception as e:
                    rejected.append({
                        "学号": student_id,
                        "姓名": name,
                        "原因": "附件解析失败"
                    })

        # 处理候补名单
        remaining = 25 - len(admitted)
        if remaining > 0:
            candidates.sort(key=lambda x: x[2])
            for student_id, name, _ in candidates[:remaining]:
                admitted.append({"学号": student_id, "姓名": name, "备注": "非新鸿基"})
            for student_id, name, _ in candidates[remaining:]:
                rejected.append({
                    "学号": student_id,
                    "姓名": name,
                    "原因": "名额已满"
                })

        # 保存结果
        save_results(admitted, rejected)
        logger.info(f"处理完成，录取{len(admitted)}人，拒绝{len(rejected)}人")

    except Exception as e:
        logger.error(f"处理过程中发生严重错误: {e}", exc_info=True)

if __name__ == "__main__":
    main()
