import logging
import re
import os
from datetime import datetime, timezone
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

# 正则表达式：增强版主题匹配
SUBJECT_PATTERN = re.compile(
    r"^\s*"
    r"[()（）\[\]【】\{\}｛｝]*"  # 任意数量的括号
    r"([\u4e00-\u9fa5]{2,4})"  # 姓名（2-4个中文字符）
    r"\s*[+＋-—\s]*"  # 分隔符（支持全角/半角/空格/破折号）
    r"(\d{8,12})"  # 学号（8-12位数字）
    r"\s*[+＋-—\s]*"
    r"书法班报名申请"
    r"[()（）\[\]【】\{\}｛｝]*"
    r"\s*$",
    re.IGNORECASE  # 忽略大小写（防止特殊情况）
)

logger = logging.getLogger(__name__)
logger.addFilter(SafeLogFilter())

def parse_subject_pattern(subject: str) -> tuple[str, str] | tuple[None, None]:
    """强化版主题解析"""
    if not subject:
        return None, None
        
    clean_subject = re.sub(r"\s+", "", subject)
    match = SUBJECT_PATTERN.match(clean_subject)  # 使用match而非search，确保整行匹配
    
    if match:
        name = match.group(1).strip()
        student_id = match.group(2).strip()
        return name, student_id
    return None, None

def parse_subject(msg: Message) -> str:
    """安全解码邮件主题"""
    subject = msg.get("Subject", "")
    decoded_parts = []
    
    for part, charset in decode_header(subject):
        try:
            if isinstance(part, bytes):
                # 优先尝试常见编码
                for encoding in [charset, 'utf-8', 'gb18030', 'big5', 'gbk']:
                    if not encoding:
                        continue
                    try:
                        decoded = part.decode(encoding)
                        break
                    except:
                        continue
                else:
                    decoded = part.decode('utf-8', errors='replace')
            else:
                decoded = str(part)
            decoded_parts.append(decoded)
        except Exception as e:
            logger.warning(f"主题解码异常: {str(e)}")
            decoded_parts.append("[解码失败]")
    
    return "".join(decoded_parts)

def main():
    """主处理流程（修复日期/筛选/排序逻辑）"""
    # 初始化处理器
    email_processor = EmailProcessor()
    
    # 读取基础数据
    logger.info("读取学生名单...")
    new_hongji = read_student_list(str(NEW_HONGJI_FILE))
    last_year = read_student_list(str(LAST_YEAR_FILE))
    blacklist = read_student_list(str(BLACKLIST_FILE))  # 新增黑名单
    
    admitted: list[dict[str, str]] = []
    rejected: list[dict[str, str]] = []
    candidates: list[tuple[str, str, datetime]] = []
    
    # 解析日期范围
    try:
        start_date_str = os.environ.get("START_DATE", "01-Mar-2025")
        end_date_str = os.environ.get("END_DATE", datetime.now().strftime("%d-%b-%Y"))
        
        # 转换为带时区的datetime
        start_date = datetime.strptime(start_date_str, "%d-%b-%Y").replace(tzinfo=timezone.utc)
        end_date = datetime.strptime(end_date_str, "%d-%b-%Y").replace(tzinfo=timezone.utc)
        logger.info(f"处理日期范围: {start_date_str} 至 {end_date_str}")
    except Exception as e:
        logger.error(f"日期解析失败，使用默认值: {e}")
        start_date = datetime(2025, 3, 1, tzinfo=timezone.utc)
        end_date = datetime.now(timezone.utc)

    try:
        # 验证基础文件
        if not new_hongji:
            raise RuntimeError("新鸿基名单为空或文件不存在")
        if not last_year:
            raise RuntimeError("去年录取名单为空或文件不存在")
        
        # 获取并处理邮件
        logger.info("连接邮箱并获取邮件...")
        with SecureIMAPClient() as client:
            for uid, msg in client.fetch_emails():
                # 解析邮件接收时间
                recv_date = None
                try:
                    date_str = msg.get("Date")
                    if date_str:
                        recv_date = parsedate_to_datetime(date_str)
                        # 转换为UTC时区（统一比较）
                        if recv_date.tzinfo is None:
                            recv_date = recv_date.replace(tzinfo=timezone.utc)
                        else:
                            recv_date = recv_date.astimezone(timezone.utc)
                except Exception as e:
                    logger.error(f"邮件{uid}日期解析失败: {e}")
                    continue
                
                # 日期过滤（包含开始/结束日期）
                if not recv_date or not (start_date <= recv_date <= end_date):
                    logger.debug(f"邮件{uid}时间不在范围内: {recv_date}")
                    continue
        
                # 解析主题
                subject = parse_subject(msg)
                try:
                    name, student_id = parse_subject_pattern(subject)
                except Exception as e:
                    logger.error(f"邮件{uid}主题解析失败: {e}")
                    name, student_id = None, None
                
                # 验证学号/姓名
                if not student_id or not name:
                    rejected.append({
                        "学号": "未知",
                        "姓名": "未知",
                        "原主题": subject,
                        "原因": "主题格式错误（正确示例：薛孜324011234书法班报名申请）"
                    })
                    continue
                
                # 黑名单过滤
                if student_id in blacklist:
                    rejected.append({
                        "学号": student_id,
                        "姓名": name,
                        "原因": "黑名单用户"
                    })
                    continue
                
                # 新鸿基直接录取
                if student_id in new_hongji:
                    admitted.append({"学号": student_id, "姓名": name, "备注": "新鸿基"})
                    logger.info(f"新鸿基录取: {name}({student_id})")
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
                docx_files = [a for a in attachments if a.suffix.lower() == ".docx"]  # 忽略大小写
                
                if not docx_files:
                    rejected.append({
                        "学号": student_id,
                        "姓名": name,
                        "原因": "缺少DOCX格式申请附件"
                    })
                    continue
                
                # 解析DOCX（只取第一个有效文件）
                try:
                    docx_info = parse_docx(str(docx_files[0]))
                    if not docx_info["is_supported"]:
                        rejected.append({
                            "学号": student_id,
                            "姓名": name,
                            "原因": "非学生资助对象，不符合申请条件"
                        })
                    elif docx_info["reason_length"] < 95:
                        rejected.append({
                            "学号": student_id,
                            "姓名": name,
                            "原因": f"申请理由字数不足（仅{docx_info['reason_length']}字，需≥95字）"
                        })
                    else:
                        # 加入候补名单（带接收时间）
                        candidates.append((student_id, name, recv_date))
                        logger.info(f"加入候补: {name}({student_id})")
                except Exception as e:
                    logger.error(f"解析{student_id}的DOCX失败: {e}")
                    rejected.append({
                        "学号": student_id,
                        "姓名": name,
                        "原因": f"附件解析失败: {str(e)[:50]}..."
                    })
            
        # 处理候补名单（按接收时间升序）
        remaining_quota = ADMISSION_QUOTA - len(admitted)
        logger.info(f"新鸿基录取{len(admitted)}人，剩余名额{remaining_quota}")
        
        if remaining_quota > 0 and candidates:
            # 按接收时间排序（先到先得）
            candidates.sort(key=lambda x: x[2])
            
            # 录取候补
            admit_candidates = candidates[:remaining_quota]
            for student_id, name, _ in admit_candidates:
                admitted.append({"学号": student_id, "姓名": name, "备注": "非新鸿基（候补）"})
                logger.info(f"候补录取: {name}({student_id})")
            
            # 名额已满拒绝
            reject_candidates = candidates[remaining_quota:]
            for student_id, name, _ in reject_candidates:
                rejected.append({
                    "学号": student_id,
                    "姓名": name,
                    "原因": "符合条件但名额已满"
                })
        elif remaining_quota <= 0 and candidates:
            logger.warning("名额已满，所有候补均被拒绝")
            for student_id, name, _ in candidates:
                rejected.append({
                    "学号": student_id,
                    "姓名": name,
                    "原因": "符合条件但名额已满"
                })
        
        # 保存结果
        save_results(admitted, rejected)
        logger.info(f"处理完成 | 录取{len(admitted)}人 | 拒绝{len(rejected)}人")
        
    except Exception as e:
        logger.error(f"处理过程中发生严重错误", exc_info=True)
        raise

if __name__ == "__main__":
    main()
