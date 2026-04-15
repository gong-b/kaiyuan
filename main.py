import logging
import re
import os
import asyncio
from datetime import datetime
from email.header import decode_header
from email.message import Message
from email.utils import parsedate_to_datetime
from config import *
from email_client import AsyncSecureIMAPClient
from email_processor import EmailProcessor
from docx_parser import parse_docx_batch
from excel_handler import read_student_list, save_results

logging.basicConfig(level=logging.WARNING, format="%(asctime)s - %(levelname)s - %(message)s")
logger = logging.getLogger(__name__)

def parse_subject_pattern(subject: str) -> tuple[str, str] | tuple[None, None]:
    """优化正则匹配效率"""
    clean_subject = re.sub(r"\s+", "", subject)
    # 预编译正则提升效率
    id_pattern = re.compile(r"(\d{6,12})")
    name_pattern = re.compile(r"([\u4e00-\u9fa5]{2,4})")
    keyword_pattern = re.compile(r"报名|申请|班|书法")
    
    id_match = id_pattern.search(clean_subject)
    name_match = name_pattern.search(clean_subject)
    is_valid = keyword_pattern.search(clean_subject) is not None
    
    if id_match and name_match and is_valid:
        return name_match.group(1), id_match.group(1)
    return None, None

def parse_subject(msg: Message) -> str:
    """优化解码逻辑"""
    subject_raw = msg.get("Subject", "")
    decoded_parts = []
    for part, charset in decode_header(subject_raw):
        try:
            if isinstance(part, bytes):
                decoded_parts.append(part.decode(charset or 'utf-8', errors='replace'))
            else:
                decoded_parts.append(str(part))
        except:
            decoded_parts.append(str(part))
    return "".join(decoded_parts)

async def main():
    email_processor = EmailProcessor()
    # 缓存加载学生名单
    new_hongji = read_student_list(str(NEW_HONGJI_FILE))
    last_year = read_student_list(str(LAST_YEAR_FILE))
    admitted, rejected, candidates = [], [], []

    try:
        async with AsyncSecureIMAPClient() as client:
            # 异步获取邮件
            email_generator = client.fetch_emails()
            # 收集所有附件路径用于批量解析
            docx_files = []
            email_data = []
            
            async for uid, msg in email_generator:
                recv_date = parsedate_to_datetime(msg.get("Date"))
                subject = parse_subject(msg)
                name, student_id = parse_subject_pattern(subject)

                if not student_id:
                    rejected.append({"学号": "未知", "姓名": "未知", "原主题": subject, "原因": "格式错误"})
                    continue

                # 保存附件（批量处理）
                attachments = email_processor.save_attachments(msg, student_id, name)
                docx_files.extend([str(f) for f in attachments if f.suffix.lower() == ".docx"])
                email_data.append((uid, student_id, name, recv_date, attachments))

            # 批量解析DOCX（多线程）
            docx_results = parse_docx_batch(docx_files)

            # 处理邮件数据
            for _, student_id, name, recv_date, attachments in email_data:
                # 去年已录取判定
                if student_id in last_year:
                    rejected.append({"学号": student_id, "姓名": name, "原因": "去年已录取"})
                    continue

                # 新鸿基直接录取
                if student_id in new_hongji:
                    admitted.append({"学号": student_id, "姓名": name, "备注": "新鸿基直接录取"})
                    continue

                # 校验DOCX附件
                docx_file = next((str(f) for f in attachments if f.suffix.lower() == ".docx"), None)
                if not docx_file:
                    rejected.append({"学号": student_id, "姓名": name, "原因": "缺少DOCX附件"})
                    continue

                # 使用批量解析结果
                info = docx_results.get(docx_file, {"is_supported": False, "reason_length": 0})
                if not info["is_supported"]:
                    rejected.append({"学号": student_id, "姓名": name, "原因": "非资助对象"})
                elif info["reason_length"] < 95:
                    rejected.append({"学号": student_id, "姓名": name, "原因": f"理由不足({info['reason_length']}字)"})
                else:
                    candidates.append((student_id, name, recv_date))

        # 处理名额逻辑
        remaining = 25 - len(admitted)
        candidates.sort(key=lambda x: x[2])
        for sid, n, _ in candidates[:remaining]:
            admitted.append({"学号": sid, "姓名": n, "备注": "普通录取"})
        for sid, n, _ in candidates[remaining:]:
            rejected.append({"学号": sid, "姓名": n, "原因": "名额已满"})

        save_results(admitted, rejected)
    except Exception as e:
        logger.error(f"运行失败: {e}")
        raise

if __name__ == "__main__":
    # 适配Windows异步环境
    if os.name == 'nt':
        asyncio.set_event_loop_policy(asyncio.WindowsSelectorEventLoopPolicy())
    asyncio.run(main())
