import logging
import re
import os
from datetime import datetime
from email.header import decode_header
from email.message import Message
from email.utils import parsedate_to_datetime
from config import *
from email_client import AsyncSecureIMAPClient  # 替换为异步客户端
from email_processor import EmailProcessor
from docx_parser import parse_docx
from excel_handler import read_student_list, save_results
import asyncio
from multiprocessing import Pool, cpu_count
from concurrent.futures import ProcessPoolExecutor

logging.basicConfig(level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s")
logger = logging.getLogger(__name__)

# 原有parse_subject_pattern、parse_subject函数保留

async def process_single_email(msg, email_processor, new_hongji, last_year, rejected, admitted, candidates):
    """单邮件异步处理逻辑"""
    recv_date = parsedate_to_datetime(msg.get("Date"))
    subject = parse_subject(msg)
    name, student_id = parse_subject_pattern(subject)

    if not student_id:
        rejected.append({"学号": "未知", "姓名": "未知", "原主题": subject, "原因": "格式错误"})
        return

    # 异步保存附件（优化IO）
    attachments = await asyncio.to_thread(email_processor.save_attachments, msg, student_id, name)
    
    if student_id in last_year:
        rejected.append({"学号": student_id, "姓名": name, "原因": "去年已录取"})
        return

    if student_id in new_hongji:
        admitted.append({"学号": student_id, "姓名": name, "备注": "新鸿基直接录取"})
        return

    docx_files = [a for a in attachments if a.suffix.lower() == ".docx"]
    if not docx_files:
        rejected.append({"学号": student_id, "姓名": name, "原因": "缺少DOCX附件"})
        return

    # 提交到进程池解析DOCX（CPU密集型任务）
    return (student_id, name, recv_date, str(docx_files[0]))

def parse_docx_mp(args):
    """多进程封装DOCX解析"""
    student_id, name, filepath = args
    try:
        info = parse_docx(filepath)
        if not info["is_supported"]:
            return ("reject", student_id, name, "非资助对象")
        elif info["reason_length"] < 95:
            return ("reject", student_id, name, f"理由不足({info['reason_length']}字)")
        else:
            return ("candidate", student_id, name)
    except:
        return ("reject", student_id, name, "文档解析失败")

async def main():
    email_processor = EmailProcessor()
    new_hongji = read_student_list(str(NEW_HONGJI_FILE))
    last_year = read_student_list(str(LAST_YEAR_FILE))
    admitted, rejected, candidates = [], [], []
    docx_tasks = []  # 待解析的DOCX任务

    try:
        async with AsyncSecureIMAPClient() as client:
            # 异步迭代邮件
            async for uid, msg in client.fetch_emails_batch(batch_size=100):
                res = await process_single_email(msg, email_processor, new_hongji, last_year, rejected, admitted, candidates)
                if res:  # 有DOCX需要解析
                    docx_tasks.append(res)

        # 多进程解析DOCX（CPU密集型）
        if docx_tasks:
            logger.info(f"开始多进程解析{len(docx_tasks)}个DOCX文件")
            # 提取参数：(student_id, name, filepath)
            mp_args = [(t[0], t[1], t[3]) for t in docx_tasks]
            with ProcessPoolExecutor(max_workers=cpu_count()) as executor:
                results = executor.map(parse_docx_mp, mp_args)
            
            # 整理解析结果
            recv_dates = {f"{t[0]}_{t[1]}": t[2] for t in docx_tasks}
            for res in results:
                if res[0] == "reject":
                    rejected.append({"学号": res[1], "姓名": res[2], "原因": res[3]})
                else:
                    candidates.append((res[1], res[2], recv_dates[f"{res[1]}_{res[2]}"]))

        # 名额逻辑（原有保留）
        remaining = 25 - len(admitted)
        candidates.sort(key=lambda x: x[2])
        for sid, n, _ in candidates[:remaining]: 
            admitted.append({"学号": sid, "姓名": n, "备注": "普通录取"})
        for sid, n, _ in candidates[remaining:]: 
            rejected.append({"学号": sid, "姓名": n, "原因": "名额已满"})

        save_results(admitted, rejected)
        logger.info("处理完成！")
    except Exception as e: 
        logger.error(f"运行失败: {e}")

if __name__ == "__main__":
    asyncio.run(main())
