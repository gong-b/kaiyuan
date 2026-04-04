import logging
from email_client import EmailClient
from excel_handler import ExcelHandler
from docx_parser import DocxParser
from email_processor import EmailProcessor
import config

logging.basicConfig(level=logging.INFO, format="%(message)s")

def main():
    logging.info("✅ 开始筛选")

    # 读取上传的3个名单
    excel = ExcelHandler()
    xhj_ids = excel.read_student_ids("新鸿基名单.xlsx")
    black_ids = excel.read_student_ids("黑名单.xlsx")
    last_ids = excel.read_student_ids("去年名单.xlsx")

    # 收邮件
    client = EmailClient(config.EMAIL_HOST, config.EMAIL_PORT, config.EMAIL_USER, config.EMAIL_PASS)
    mails = client.fetch_mails()
    logging.info(f"📩 共收取邮件：{len(mails)}封")

    accept = []
    reject = []

    for mail in mails:
        sid = mail.get("student_id")
        name = mail.get("name")
        if not sid:
            continue

        # 审核规则
        if sid in black_ids:
            reject.append([sid, name, "黑名单"])
            continue
        if sid in last_ids:
            reject.append([sid, name, "去年已参加"])
            continue
        if sid not in xhj_ids:
            reject.append([sid, name, "非新鸿基学生"])
            continue

        # 检查附件
        doc_path = mail.get("attachment_path")
        if not doc_path:
            reject.append([sid, name, "未上传申请表"])
            continue

        # 解析Word
        parser = DocxParser(doc_path)
        is_subsidy = parser.is_subsidy()
        word_count = parser.count_reason()
        logging.info(f"{sid} 理由字数：{word_count}")

        if not is_subsidy:
            reject.append([sid, name, "非资助对象"])
            continue
        if word_count < config.REASON_MIN_WORDS:
            reject.append([sid, name, f"理由字数不足{config.REASON_MIN_WORDS}"])
            continue

        accept.append([sid, name, "审核通过"])

    # 最多录取25人
    accept = accept[:config.MAX_ACCEPT]
    
    # 输出Excel
    excel.write_accept(accept)
    excel.write_reject(reject)

    logging.info(f"\n🎯 最终录取：{len(accept)}人")
    logging.info("✅ 录取名单.xlsx 已生成")
    logging.info("✅ 拒绝名单.xlsx 已生成")

if __name__ == "__main__":
    main()
