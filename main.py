import logging
import os
from email_client import EmailClient
from excel_handler import ExcelHandler
from docx_parser import DocxParser
import config

logging.basicConfig(level=logging.INFO, format="%(message)s")

def main():
    logging.info("✅ 开始筛选")

    # 读取名单
    excel = ExcelHandler()
    xhj_ids = excel.read_student_ids("新鸿基名单.xlsx")
    last_ids = excel.read_student_ids("去年名单.xlsx")

    # 收邮件
    client = EmailClient()
    mails = client.fetch_mails()
    logging.info(f"📩 共收取邮件：{len(mails)}封")

    accept = []
    reject = []

    for mail in mails:
        sid = mail.get("student_id", "")
        name = mail.get("name", "")
        if not sid:
            reject.append(["未知", "未知", "主题格式错误"])
            continue

        # 规则
        if sid in last_ids:
            reject.append([sid, name, "去年已参加"])
            continue

        # 新鸿基直接进入候选
        if sid not in xhj_ids:
            reject.append([sid, name, "非新鸿基学生"])
            continue

        # 附件检查
        doc_path = mail.get("attachment_path")
        if not doc_path:
            reject.append([sid, name, "未上传Word附件"])
            continue

        # 解析申请表
        parser = DocxParser(doc_path)
        is_supported = parser.is_supported()
        reason_len = parser.get_reason_length()

        if not is_supported:
            reject.append([sid, name, "非资助对象"])
            continue
        if reason_len < 95:
            reject.append([sid, name, f"申请理由字数不足({reason_len})"])
            continue

        accept.append([sid, name, "通过审核"])

    # 最多录取25人
    accept = accept[:25]
    excel.write_accept(accept)
    excel.write_reject(reject)

    logging.info(f"\n🎯 录取：{len(accept)}人")

if __name__ == "__main__":
    main()
