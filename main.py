import logging
import os
from email_client import EmailClient
from excel_handler import ExcelHandler
from docx_parser import DocxParser

logging.basicConfig(level=logging.INFO, format="%(message)s")

def main():
    logging.info("✅ 开始筛选")
    excel = ExcelHandler()
    xhj = excel.read_student_ids("新鸿基名单.xlsx")
    black = excel.read_student_ids("黑名单.xlsx")
    last = excel.read_student_ids("去年名单.xlsx")

    client = EmailClient()
    mails = client.fetch_mails()
    logging.info(f"📩 共收取邮件：{len(mails)}封")

    accept = []
    reject = []

    for mail in mails:
        sid = mail.get("student_id", "")
        name = mail.get("name", "")
        if not sid:
            continue

        if sid in black:
            reject.append([sid, name, "黑名单"])
            continue
        if sid in last:
            reject.append([sid, name, "已参加"])
            continue
        if sid in xhj:
            accept.append([sid, name, "新鸿基直录"])
            continue

        path = mail.get("attachment_path")
        if not path:
            reject.append([sid, name, "无附件"])
            continue

        try:
            p = DocxParser(path)
            sub = p.is_subsidy()
            cnt = p.get_reason_length()

            if not sub:
                reject.append([sid, name, "非资助对象"])
            elif cnt >= 100:
                accept.append([sid, name, f"通过({cnt}字)"])
            else:
                reject.append([sid, name, f"字数不足({cnt}/100)"])
        except:
            reject.append([sid, name, "文件错误"])

    accept = accept[:25]
    excel.write_accept(accept)
    excel.write_reject(reject)
    logging.info(f"🎯 录取：{len(accept)}人")

if __name__ == "__main__":
    main()
