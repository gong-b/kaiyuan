import logging
import os
from email_client import EmailClient
from excel_handler import ExcelHandler
from docx_parser import DocxParser
import config

logging.basicConfig(level=logging.INFO, format="%(message)s")

def main():
    logging.info("✅ 开始筛选")

    excel = ExcelHandler()
    xhj_ids = excel.read_student_ids("新鸿基名单.xlsx")
    black_ids = excel.read_student_ids("黑名单.xlsx")
    last_ids = excel.read_student_ids("去年名单.xlsx")

    client = EmailClient()
    mails = client.fetch_mails()
    logging.info(f"📩 共扫描邮件：{len(mails)}封")

    accept = []
    reject = []

    for mail in mails:
        sid = mail.get("student_id", "")
        name = mail.get("name", "")
        if not sid:
            reject.append(["未知", "未知", "无学号"])
            continue

        if sid in black_ids:
            reject.append([sid, name, "黑名单"])
            continue
        if sid in last_ids:
            reject.append([sid, name, "本年已参加"])
            continue

        if sid in xhj_ids:
            accept.append([sid, name, "新鸿基(直录)"])
            continue

        doc_path = mail.get("attachment_path")
        if not doc_path:
            reject.append([sid, name, "无Word附件"])
            continue

        try:
            parser = DocxParser(doc_path)
            is_sub = parser.is_subsidy()
            cnt = parser.get_reason_length()

            logging.info(f"✅ {sid} {name} | 资助:{is_sub} | 字数:{cnt}")

            if not is_sub:
                reject.append([sid, name, "非资助对象"])
                continue

            if cnt >= 100:
                accept.append([sid, name, f"资助通过({cnt}字)"])
            else:
                reject.append([sid, name, f"字数不足({cnt}/100)"])

        except Exception as e:
            reject.append([sid, name, "文件异常"])
            continue

    accept = accept[:config.MAX_ACCEPT]
    excel.write_accept(accept)
    excel.write_reject(reject)

    logging.info(f"\n🎯 最终录取：{len(accept)}人")

if __name__ == "__main__":
    main()
