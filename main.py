import logging
import os
from email_client import EmailClient
from excel_handler import ExcelHandler
from docx_parser import DocxParser
import config

logging.basicConfig(level=logging.INFO, format="%(message)s")

def main():
    logging.info("✅ 开始筛选")

    # 1. 读取名单
    excel = ExcelHandler()
    xhj_ids = excel.read_student_ids("新鸿基名单.xlsx")
    black_ids = excel.read_student_ids("黑名单.xlsx")
    last_ids = excel.read_student_ids("去年名单.xlsx")

    # 2. 收邮件（使用日期范围筛选）
    client = EmailClient()
    mails = client.fetch_mails()
    logging.info(f"📩 共扫描邮件：{len(mails)}封")

    accept = []
    reject = []

    # 3. 遍历每封邮件进行审核
    for mail in mails:
        sid = mail.get("student_id", "")
        name = mail.get("name", "")
        if not sid:
            reject.append(["未知", "未知", "主题无学号"])
            continue

        # -------- 规则1：黑名单和去年已参加 直接拒绝 --------
        if sid in black_ids:
            reject.append([sid, name, "黑名单"])
            continue
        if sid in last_ids:
            reject.append([sid, name, "本年已参加"])
            continue

        # -------- 规则2：新鸿基学生 直接录取 --------
        if sid in xhj_ids:
            accept.append([sid, name, "新鸿基(直录)"])
            continue

        # -------- 规则3：非新鸿基学生 审核资助对象 + 字数 --------
        # 解析Word文档
        doc_path = mail.get("attachment_path")
        if not doc_path:
            reject.append([sid, name, "无Word附件"])
            continue

        try:
            parser = DocxParser(doc_path)
            is_subsidy = parser.is_subsidy()
            reason_len = parser.get_reason_length()

            if not is_subsidy:
                reject.append([sid, name, "非资助对象"])
                continue
            
            # 是资助对象，检查字数
            if reason_len >= 100:
                accept.append([sid, name, f"资助对象(字数{reason_len})"])
            else:
                reject.append([sid, name, f"理由不足({reason_len}/100)"])
                
        except Exception as e:
            reject.append([sid, name, f"解析失败"])
            continue

    # 4. 最多录取25人
    accept = accept[:config.MAX_ACCEPT]
    
    # 5. 输出结果
    excel.write_accept(accept)
    excel.write_reject(reject)

    logging.info(f"\n🎯 最终录取：{len(accept)}人")
    logging.info(f"❌ 拒绝人数：{len(reject)}人")

if __name__ == "__main__":
    main()
