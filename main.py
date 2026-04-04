import logging
import os
from email_client import EmailClient
from docx_parser import DocxParser
import pandas as pd

logging.basicConfig(level=logging.INFO, format="%(message)s")

def main():
    logging.info("✅ 开始筛选")

    def get_ids(path):
        try:
            df = pd.read_excel(path, dtype=str)
            return set(df.iloc[:, 0].dropna().str.strip())
        except:
            return set()

    xhj_ids = get_ids("新鸿基名单.xlsx")
    black_ids = get_ids("黑名单.xlsx")
    last_ids = get_ids("去年名单.xlsx")

    client = EmailClient()
    mails = client.fetch_mails()
    logging.info(f"📩 共收取邮件：{len(mails)}封")

    accept_list = []
    reject_list = []

    for mail in mails:
        sid = mail.get("student_id", "")
        name = mail.get("name", "")
        receive_time = mail.get("receive_time", "")
        attach_path = mail.get("attachment_path", "")

        grade = "未知"
        apply_class = "未知"
        is_subsidy = False
        reason_count = 0
        reject_reason = ""
        parse_success = False

        if attach_path:
            try:
                parser = DocxParser(attach_path)
                grade = parser.get_grade()
                apply_class = parser.get_apply_class()
                is_subsidy = parser.is_subsidy()
                reason_count = parser.get_reason_count()
                parse_success = True
            except:
                reject_reason = "解析失败"

        # 筛选规则
        if sid in black_ids:
            reject_reason = "黑名单"
        elif sid in last_ids:
            reject_reason = "本年已参加"
        elif not attach_path:
            reject_reason = "无附件"
        elif not parse_success:
            reject_reason = "解析失败"
        else:
            if sid in xhj_ids:
                pass
            else:
                if not is_subsidy:
                    reject_reason = "非资助对象"
                elif reason_count < 100:
                    reject_reason = f"字数不足({reason_count}/100)"

        # 基础信息
        row = [sid, name, grade, reason_count, "是" if is_subsidy else "否", apply_class]

        if reject_reason:
            reject_list.append([*row, reject_reason, receive_time])
        else:
            accept_list.append([*row, receive_time])

    # 按时间排序
    accept_list.sort(key=lambda x: x[-1])
    reject_list.sort(key=lambda x: x[-1])

    # 去掉时间字段
    accept_final = [r[:-1] for r in accept_list]
    reject_final = [r[:-1] for r in reject_list]

    # 列表头
    accept_cols = ["学号", "姓名", "年级", "申请理由字数", "是否资助", "报名班级"]
    reject_cols = ["学号", "姓名", "年级", "申请理由字数", "是否资助", "报名班级", "拒绝原因"]

    # 导出总表
    pd.DataFrame(accept_final, columns=accept_cols).to_excel("录取名单.xlsx", index=False)
    pd.DataFrame(reject_final, columns=reject_cols).to_excel("拒绝名单.xlsx", index=False)

    # ==============================
    # 按班级分类导出（录取 + 拒绝）
    # ==============================
    if accept_final:
        accept_df = pd.DataFrame(accept_final, columns=accept_cols)
        for cls_name, group in accept_df.groupby("报名班级"):
            group.to_excel(f"录取_{cls_name}.xlsx", index=False)

    if reject_final:
        reject_df = pd.DataFrame(reject_final, columns=reject_cols)
        for cls_name, group in reject_df.groupby("报名班级"):
            group.to_excel(f"拒绝_{cls_name}.xlsx", index=False)

    logging.info(f"🎯 录取：{len(accept_final)} 人")
    logging.info(f"❌ 拒绝：{len(reject_final)} 人")

if __name__ == "__main__":
    main()
