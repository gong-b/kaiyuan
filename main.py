import logging
import os
from email_client import EmailClient
from docx_parser import DocxParser
import pandas as pd

logging.basicConfig(level=logging.INFO, format="%(message)s")

def main():
    logging.info("✅ 开始筛选")

    # 读取名单工具
    def get_ids(path):
        try:
            df = pd.read_excel(path, dtype=str)
            return set(df.iloc[:, 0].dropna().str.strip())
        except:
            return set()

    xhj_ids   = get_ids("新鸿基名单.xlsx")
    black_ids = get_ids("黑名单.xlsx")
    last_ids  = get_ids("去年名单.xlsx")

    # 获取邮件
    client = EmailClient()
    mails = client.fetch_mails()
    logging.info(f"📩 共收取邮件：{len(mails)}封")

    accept_list = []  # 录取
    reject_list = []  # 拒绝

    for mail in mails:
        # 邮件基础信息
        sid          = mail.get("student_id", "")
        name         = mail.get("name", "")
        receive_time = mail.get("receive_time", "")
        attach_path  = mail.get("attachment_path", "")

        # 默认值
        grade         = "无"
        apply_class   = "无"
        is_subsidy    = False
        reason_len    = 0
        reject_reason = ""
        parse_ok      = False

        # 先解析附件
        if attach_path:
            try:
                parser = DocxParser(attach_path)
                grade       = parser.get_grade()
                apply_class = parser.get_apply_class()
                is_subsidy  = parser.get_subsidy()
                reason_len  = parser.get_reason_count()
                parse_ok    = True
            except:
                reject_reason = "文件解析失败"

        # ======================
        # 审核规则（严格按你要求）
        # ======================
        if sid in black_ids:
            reject_reason = "黑名单"
        elif sid in last_ids:
            reject_reason = "本年已参加"
        elif not attach_path:
            reject_reason = "无Word附件"
        elif not parse_ok:
            reject_reason = "文件解析失败"
        else:
            if sid in xhj_ids:
                # 新鸿基直接录取
                pass
            else:
                # 非新鸿基：必须资助 + 字数≥100
                if not is_subsidy:
                    reject_reason = "非资助对象"
                elif reason_len < 100:
                    reject_reason = f"理由字数不足({reason_len}/100)"

        # ======================
        # 组装字段
        # ======================
        base_row = [
            sid,
            name,
            grade,
            reason_len,
            "是" if is_subsidy else "否",
            apply_class
        ]

        if reject_reason:
            # 拒绝名单（带原因）
            reject_list.append([*base_row, reject_reason, receive_time])
        else:
            # 录取名单（不带拒绝原因、不带时间）
            accept_list.append([*base_row, receive_time])

    # ======================
    # 按【报名时间先后】排序
    # ======================
    accept_list.sort(key=lambda x: x[-1])  # 按时间排序
    reject_list.sort(key=lambda x: x[-1])

    # 去掉排序用的时间字段（不导出）
    accept_export = [row[:-1] for row in accept_list]
    reject_export = [row[:-1] for row in reject_list]

    # ======================
    # 列表头
    # ======================
    accept_cols = ["学号", "姓名", "年级", "申请理由字数", "是否资助", "报名班级"]
    reject_cols = ["学号", "姓名", "年级", "申请理由字数", "是否资助", "报名班级", "拒绝原因"]

    # 导出总表
    pd.DataFrame(accept_export, columns=accept_cols).to_excel("录取名单.xlsx", index=False)
    pd.DataFrame(reject_export, columns=reject_cols).to_excel("拒绝名单.xlsx", index=False)

    # ======================
    # 按班级分班导出
    # ======================
    if accept_export:
        accept_df = pd.DataFrame(accept_export, columns=accept_cols)
        for class_name, group in accept_df.groupby("报名班级"):
            group.to_excel(f"录取_{class_name}.xlsx", index=False)

    logging.info(f"🎯 录取：{len(accept_export)} 人")
    logging.info(f"❌ 拒绝：{len(reject_export)} 人")

if __name__ == "__main__":
    main()
