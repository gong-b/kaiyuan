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
        except Exception as e:
            print(f"⚠️ 读取名单 {path} 失败: {e}")
            return set()

    xhj_ids = get_ids("新鸿基名单.xlsx")
    black_ids = get_ids("黑名单.xlsx")
    last_ids = get_ids("去年名单.xlsx")

    client = EmailClient()
    mails = client.fetch_mails()
    logging.info(f"📩 共收取邮件：{len(mails)}封")

    accept_list = []
    reject_list = []
    processed_students = set()  # 记录已经录取的学生学号

    for mail in mails:
        sid = mail.get("student_id", "")
        name = mail.get("name", "")
        receive_time = mail.get("receive_time", "")
        attach_path = mail.get("attachment_path", "")

        grade = "未知"
        apply_class = "未知班级"
        is_subsidy = False
        reason_count = 0
        reject_reason = ""
        parse_success = False

        if attach_path:
            try:
                parser = DocxParser(attach_path)
                grade = parser.get_grade()
                apply_class = parser.get_apply_class()
                is_subsidy = parser.get_subsidy()
                reason_count = parser.get_reason_count()
                parse_success = True
            except Exception as e:
                reject_reason = "文件解析失败"
                print(f"⚠️ {sid} {name} 解析失败: {e}")

        # 审核规则
        if sid in black_ids:
            reject_reason = "黑名单"
        elif sid in last_ids:
            reject_reason = "本年已参加"
        elif sid in xhj_ids:
            reject_reason = ""
        elif not attach_path:
            reject_reason = "无Word附件"
        elif not parse_success:
            reject_reason = "文件解析失败"
        else:
            if not is_subsidy:
                reject_reason = "非资助对象"
            elif reason_count < 100:
                reject_reason = f"理由字数不足({reason_count}/100)"
            else:
                reject_reason = ""

        # ======================
        # 🔥 核心：一人多报处理 - 只录取最先报名的
        # ======================
        if not reject_reason:
            if sid in processed_students:
                reject_reason = "重复报名，仅保留最先报名的班级"
            else:
                processed_students.add(sid)

        # 组装数据
        base_row = [sid, name, grade, reason_count, "是" if is_subsidy else "否", apply_class]

        if reject_reason:
            reject_list.append([*base_row, reject_reason, receive_time])
        else:
            accept_list.append([*base_row, receive_time])

    # ======================
    # 🔥 核心：日期正序排序（最早报名 → 排在最前）
    # ======================
    accept_list.sort(key=lambda x: x[-1])  # 时间正序
    reject_list.sort(key=lambda x: x[-1])

    # 去掉时间字段
    accept_final = [row[:-1] for row in accept_list]
    reject_final = [row[:-1] for row in reject_list]

    # 表头
    accept_cols = ["学号", "姓名", "年级", "申请理由字数", "是否资助", "报名班级"]
    reject_cols = ["学号", "姓名", "年级", "申请理由字数", "是否资助", "报名班级", "拒绝原因"]

    # 导出总表
    pd.DataFrame(accept_final, columns=accept_cols).to_excel("录取名单.xlsx", index=False)
    pd.DataFrame(reject_final, columns=reject_cols).to_excel("拒绝名单.xlsx", index=False)

    # 按班级分类导出（同一个班级在一起）
    if accept_final:
        accept_df = pd.DataFrame(accept_final, columns=accept_cols)
        for cls_name, group in accept_df.groupby("报名班级"):
            group.to_excel(f"录取_{cls_name}.xlsx", index=False)

    if reject_final:
        reject_df = pd.DataFrame(reject_final, columns=reject_cols)
        for cls_name, group in reject_df.groupby("报名班级"):
            group.to_excel(f"拒绝_{cls_name}.xlsx", index=False)

    logging.info(f"🎯 最终录取：{len(accept_final)} 人")
    logging.info(f"❌ 最终拒绝：{len(reject_final)} 人")

if __name__ == "__main__":
    main()
