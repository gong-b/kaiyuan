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
        except Exception as e:
            print(f"⚠️ 读取名单 {path} 失败: {e}")
            return set()

    xhj_ids = get_ids("新鸿基名单.xlsx")
    black_ids = get_ids("黑名单.xlsx")
    last_ids = get_ids("去年名单.xlsx")

    # 获取邮件
    client = EmailClient()
    mails = client.fetch_mails()
    logging.info(f"📩 共收取邮件：{len(mails)}封")

    accept_list = []
    reject_list = []

    for mail in mails:
        # 邮件基础信息
        sid = mail.get("student_id", "")
        name = mail.get("name", "")
        receive_time = mail.get("receive_time", "")
        attach_path = mail.get("attachment_path", "")

        # 默认值
        grade = "未知"
        apply_class = "未知班级"
        is_subsidy = False
        reason_count = 0
        reject_reason = ""
        parse_success = False

        # 解析附件（即使解析失败，也不影响新鸿基直录）
        if attach_path:
            try:
                parser = DocxParser(attach_path)
                grade = parser.get_grade()
                apply_class = parser.get_apply_class()
                is_subsidy = parser.is_subsidy()
                reason_count = parser.get_reason_count()
                parse_success = True
            except Exception as e:
                reject_reason = "文件解析失败"
                print(f"⚠️ {sid} {name} 解析失败: {e}")

        # ======================
        # 核心审核规则（修复全拒绝）
        # ======================
        # 1. 黑名单/已参加 直接拒绝（最高优先级）
        if sid in black_ids:
            reject_reason = "黑名单"
        elif sid in last_ids:
            reject_reason = "本年已参加"
        # 2. 新鸿基学生 直接录取（不受解析失败影响）
        elif sid in xhj_ids:
            reject_reason = ""  # 直接录取，清空拒绝原因
        # 3. 无附件 直接拒绝
        elif not attach_path:
            reject_reason = "无Word附件"
        # 4. 非新鸿基，且解析失败 拒绝
        elif not parse_success:
            reject_reason = "文件解析失败"
        # 5. 非新鸿基，正常审核
        else:
            if not is_subsidy:
                reject_reason = "非资助对象"
            elif reason_count < 100:
                reject_reason = f"理由字数不足({reason_count}/100)"
            else:
                reject_reason = ""  # 符合条件，录取

        # ======================
        # 组装数据
        # ======================
        base_row = [
            sid,
            name,
            grade,
            reason_count,
            "是" if is_subsidy else "否",
            apply_class
        ]

        if reject_reason:
            reject_list.append([*base_row, reject_reason, receive_time])
        else:
            accept_list.append([*base_row, receive_time])

    # ======================
    # 按报名时间排序
    # ======================
    accept_list.sort(key=lambda x: x[-1])
    reject_list.sort(key=lambda x: x[-1])

    # 去掉排序用的时间字段
    accept_final = [row[:-1] for row in accept_list]
    reject_final = [row[:-1] for row in reject_list]

    # ======================
    # 表头
    # ======================
    accept_cols = ["学号", "姓名", "年级", "申请理由字数", "是否资助", "报名班级"]
    reject_cols = ["学号", "姓名", "年级", "申请理由字数", "是否资助", "报名班级", "拒绝原因"]

    # 导出总表
    pd.DataFrame(accept_final, columns=accept_cols).to_excel("录取名单.xlsx", index=False)
    pd.DataFrame(reject_final, columns=reject_cols).to_excel("拒绝名单.xlsx", index=False)

    # ======================
    # 按班级分类导出（录取+拒绝）
    # ======================
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
