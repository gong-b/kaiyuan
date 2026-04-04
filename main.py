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

        # 组装数据
        base_row = [sid, name, grade, reason_count, "是" if is_subsidy else "否", apply_class, receive_time]

        if reject_reason:
            reject_list.append([*base_row, reject_reason])
        else:
            accept_list.append(base_row)

    # ==============================================
    # 🔥 关键：先按班级排序 → 再按时间排序
    # ==============================================
    accept_list.sort(key=lambda x: (x[5], x[6]))  # 班级、时间
    reject_list.sort(key=lambda x: (x[5], x[6]))

    # 去掉时间字段
    accept_final = [[x[0],x[1],x[2],x[3],x[4],x[5]] for x in accept_list]
    reject_final = [[x[0],x[1],x[2],x[3],x[4],x[5],x[7]] for x in reject_list]

    # 表头
    accept_cols = ["学号", "姓名", "年级", "申请理由字数", "是否资助", "报名班级"]
    reject_cols = ["学号", "姓名", "年级", "申请理由字数", "是否资助", "报名班级", "拒绝原因"]

    # ==============================
    # 导出一个Excel，但班级已归类在一起
    # ==============================
    pd.DataFrame(accept_final, columns=accept_cols).to_excel("录取名单.xlsx", index=False)
    pd.DataFrame(reject_final, columns=reject_cols).to_excel("拒绝名单.xlsx", index=False)

    logging.info(f"🎯 最终录取：{len(accept_final)} 人")
    logging.info(f"❌ 最终拒绝：{len(reject_final)} 人")

if __name__ == "__main__":
    main()
