import logging
import os
from email_client import EmailClient
from docx_parser import DocxParser
import pandas as pd

logging.basicConfig(level=logging.INFO, format="%(message)s")

def main():
    logging.info("✅ 开始筛选")

    # 读取名单
    def get_ids(path):
        try:
            return set(pd.read_excel(path, dtype=str).iloc[:,0].dropna().str.strip())
        except:
            return set()

    xhj_ids = get_ids("新鸿基名单.xlsx")
    black_ids = get_ids("黑名单.xlsx")
    last_ids = get_ids("去年名单.xlsx")

    client = EmailClient()
    mails = client.fetch_mails()
    logging.info(f"📩 共收取邮件：{len(mails)}封")

    final_accept = []
    final_reject = []

    for mail in mails:
        sid = mail.get("student_id", "")
        name = mail.get("name", "")
        receive_time = mail.get("receive_time", "")
        path = mail.get("attachment_path", "")

        grade = sid[:4] if len(sid) >=4 else ""
        status = "报名成功"
        reason = ""

        # 拦截规则
        if sid in black_ids:
            reason = "黑名单"
        elif sid in last_ids:
            reason = "本年已参加"
        elif not path:
            reason = "无附件"
        else:
            try:
                parser = DocxParser(path)
                is_sub = parser.get_subsidy()
                cnt = parser.get_reason_count()
                cls = parser.get_class()

                if sid in xhj_ids:
                    status = "报名成功"
                else:
                    if not is_sub:
                        reason = "非资助对象"
                    elif cnt < 100:
                        reason = f"字数不足({cnt}/100)"
                    else:
                        status = "报名成功"
            except:
                reason = "文件解析失败"

        # 最终信息
        is_sub = False
        cnt = 0
        cls = ""
        try:
            parser = DocxParser(path)
            is_sub = parser.get_subsidy()
            cnt = parser.get_reason_count()
            cls = parser.get_class() or "未填写"
        except:
            pass

        row = [
            sid, name, grade,
            cnt, "是" if is_sub else "否",
            cls, receive_time,
            status if not reason else "报名失败",
            reason
        ]

        if reason:
            final_reject.append(row)
        else:
            final_accept.append(row)

    # 按邮件时间排序
    final_accept.sort(key=lambda x: x[6])
    final_reject.sort(key=lambda x: x[6])

    # 导出完整表格
    cols = ["学号","姓名","年级","申请理由字数","是否资助","报名班级","报名时间","状态","拒绝原因"]
    pd.DataFrame(final_accept, columns=cols).to_excel("录取名单.xlsx", index=False)
    pd.DataFrame(final_reject, columns=cols).to_excel("拒绝名单.xlsx", index=False)

    # 分班导出
    accept_df = pd.DataFrame(final_accept, columns=cols)
    for cls_name, group in accept_df.groupby("报名班级"):
        group.to_excel(f"录取_{cls_name}.xlsx", index=False)

    logging.info(f"🎯 录取：{len(final_accept)}人")
    logging.info(f"❌ 拒绝：{len(final_reject)}人")

if __name__ == "__main__":
    main()
