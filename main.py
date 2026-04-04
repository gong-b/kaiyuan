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
            return set(df.iloc[:,0].dropna().str.strip())
        except:
            return set()

    xhj_ids = get_ids("新鸿基名单.xlsx")
    black_ids = get_ids("黑名单.xlsx")
    last_ids = get_ids("去年名单.xlsx")

    client = EmailClient()
    mails = client.fetch_mails()
    logging.info(f"📩 共收取邮件：{len(mails)}封")

    accept = []
    reject = []

    for mail in mails:
        sid = mail.get("student_id", "")
        name = mail.get("name", "")
        rcv_time = mail.get("receive_time", "")
        path = mail.get("attachment_path", "")

        grade = "未知"
        cls = "未知班级"
        sub = False
        cnt = 0
        reason = ""
        ok = False

        if path:
            try:
                p = DocxParser(path)
                grade = p.get_grade()
                cls = p.get_apply_class()
                sub = p.get_subsidy()
                cnt = p.get_reason_count()
                ok = True
            except:
                reason = "解析失败"

        # 筛选规则
        if sid in black_ids:
            reason = "黑名单"
        elif sid in last_ids:
            reason = "本年已参加"
        elif sid in xhj_ids:
            reason = ""
        elif not path:
            reason = "无附件"
        elif not ok:
            reason = "解析失败"
        else:
            if not sub:
                reason = "非资助对象"
            elif cnt < 100:
                reason = f"字数不足({cnt}/100)"

        row = [sid, name, grade, cnt, "是" if sub else "否", cls]

        if reason:
            reject.append([*row, reason, rcv_time])
        else:
            accept.append([*row, rcv_time])

    # 按时间排序
    accept.sort(key=lambda x: x[-1])
    reject.sort(key=lambda x: x[-1])

    acc = [r[:-1] for r in accept]
    rej = [r[:-1] for r in reject]

    cols_acc = ["学号","姓名","年级","申请理由字数","是否资助","报名班级"]
    cols_rej = ["学号","姓名","年级","申请理由字数","是否资助","报名班级","拒绝原因"]

    # 导出总表
    pd.DataFrame(acc, columns=cols_acc).to_excel("录取名单.xlsx", index=False)
    pd.DataFrame(rej, columns=cols_rej).to_excel("拒绝名单.xlsx", index=False)

    # ==========================================
    # 🔥 关键：按班级 100% 分开导出
    # ==========================================
    if acc:
        df_acc = pd.DataFrame(acc, columns=cols_acc)
        for c, g in df_acc.groupby("报名班级"):
            g.to_excel(f"录取_{c}.xlsx", index=False)

    if rej:
        df_rej = pd.DataFrame(rej, columns=cols_rej)
        for c, g in df_rej.groupby("报名班级"):
            g.to_excel(f"拒绝_{c}.xlsx", index=False)

    logging.info(f"🎯 录取：{len(acc)} 人")
    logging.info(f"❌ 拒绝：{len(rej)} 人")

if __name__ == "__main__":
    main()
