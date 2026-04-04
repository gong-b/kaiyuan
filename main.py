import os
import shutil
import zipfile
import pandas as pd
from email.utils import parsedate_to_datetime
from email_client import EmailClient
from docx_parser import DocxParser

# 关闭所有警告日志
import logging
logging.basicConfig(level=logging.ERROR)
import warnings
warnings.filterwarnings("ignore")

def main():
    # 清理附件目录
    if os.path.exists("attachments"):
        shutil.rmtree("attachments")
    os.makedirs("attachments", exist_ok=True)

    # 读取名单
    def get_ids(path):
        try:
            df = pd.read_excel(path, dtype=str)
            return set(df.iloc[:, 0].dropna().astype(str).str.strip())
        except:
            return set()

    xhj_ids = get_ids("新鸿基名单.xlsx")
    black_ids = get_ids("黑名单.xlsx")
    last_ids = get_ids("副本去年报名名单.xlsx")

    client = EmailClient()
    mails = client.fetch_mails()

    accept_list = []
    reject_list = []
    processed = set()
    all_files = []

    for mail in mails:
        sid = mail.get("student_id", "")
        name = mail.get("name", "")
        rcv_time = mail.get("receive_time", "")
        attach_path = mail.get("attachment_path", "")

        grade = "未知"
        cls = "未知班级"
        sub = False
        cnt = 0
        reason = ""
        ok = False
        dt = None

        try:
            dt = parsedate_to_datetime(rcv_time)
        except:
            dt = None

        if attach_path:
            try:
                p = DocxParser(attach_path)
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
        elif not attach_path:
            reason = "无附件"
        elif not ok:
            reason = "解析失败"
        else:
            if not sub:
                reason = "非资助对象"
            elif cnt < 100:
                reason = f"字数不足({cnt}/100)"

        # 一人多报只录最早
        if not reason:
            if sid in processed:
                reason = "重复报名，仅录取最先报名班级"
            else:
                processed.add(sid)

        # 复制附件
        if attach_path and os.path.exists(attach_path):
            try:
                ext = os.path.splitext(attach_path)[1]
                new_name = f"{sid}_{name}{ext}"
                target = os.path.join("attachments", new_name)
                shutil.copy(attach_path, target)
                all_files.append(target)
            except:
                pass

        row = [sid, name, grade, cnt, "是" if sub else "否", cls, dt]
        if reason:
            reject_list.append([*row, reason])
        else:
            accept_list.append(row)

    # 排序：先班级，再时间正序
    accept_list.sort(key=lambda x: (x[5], x[6] if x[6] else ""))
    reject_list.sort(key=lambda x: (x[5], x[6] if x[6] else ""))

    acc = [[x[0],x[1],x[2],x[3],x[4],x[5]] for x in accept_list]
    rej = [[x[0],x[1],x[2],x[3],x[4],x[5],x[7]] for x in reject_list]

    cols_acc = ["学号","姓名","年级","申请理由字数","是否资助","报名班级"]
    cols_rej = ["学号","姓名","年级","申请理由字数","是否资助","报名班级","拒绝原因"]

    df_acc = pd.DataFrame(acc, columns=cols_acc)
    df_rej = pd.DataFrame(rej, columns=cols_rej)

    # 导出Excel
    df_acc.to_excel("录取名单.xlsx", index=False)
    df_rej.to_excel("拒绝名单.xlsx", index=False)

    for c, g in df_acc.groupby("报名班级"):
        g.to_excel(f"录取_{c}.xlsx", index=False)
    for c, g in df_rej.groupby("报名班级"):
        g.to_excel(f"拒绝_{c}.xlsx", index=False)

    # 打包附件
    zip_path = "所有报名附件.zip"
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
        seen = set()
        for f in all_files:
            if os.path.exists(f) and f not in seen:
                zf.write(f, os.path.basename(f))
                seen.add(f)

    print(f"✅ 筛选完成")
    print(f"🎯 录取：{len(acc)} 人")
    print(f"❌ 拒绝：{len(rej)} 人")

if __name__ == "__main__":
    main()
