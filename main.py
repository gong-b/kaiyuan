import streamlit as st
import pandas as pd
import os
from email_client import EmailClient
from docx_parser import DocxParser
from email.utils import parsedate_to_datetime
from io import BytesIO

# ------------------- 页面标题 -------------------
st.title("🎓 书法班报名自动筛选系统")

# ------------------- 用户输入 -------------------
email = st.text_input("浙大邮箱")
pwd = st.text_input("客户端专用密码", type="password")
start_date = st.text_input("开始日期 (例：2025-10-02)")
end_date = st.text_input("结束日期 (例：2025-10-10)")

st.subheader("📂 上传名单")
xhj_file = st.file_uploader("新鸿基名单", type="xlsx")
black_file = st.file_uploader("黑名单", type="xlsx")
last_file = st.file_uploader("去年已参加名单", type="xlsx")

# ------------------- 开始筛选 -------------------
if st.button("✅ 开始筛选"):
    if not all([email, pwd, start_date, end_date, xhj_file, black_file, last_file]):
        st.warning("请填写完整信息！")
        st.stop()

    # 读取名单
    def get_ids(file):
        try:
            return set(pd.read_excel(BytesIO(file.getvalue()), dtype=str).iloc[:,0].dropna().str.strip())
        except:
            return set()

    xhj_ids = get_ids(xhj_file)
    black_ids = get_ids(black_file)
    last_ids = get_ids(last_file)

    # 环境变量
    os.environ["EMAIL_USER"] = email
    os.environ["EMAIL_PASS"] = pwd
    os.environ["START_DATE"] = start_date
    os.environ["END_DATE"] = end_date

    # 收邮件
    client = EmailClient()
    mails = client.fetch_mails()
    st.success(f"📩 共收取邮件：{len(mails)}封")

    # 筛选逻辑
    accept_list = []
    reject_list = []
    processed = set()

    for mail in mails:
        sid = mail.get("student_id", "")
        name = mail.get("name", "")
        rt = mail.get("receive_time", "")
        ap = mail.get("attachment_path", "")

        g, c, is_sub, cnt, err, ok = "未知", "未知", False, 0, "", True
        try:
            t = parsedate_to_datetime(rt)
        except:
            t = None

        if ap:
            try:
                p = DocxParser(ap)
                g = p.get_grade()
                c = p.get_apply_class()
                is_sub = p.get_subsidy()
                cnt = p.get_reason_count()
            except:
                err = "文件解析失败"

        if sid in black_ids:
            err = "黑名单"
        elif sid in last_ids:
            err = "本年已参加"
        elif not ap:
            err = "无Word附件"
        elif not is_sub:
            err = "非资助对象"
        elif cnt < 100:
            err = f"字数不足({cnt})"

        if not err:
            if sid in processed:
                err = "重复报名"
            else:
                processed.add(sid)

        row = [sid, name, g, cnt, "是" if is_sub else "否", c, t]
        if err:
            reject_list.append(row + [err])
        else:
            accept_list.append(row)

    # 排序
    accept_list.sort(key=lambda x: (x[5], x[6] if x[6] else ""))
    reject_list.sort(key=lambda x: (x[5], x[6] if x[6] else ""))

    # 构造表格
    cols1 = ["学号","姓名","年级","申请理由字数","是否资助","报名班级"]
    cols2 = ["学号","姓名","年级","申请理由字数","是否资助","报名班级","拒绝原因"]
    df_a = pd.DataFrame([x[:6] for x in accept_list], columns=cols1)
    df_r = pd.DataFrame([x[:7] for x in reject_list], columns=cols2)

    # 输出结果
    st.success(f"✅ 筛选完成！录取 {len(df_a)} 人，拒绝 {len(df_r)} 人")

    st.subheader("✅ 录取名单")
    st.write(df_a)

    st.subheader("❌ 拒绝名单")
    st.write(df_r)

    # 下载
    df_a.to_excel("录取.xlsx", index=False)
    df_r.to_excel("拒绝.xlsx", index=False)

    col1, col2 = st.columns(2)
    with col1:
        st.download_button("下载录取名单", open("录取.xlsx","rb"), "录取名单.xlsx")
    with col2:
        st.download_button("下载拒绝名单", open("拒绝.xlsx","rb"), "拒绝名单.xlsx")
