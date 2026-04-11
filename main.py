import streamlit as st
import pandas as pd
import os
from email_client import EmailClient
from docx_parser import DocxParser
from email.utils import parsedate_to_datetime
from io import BytesIO

# ------------------- 页面 -------------------
st.title("🎓 书法班报名自动筛选系统")

email_account = st.text_input("浙大邮箱")
password = st.text_input("客户端专用密码", type="password")
start_date = st.text_input("开始日期 (例：2025-10-02)")
end_date = st.text_input("结束日期 (例：2025-10-10)")

st.subheader("📂 上传名单")
xhj_file = st.file_uploader("新鸿基名单", type="xlsx")
black_file = st.file_uploader("黑名单", type="xlsx")
last_file = st.file_uploader("去年已参加名单", type="xlsx")

# ------------------- 开始筛选 -------------------
if st.button("✅ 开始筛选"):
    if not all([email_account, password, start_date, end_date, xhj_file, black_file, last_file]):
        st.warning("请填写完整信息！")
        st.stop()

    # 读取名单
    def get_ids(file):
        try:
            df = pd.read_excel(BytesIO(file.getvalue()), dtype=str)
            return set(df.iloc[:,0].dropna().str.strip())
        except:
            return set()

    xhj_ids = get_ids(xhj_file)
    black_ids = get_ids(black_file)
    last_ids = get_ids(last_file)

    # 环境变量
    os.environ["EMAIL_USER"] = email_account
    os.environ["EMAIL_PASS"] = password
    os.environ["START_DATE"] = start_date
    os.environ["END_DATE"] = end_date

    # 收邮件
    client = EmailClient()
    mails = client.fetch_mails()
    st.success(f"📩 共收取邮件：{len(mails)}封")

    accept_list = []
    reject_list = []
    processed = set()

    for mail in mails:
        sid = mail.get("student_id", "")
        name = mail.get("name", "")
        rt = mail.get("receive_time", "")
        attach_io = mail.get("attach_io")

        g, c, is_sub, cnt, err = "未知", "未知", False, 0, ""

        try:
            t = parsedate_to_datetime(rt)
        except:
            t = None

        # 从内存解析，不读本地文件
        if attach_io:
            try:
                p = DocxParser(attach_io)  # 关键修复
                g = p.get_grade()
                c = p.get_apply_class()
                is_sub = p.get_subsidy()
                cnt = p.get_reason_count()
            except:
                err = "文件解析失败"
        else:
            err = "无Word附件"

        # 筛选规则
        if sid in black_ids:
            err = "黑名单"
        elif sid in last_ids:
            err = "本年已参加"
        elif not err and not is_sub:
            err = "非资助对象"
        elif not err and cnt < 100:
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

    # 表格
    cols1 = ["学号","姓名","年级","申请理由字数","是否资助","报名班级"]
    cols2 = ["学号","姓名","年级","申请理由字数","是否资助","报名班级","拒绝原因"]
    df_a = pd.DataFrame([x[:6] for x in accept_list], columns=cols1)
    df_r = pd.DataFrame([x[:7] for x in reject_list], columns=cols2)

    # ===================== 修复：不写磁盘，内存下载 =====================
    st.success(f"✅ 筛选完成！录取 {len(df_a)} 人，拒绝 {len(df_r)} 人")

    st.subheader("✅ 录取名单")
    st.markdown(df_a.to_markdown(index=False))  # 不触发前端报错

    st.subheader("❌ 拒绝名单")
    st.markdown(df_r.to_markdown(index=False))

    # 内存生成 Excel
    buf_a = BytesIO()
    buf_r = BytesIO()
    df_a.to_excel(buf_a, index=False)
    df_r.to_excel(buf_r, index=False)

    col1, col2 = st.columns(2)
    with col1:
        st.download_button("下载录取名单", buf_a.getvalue(), "录取名单.xlsx")
    with col2:
        st.download_button("下载拒绝名单", buf_r.getvalue(), "拒绝名单.xlsx")
