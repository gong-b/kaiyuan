import streamlit as st
import logging
import os
import pandas as pd
from email_client import EmailClient
from docx_parser import DocxParser
from email.utils import parsedate_to_datetime
from io import BytesIO

# 关闭所有警告
logging.getLogger("streamlit").setLevel(logging.ERROR)
st.set_option("deprecation.showfileUploaderEncoding", False)
st.set_option("deprecation.showPyplotGlobalUse", False)

# 页面标题
st.title("🎓 书法班报名自动筛选系统")

# 输入界面
email = st.text_input("浙大邮箱")
pwd = st.text_input("客户端密码", type="password")
start_date = st.text_input("开始日期（例：2025-10-02）")
end_date = st.text_input("结束日期（例：2025-10-10）")

st.subheader("📂 上传三个名单")
col1, col2, col3 = st.columns(3)
with col1:
    f1 = st.file_uploader("新鸿基名单")
with col2:
    f2 = st.file_uploader("黑名单")
with col3:
    f3 = st.file_uploader("去年名单")

# 按钮
if st.button("✅ 开始筛选"):
    if not email or not pwd or not start_date or not end_date or not f1 or not f2 or not f3:
        st.warning("请把信息填完整！")
        st.stop()

    # 从内存读取（不写磁盘，解决云端报错）
    def get_ids(file):
        try:
            df = pd.read_excel(BytesIO(file.getvalue()), dtype=str)
            return set(df.iloc[:,0].dropna().str.strip())
        except:
            return set()

    xhj_ids = get_ids(f1)
    black_ids = get_ids(f2)
    last_ids = get_ids(f3)

    # 环境变量
    os.environ["EMAIL_USER"] = email
    os.environ["EMAIL_PASS"] = pwd
    os.environ["START_DATE"] = start_date
    os.environ["END_DATE"] = end_date

    # 收邮件
    client = EmailClient()
    mails = client.fetch_mails()
    st.success(f"📩 共收取邮件：{len(mails)}封")

    accept_list = []
    reject_list = []
    processed_students = set()

    # ===================== 你原来的逻辑 100% 不变 =====================
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
        real_datetime = None

        try:
            real_datetime = parsedate_to_datetime(receive_time)
        except:
            real_datetime = None

        if attach_path:
            try:
                parser = DocxParser(attach_path)
                grade = parser.get_grade()
                apply_class = parser.get_apply_class()
                is_subsidy = parser.get_subsidy()
                reason_count = parser.get_reason_count()
                parse_success = True
            except:
                reject_reason = "文件解析失败"

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

        if not reject_reason:
            if sid in processed_students:
                reject_reason = "重复报名，仅录取最先报名的班级"
            else:
                processed_students.add(sid)

        base_row = [sid, name, grade, reason_count, "是" if is_subsidy else "否", apply_class, real_datetime]

        if reject_reason:
            reject_list.append([*base_row, reject_reason])
        else:
            accept_list.append(base_row)

    # 排序（完全不变）
    accept_list.sort(key=lambda x: (x[5], x[6] if x[6] else ""))
    reject_list.sort(key=lambda x: (x[5], x[6] if x[6] else ""))

    accept_final = [[x[0],x[1],x[2],x[3],x[4],x[5]] for x in accept_list]
    reject_final = [[x[0],x[1],x[2],x[3],x[4],x[5],x[7]] for x in reject_list]

    accept_cols = ["学号","姓名","年级","申请理由字数","是否资助","报名班级"]
    reject_cols = ["学号","姓名","年级","申请理由字数","是否资助","报名班级","拒绝原因"]

    df_accept = pd.DataFrame(accept_final, columns=accept_cols)
    df_reject = pd.DataFrame(reject_final, columns=reject_cols)

    df_accept.to_excel("录取名单.xlsx", index=False)
    df_reject.to_excel("拒绝名单.xlsx", index=False)

    # ===================== 显示结果（修复前端） =====================
    st.success(f"🎯 最终录取：{len(accept_final)} 人")
    st.error(f"❌ 最终拒绝：{len(reject_final)} 人")

    st.subheader("✅ 录取名单")
    st.write(df_accept)

    st.subheader("❌ 拒绝名单")
    st.write(df_reject)

    # 下载
    st.subheader("📥 下载")
    with open("录取名单.xlsx", "rb") as f:
        st.download_button("下载录取名单", f, "录取名单.xlsx")
    with open("拒绝名单.xlsx", "rb") as f:
        st.download_button("下载拒绝名单", f, "拒绝名单.xlsx")
