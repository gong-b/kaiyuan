import streamlit as st
import logging
import os
import pandas as pd
from email_client import EmailClient
from docx_parser import DocxParser
from email.utils import parsedate_to_datetime

# --------------------- 关闭所有警告 & 错误 ---------------------
st.set_option('deprecation.showfileUploaderEncoding', False)
st.set_option('deprecation.showPyplotGlobalUse', False)
import warnings
warnings.filterwarnings('ignore')
logging.basicConfig(level=logging.ERROR)

# --------------------- 极简单栏页面（永不报错） ---------------------
st.title("🎓 书法班报名自动筛选系统")

# 输入（全部用最简单组件，永不报错）
email = st.text_input("浙大邮箱")
pwd = st.text_input("客户端专用密码", type="password")
start_date = st.text_input("开始日期 (例：2025-10-02)")
end_date = st.text_input("结束日期 (例：2025-10-10)")

# 文件上传（极简，不触发动态组件）
st.subheader("📂 上传名单")
xhj = st.file_uploader("新鸿基名单", type="xlsx")
black = st.file_uploader("黑名单", type="xlsx")
last = st.file_uploader("去年已参加", type="xlsx")

# 按钮
if st.button("✅ 开始筛选"):
    if not all([email, pwd, start_date, end_date, xhj, black, last]):
        st.warning("请填写完整信息")
        st.stop()

    # 保存文件
    with open("新鸿基名单.xlsx", "wb") as f:
        f.write(xhj.getbuffer())
    with open("黑名单.xlsx", "wb") as f:
        f.write(black.getbuffer())
    with open("去年名单.xlsx", "wb") as f:
        f.write(last.getbuffer())

    # 环境变量
    os.environ["EMAIL_USER"] = email
    os.environ["EMAIL_PASS"] = pwd
    os.environ["START_DATE"] = start_date
    os.environ["END_DATE"] = end_date

    # 读取名单
    def get_ids(path):
        try:
            return set(pd.read_excel(path).iloc[:,0].dropna().astype(str).str.strip())
        except:
            return set()

    xhj_ids = get_ids("新鸿基名单.xlsx")
    black_ids = get_ids("黑名单.xlsx")
    last_ids = get_ids("去年名单.xlsx")

    # 收邮件
    client = EmailClient()
    mails = client.fetch_mails()
    st.success(f"📩 共收取邮件：{len(mails)}封")

    accept_list = []
    reject_list = []
    processed_students = set()

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

        if not reject_reason:
            if sid in processed_students:
                reject_reason = "重复报名"
            else:
                processed_students.add(sid)

        base_row = [sid, name, grade, reason_count, "是" if is_subsidy else "否", apply_class, real_datetime]
        if reject_reason:
            reject_list.append([*base_row, reject_reason])
        else:
            accept_list.append(base_row)

    # 排序
    accept_list.sort(key=lambda x: (x[5], x[6] if x[6] else ""))
    reject_list.sort(key=lambda x: (x[5], x[6] if x[6] else ""))

    accept_final = [[x[0],x[1],x[2],x[3],x[4],x[5]] for x in accept_list]
    reject_final = [[x[0],x[1],x[2],x[3],x[4],x[5],x[7]] for x in reject_list]

    cols_acc = ["学号","姓名","年级","申请理由字数","是否资助","报名班级"]
    cols_rej = ["学号","姓名","年级","申请理由字数","是否资助","报名班级","拒绝原因"]

    df_acc = pd.DataFrame(accept_final, columns=cols_acc)
    df_rej = pd.DataFrame(reject_final, columns=cols_rej)

    df_acc.to_excel("录取名单.xlsx", index=False)
    df_rej.to_excel("拒绝名单.xlsx", index=False)

    # 输出结果
    st.subheader("✅ 筛选完成")
    st.success(f"录取：{len(accept_final)} 人")
    st.success(f"拒绝：{len(reject_final)} 人")

    # 下载
    with open("录取名单.xlsx", "rb") as f:
        st.download_button("下载录取名单", f, "录取名单.xlsx")
    with open("拒绝名单.xlsx", "rb") as f:
        st.download_button("下载拒绝名单", f, "拒绝名单.xlsx")
