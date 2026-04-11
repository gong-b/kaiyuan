import logging
import os
from email_client import EmailClient
from docx_parser import DocxParser
import pandas as pd
from email.utils import parsedate_to_datetime
import streamlit as st
import warnings
warnings.filterwarnings("ignore")

# 页面基础配置
st.set_page_config(
    page_title="开源课堂班报名邮件筛选系统",
    page_icon="🎓",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# 自定义CSS美化页面，解决乱码、重叠问题
st.markdown("""
<style>
/* 全局样式 */
.stApp {
    background-color: #0e1117;
    color: #fafafa;
}
/* 标题样式 */
h1 {
    color: #ffffff;
    font-size: 2.5rem;
    font-weight: 700;
    margin-bottom: 2rem;
}
h2 {
    color: #ffffff;
    font-size: 1.8rem;
    font-weight: 600;
    margin: 1.5rem 0 1rem 0;
}
/* 输入框样式 */
.stTextInput > div > div > input,
.stDateInput > div > div > input,
.stFileUploader > div > div > button {
    background-color: #262730;
    color: #fafafa;
    border: 1px solid #4e4f60;
    border-radius: 8px;
    padding: 0.8rem;
}
/* 按钮样式 */
.stButton > button {
    background-color: #ff4b4b;
    color: white;
    border: none;
    border-radius: 8px;
    padding: 0.8rem 2rem;
    font-size: 1.1rem;
    font-weight: 600;
}
.stButton > button:hover {
    background-color: #ff6b6b;
}
/* 表格样式 */
.stDataFrame {
    background-color: #262730;
    border-radius: 8px;
}
/* 下载按钮样式 */
.stDownloadButton > button {
    background-color: #00cc96;
    color: white;
    border: none;
    border-radius: 8px;
    padding: 0.6rem 1.5rem;
    margin: 0.5rem 0;
}
.stDownloadButton > button:hover {
    background-color: #00e6ac;
}
</style>
""", unsafe_allow_html=True)

# 主标题
st.title("🎓 书法班报名自动筛选系统")

# 分栏布局，让页面更整齐
col1, col2 = st.columns(2)

# 左侧：邮箱登录
with col1:
    st.subheader("📩 邮箱登录")
    email = st.text_input("浙大邮箱", placeholder="请输入浙大邮箱地址")
    pwd = st.text_input("客户端专用密码", type="password", placeholder="请输入客户端专用密码")

# 右侧：筛选日期
with col2:
    st.subheader("⏰ 筛选日期")
    start_date = st.date_input("开始日期", value=None, placeholder="请选择开始日期")
    end_date = st.date_input("结束日期", value=None, placeholder="请选择结束日期")

# 名单上传区（单独一行，布局整齐）
st.subheader("📂 上传名单")
col3, col4, col5 = st.columns(3)

with col3:
    xhj_file = st.file_uploader("新鸿基名单", type=["xlsx"], label_visibility="visible")
with col4:
    black_file = st.file_uploader("黑名单", type=["xlsx"], label_visibility="visible")
with col5:
    last_file = st.file_uploader("去年已参加", type=["xlsx"], label_visibility="visible")

# 开始筛选按钮（居中）
st.markdown("<br>", unsafe_allow_html=True)
if st.button("✅ 开始筛选", use_container_width=True):
    # 校验输入
    if not email or not pwd or not start_date or not end_date or not xhj_file or not black_file or not last_file:
        st.error("❌ 请填写完整信息并上传所有名单！")
        st.stop()

    # 保存上传的名单
    with open("新鸿基名单.xlsx", "wb") as f:
        f.write(xhj_file.getbuffer())
    with open("黑名单.xlsx", "wb") as f:
        f.write(black_file.getbuffer())
    with open("去年名单.xlsx", "wb") as f:
        f.write(last_file.getbuffer())

    # 配置环境变量
    os.environ["EMAIL_USER"] = email
    os.environ["EMAIL_PASS"] = pwd
    os.environ["START_DATE"] = str(start_date)
    os.environ["END_DATE"] = str(end_date)

    # 读取名单
    def get_ids(path):
        try:
            df = pd.read_excel(path, dtype=str)
            return set(df.iloc[:, 0].dropna().str.strip())
        except Exception as e:
            st.error(f"⚠️ 读取名单 {path} 失败: {e}")
            return set()

    xhj_ids = get_ids("新鸿基名单.xlsx")
    black_ids = get_ids("黑名单.xlsx")
    last_ids = get_ids("去年名单.xlsx")

    # 收取邮件
    with st.spinner("📩 正在收取邮件..."):
        client = EmailClient()
        mails = client.fetch_mails()
    st.success(f"✅ 共收取邮件：{len(mails)}封")

    # 筛选逻辑（完全保留你原来的逻辑）
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
            except Exception as e:
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

    # 排序：先按班级 → 再按真实时间正序
    accept_list.sort(key=lambda x: (x[5], x[6] if x[6] else ""))
    reject_list.sort(key=lambda x: (x[5], x[6] if x[6] else ""))

    accept_final = [[x[0], x[1], x[2], x[3], x[4], x[5]] for x in accept_list]
    reject_final = [[x[0], x[1], x[2], x[3], x[4], x[5], x[7]] for x in reject_list]

    accept_cols = ["学号", "姓名", "年级", "申请理由字数", "是否资助", "报名班级"]
    reject_cols = ["学号", "姓名", "年级", "申请理由字数", "是否资助", "报名班级", "拒绝原因"]

    # 导出Excel
    df_accept = pd.DataFrame(accept_final, columns=accept_cols)
    df_reject = pd.DataFrame(reject_final, columns=reject_cols)

    df_accept.to_excel("录取名单.xlsx", index=False)
    df_reject.to_excel("拒绝名单.xlsx", index=False)

    # 分班导出
    if not df_accept.empty:
        for cls_name, group in df_accept.groupby("报名班级"):
            group.to_excel(f"录取_{cls_name}.xlsx", index=False)
    if not df_reject.empty:
        for cls_name, group in df_reject.groupby("报名班级"):
            group.to_excel(f"拒绝_{cls_name}.xlsx", index=False)

    # 展示结果
    st.markdown("---")
    st.subheader("🎯 筛选结果")
    st.success(f"✅ 最终录取：{len(accept_final)} 人 | ❌ 最终拒绝：{len(reject_final)} 人")

    # 录取名单
    st.subheader("✅ 录取名单")
    st.dataframe(df_accept, use_container_width=True, hide_index=True)
    with open("录取名单.xlsx", "rb") as f:
        st.download_button("📥 下载录取名单.xlsx", f, file_name="录取名单.xlsx")

    # 拒绝名单
    st.subheader("❌ 拒绝名单")
    st.dataframe(df_reject, use_container_width=True, hide_index=True)
    with open("拒绝名单.xlsx", "rb") as f:
        st.download_button("📥 下载拒绝名单.xlsx", f, file_name="拒绝名单.xlsx")

    # 分班名单下载
    st.subheader("📁 分班名单下载")
    for f in os.listdir("."):
        if f.startswith("录取_") and f.endswith(".xlsx"):
            with open(f, "rb") as fp:
                st.download_button(f"📥 下载{f}", fp, file_name=f)
        elif f.startswith("拒绝_") and f.endswith(".xlsx"):
            with open(f, "rb") as fp:
                st.download_button(f"📥 下载{f}", fp, file_name=f)

logging.basicConfig(level=logging.INFO, format="%(message)s")

if __name__ == "__main__":
    # 仅在Streamlit环境下运行main函数
    if st.runtime.exists():
        main()
