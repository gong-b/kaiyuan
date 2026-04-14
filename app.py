# -*- coding: utf-8 -*-
# Streamlit 入口：页面UI，调用主逻辑
import streamlit as st
from datetime import datetime
from main import KaiYuanAuditSystem
from config import DEFAULT_START_DATE, DEFAULT_END_DATE

# 页面配置
st.set_page_config(
    page_title="浙江大学开源课堂报名自动审核系统",
    page_icon="📧",
    layout="wide"
)
st.title("📧 浙江大学开源课堂报名自动审核系统")

# 初始化系统（单例）
if "audit_system" not in st.session_state:
    st.session_state.audit_system = KaiYuanAuditSystem()
audit_system = st.session_state.audit_system

# --------------------------
# 侧边栏：配置区
# --------------------------
with st.sidebar:
    st.header("⚙️ 邮箱配置")
    user = st.text_input("浙大邮箱")
    pwd = st.text_input("客户端密码", type="password")

    st.markdown("---")
    st.header("📅 时间范围")
    start_date = st.date_input("开始日期", datetime.strptime(DEFAULT_START_DATE, "%Y-%m-%d"))
    end_date = st.date_input("截止日期", datetime.strptime(DEFAULT_END_DATE, "%Y-%m-%d"))

    st.markdown("---")
    st.header("📋 名单上传")
    f_hongji = st.file_uploader("新鸿基名单", type="xlsx")
    f_last = st.file_uploader("去年已参加名单", type="xlsx")
    f_black = st.file_uploader("黑名单", type="xlsx")

    # 加载名单按钮
    if st.button("📥 加载名单"):
        if audit_system.load_lists(f_hongji, f_last, f_black):
            st.success("✅ 名单加载完成")

# --------------------------
# 主页面：功能区
# --------------------------
# 1. 抓取邮件
st.subheader("1️⃣ 抓取报名邮件")
if st.button("🔍 开始抓取邮件"):
    if not user or not pwd:
        st.warning("⚠️ 请填写完整的浙大邮箱和客户端密码")
    else:
        progress_bar = st.progress(0)
        status_text = st.empty()
        status, msg = audit_system.fetch_mails(
            user, pwd, start_date, end_date, progress_bar, status_text
        )
        if status:
            st.success(msg)
        else:
            st.error(msg)
        progress_bar.empty()
        status_text.empty()

# 展示抓取到的邮件
if audit_system.filtered_mails:
    with st.expander("📋 查看抓取到的报名邮件"):
        mail_list = [{
            "学号": m["sid"] or "未知",
            "姓名": m["name"] or "未知",
            "邮件主题": m["subject"]
        } for m in audit_system.filtered_mails]
        st.dataframe(mail_list, use_container_width=True)

# 2. 自动审核
st.subheader("2️⃣ 自动审核报名")
if st.button("✅ 开始自动审核"):
    if not audit_system.filtered_mails:
        st.warning("⚠️ 请先抓取邮件")
    else:
        with st.spinner("🔍 正在审核中..."):
            status, msg = audit_system.audit_mails()
            if status:
                st.success(msg)
            else:
                st.error(msg)

# 3. 结果导出
st.subheader("3️⃣ 审核结果与导出")
if audit_system.admit_list or audit_system.reject_list:
    tab1, tab2 = st.tabs(["✅ 录取名单", "❌ 拒绝名单"])
    with tab1:
        st.dataframe(audit_system.admit_list, use_container_width=True)
        admit_data, _ = audit_system.export_results()[0]
        st.download_button(
            label="📥 下载录取名单",
            data=admit_data,
            file_name="录取名单.xlsx"
        )
    with tab2:
        st.dataframe(audit_system.reject_list, use_container_width=True)
        reject_data, _ = audit_system.export_results()[1]
        st.download_button(
            label="📥 下载拒绝名单",
            data=reject_data,
            file_name="拒绝名单.xlsx"
        )
else:
    st.info("ℹ️ 点击「开始自动审核」查看结果")

st.markdown("---")
st.caption("审核规则：新鸿基直接录取 → 黑名单/去年参加过 → 拒绝 → 非资助对象 → 拒绝 → 理由<95字 → 拒绝")
