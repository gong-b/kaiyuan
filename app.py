import streamlit as st
import subprocess
import sys
import os
import pandas as pd

st.set_page_config(page_title="书法班报名筛选系统", page_icon="🎓", layout="wide")
st.title("🎓 浙江大学书法班报名自动筛选系统")

# 上传3个名单
st.subheader("📂 上传名单文件")
file_xhj = st.file_uploader("新鸿基学生名单", type=["xlsx"])
file_black = st.file_uploader("黑名单", type=["xlsx"])
file_last = st.file_uploader("去年已参加名单", type=["xlsx"])

# 运行按钮
if st.button("▶️ 开始自动筛选"):
    if not (file_xhj and file_black and file_last):
        st.warning("请上传全部3个Excel文件")
        st.stop()

    # 保存上传文件
    with open("新鸿基名单.xlsx", "wb") as f:
        f.write(file_xhj.getbuffer())
    with open("黑名单.xlsx", "wb") as f:
        f.write(file_black.getbuffer())
    with open("去年名单.xlsx", "wb") as f:
        f.write(file_last.getbuffer())

    with st.spinner("正在收取邮件、自动审核..."):
        try:
            result = subprocess.run(
                [sys.executable, "main.py"],
                capture_output=True,
                text=True,
                encoding="utf-8",
                errors="replace"
            )

            st.subheader("📝 运行日志")
            st.code(result.stdout + result.stderr)

            # 录取名单
            if os.path.exists("录取名单.xlsx"):
                st.success("✅ 筛选完成！")
                df1 = pd.read_excel("录取名单.xlsx")
                st.subheader("🥳 录取名单")
                st.dataframe(df1, use_container_width=True)
                with open("录取名单.xlsx", "rb") as f:
                    st.download_button("下载录取名单", f, "录取名单.xlsx")

            # 拒绝名单
            if os.path.exists("拒绝名单.xlsx"):
                df2 = pd.read_excel("拒绝名单.xlsx")
                st.subheader("❌ 拒绝名单")
                st.dataframe(df2, use_container_width=True)
                with open("拒绝名单.xlsx", "rb") as f:
                    st.download_button("下载拒绝名单", f, "拒绝名单.xlsx")

        except Exception as e:
            st.error(f"失败：{str(e)}")

st.info("💡 系统会自动收取浙大邮箱邮件 → 自动审核 → 自动生成名单")
