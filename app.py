import streamlit as st
import subprocess
import sys
import os
import pandas as pd
from datetime import datetime

st.set_page_config(page_title="书法班筛选", page_icon="🎓", layout="wide")
st.title("🎓 书法班报名自动筛选系统")

st.subheader("📩 邮箱登录")
email = st.text_input("浙大邮箱")
pwd = st.text_input("客户端专用密码", type="password")

st.subheader("⏰ 筛选日期")
d1 = st.date_input("开始日期")
d2 = st.date_input("结束日期")

st.subheader("📂 上传名单")
f1 = st.file_uploader("新鸿基名单", type="xlsx")
f2 = st.file_uploader("黑名单", type="xlsx")
f3 = st.file_uploader("去年已参加", type="xlsx")

if st.button("▶️ 开始筛选"):
    if not all([email, pwd, f1, f2, f3]):
        st.warning("请填完所有信息")
        st.stop()

    with open("新鸿基名单.xlsx", "wb") as f:
        f.write(f1.getbuffer())
    with open("黑名单.xlsx", "wb") as f:
        f.write(f2.getbuffer())
    with open("去年名单.xlsx", "wb") as f:
        f.write(f3.getbuffer())

    os.environ["EMAIL_USER"] = email
    os.environ["EMAIL_PASS"] = pwd
    os.environ["START_DATE"] = str(d1)
    os.environ["END_DATE"] = str(d2)

    with st.spinner("运行中..."):
        res = subprocess.run(
            [sys.executable, "main.py"],
            capture_output=True,
            encoding="utf-8",
            errors="replace"
        )

        st.code(res.stdout + res.stderr)

        for fn in ["录取名单.xlsx", "拒绝名单.xlsx"]:
            if os.path.exists(fn):
                df = pd.read_excel(fn)
                st.dataframe(df)
                with open(fn, "rb") as f:
                    st.download_button(f"下载{fn}", f, fn)
