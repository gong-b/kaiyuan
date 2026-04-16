import streamlit as st
import subprocess, os
from datetime import datetime

st.title("🎓 书法班自动化审核系统")

with st.sidebar:
    email_user = st.text_input("邮箱", value="zzbgs@zju.edu.cn")
    email_pass = st.text_input("授权码", type="password")
    # 增加截止日期
    start_date = st.date_input("开始日期", value=datetime(2025, 3, 1))
    end_date = st.date_input("截止日期", value=datetime.now())

st.header("文件上传")
c1, c2, c3 = st.columns(3)
with c1:
    f_hj = st.file_uploader("新鸿基名单", type=['xlsx'])
with c2:
    f_ly = st.file_uploader("往年录取名单", type=['xlsx'])
with c3:
    # 增加黑名单上传
    f_bl = st.file_uploader("黑名单 (可选)", type=['xlsx'])

if st.button("开始审核", type="primary"):
    if not (f_hj and f_ly and email_pass):
        st.error("请补充必要信息")
    else:
        # 保存文件
        os.makedirs("data", exist_ok=True)
        with open("data/new_hongji.xlsx", "wb") as f: f.write(f_hj.getbuffer())
        with open("data/last_year.xlsx", "wb") as f: f.write(f_ly.getbuffer())
        if f_bl:
            with open("data/blacklist.xlsx", "wb") as f: f.write(f_bl.getbuffer())
        else:
            # 如果没传黑名单，创建一个空的Excel防止报错
            import pandas as pd
            pd.DataFrame(columns=["学号"]).to_excel("data/blacklist.xlsx", index=False)

        # 传参
        env = os.environ.copy()
        env["EMAIL_USER"] = email_user
        env["EMAIL_PASSWORD"] = email_pass
        env["START_DATE"] = start_date.strftime("%Y-%m-%d")
        env["END_DATE"] = end_date.strftime("%Y-%m-%d")

        with st.spinner("后台审核中..."):
            res = subprocess.run(["python", "main.py"], env=env, capture_output=True, text=True)
            if res.returncode == 0:
                st.success("审核完成")
                st.download_button("下载录取名单", open("data/admitted_students.xlsx", "rb"), "录取名单.xlsx")
            else:
                st.error("运行失败")
                st.code(res.stderr)
