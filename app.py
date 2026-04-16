import streamlit as st
import os, subprocess
from datetime import datetime
from pathlib import Path

st.title("开源课堂报名自动审核系统")

with st.sidebar:
    email_user = st.text_input("邮箱地址", value="zzbgs@zju.edu.cn")
    email_pass = st.text_input("邮箱授权码", type="password")
    target_folder = st.text_input("提取文件夹", value="开源课堂")
    start_date = st.date_input("开始日期", value=datetime(2025, 3, 1))

col1, col2 = st.columns(2)
with col1:
    hj_file = st.file_uploader("上传新鸿基推荐名单", type=['xlsx'])
with col2:
    ly_file = st.file_uploader("上传去年录取名单", type=['xlsx'])

if st.button("开始处理", type="primary"):
    if not (hj_file and ly_file and email_pass):
        st.error("请完整填写配置并上传文件！")
    else:
        # 保存上传的文件到 data 目录供后台读取
        with open("data/new_hongji.xlsx", "wb") as f: f.write(hj_file.getbuffer())
        with open("data/last_year.xlsx", "wb") as f: f.write(ly_file.getbuffer())
        
        # 设置环境变量传参
        env = os.environ.copy()
        env["EMAIL_USER"] = email_user
        env["EMAIL_PASSWORD"] = email_pass
        env["START_DATE"] = start_date.strftime("%d-%b-%Y")
        env["TARGET_FOLDER"] = target_folder

        with st.spinner("正在执行邮箱提取与逻辑审核..."):
            result = subprocess.run(["python", "main.py"], env=env, capture_output=True, text=True)
            
            if result.returncode == 0:
                st.success("处理完成！")
                st.download_button("📥 下载录取名单", open("data/admitted_students.xlsx", "rb"), "录取名单.xlsx")
                st.download_button("📥 下载拒绝名单", open("data/rejected_students.xlsx", "rb"), "拒绝名单.xlsx")
            else:
                st.error(f"运行出错: {result.stderr}")
