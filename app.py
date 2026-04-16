import streamlit as st
import subprocess, os, shutil
from datetime import datetime
from pathlib import Path

st.set_page_config(page_title="书法班审核系统", layout="wide")
st.title("🎓 书法班自动化审核系统 (精简版)")

# 输入区
with st.sidebar:
    st.header("1. 账户配置")
    email_user = st.text_input("邮箱地址", value="zzbgs@zju.edu.cn")
    email_pass = st.text_input("邮箱授权码", type="password")
    start_date = st.date_input("邮件搜索开始日期", value=datetime(2025, 3, 1))

st.header("2. 名单上传")
c1, c2 = st.columns(2)
with c1:
    f_hj = st.file_uploader("上传新鸿基推荐名单 (.xlsx)", type=['xlsx'])
with c2:
    f_ly = st.file_uploader("上传去年录取名单 (.xlsx)", type=['xlsx'])

if st.button("🚀 开始自动化审核", type="primary"):
    if not (f_hj and f_ly and email_pass):
        st.error("❌ 请先完成配置并上传必要的名单文件！")
    else:
        # 保存文件
        with open("data/new_hongji.xlsx", "wb") as f: f.write(f_hj.getbuffer())
        with open("data/last_year.xlsx", "wb") as f: f.write(f_ly.getbuffer())
        
        # 传递环境变量
        env = os.environ.copy()
        env["EMAIL_USER"] = email_user
        env["EMAIL_PASSWORD"] = email_pass
        env["START_DATE"] = start_date.strftime("%d-%b-%Y")

        with st.spinner("正在抓取邮件并进行规则比对..."):
            # 运行后台任务
            process = subprocess.run(["python", "main.py"], env=env, capture_output=True, text=True)
            
            if process.returncode == 0:
                st.success("✅ 审核任务圆满完成！")
                # 显示下载按钮
                if Path("data/admitted_students.xlsx").exists():
                    st.download_button("📥 下载：录取名单.xlsx", open("data/admitted_students.xlsx", "rb"), "录取名单.xlsx")
                if Path("data/rejected_students.xlsx").exists():
                    st.download_button("📥 下载：拒绝名单.xlsx", open("data/rejected_students.xlsx", "rb"), "拒绝名单.xlsx")
                
                # 可选：显示日志
                with st.expander("查看处理日志"):
                    st.code(process.stdout)
            else:
                st.error("❌ 处理过程中出现异常")
                st.code(process.stderr)
