import streamlit as st
import subprocess, os, pandas as pd
from datetime import datetime
from pathlib import Path

st.set_page_config(page_title="书法班审核", layout="wide")
st.title("🖌️ 书法班自动化审核系统 (免依赖版)")

# 配置区
with st.sidebar:
    st.header("🔑 账户配置")
    email_user = st.text_input("IMAP邮箱", value="zzbgs@zju.edu.cn")
    email_pass = st.text_input("授权码", type="password")
    
    st.header("📅 筛选范围")
    sd = st.date_input("开始日期", value=datetime(2025, 3, 1))
    ed = st.date_input("截止日期", value=datetime.now())

# 文件上传区
st.header("📁 名单上传")
c1, c2, c3 = st.columns(3)
with c1: f_hj = st.file_uploader("新鸿基名单", type=['xlsx'])
with c2: f_ly = st.file_uploader("往年录取名单", type=['xlsx'])
with c3: f_bl = st.file_uploader("黑名单人员 (可选)", type=['xlsx'])

if st.button("🚀 开始自动化审核", type="primary"):
    if not (f_hj and f_ly and email_pass):
        st.error("请完整填写授权码并上传必要名单！")
    else:
        # 创建数据目录
        DATA_DIR = Path("data")
        DATA_DIR.mkdir(exist_ok=True)
        
        # 保存上传的文件
        with open(DATA_DIR / "new_hongji.xlsx", "wb") as f: f.write(f_hj.getbuffer())
        with open(DATA_DIR / "last_year.xlsx", "wb") as f: f.write(f_ly.getbuffer())
        
        if f_bl:
            with open(DATA_DIR / "blacklist.xlsx", "wb") as f: f.write(f_bl.getbuffer())
        elif (DATA_DIR / "blacklist.xlsx").exists():
            os.remove(DATA_DIR / "blacklist.xlsx") # 如果不传且存在旧的，则清理

        # 传递环境变量
        env = os.environ.copy()
        env["EMAIL_USER"] = email_user
        env["EMAIL_PASSWORD"] = email_pass
        env["START_DATE"] = sd.strftime("%Y-%m-%d")
        env["END_DATE"] = ed.strftime("%Y-%m-%d")

        with st.spinner("正在检索邮件并应用审核规则..."):
            res = subprocess.run(["python", "main.py"], env=env, capture_output=True, text=True)
            
            if res.returncode == 0:
                st.success("✅ 审核完成！")
                dl1, dl2 = st.columns(2)
                with dl1:
                    if Path("data/admitted_students.xlsx").exists():
                        st.download_button("📥 录取名单.xlsx", open("data/admitted_students.xlsx", "rb"), "录取名单.xlsx")
                with dl2:
                    if Path("data/rejected_students.xlsx").exists():
                        st.download_button("📥 拒绝名单.xlsx", open("data/rejected_students.xlsx", "rb"), "拒绝名单.xlsx")
            else:
                st.error("❌ 运行失败")
                st.code(res.stderr)
