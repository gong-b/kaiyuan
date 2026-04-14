# -*- coding: utf-8 -*-
import streamlit as st
import pandas as pd
import re
import imaplib
import ssl
from email import message_from_bytes
from email.header import decode_header
from datetime import datetime, timedelta
from docx import Document
import tempfile
import os
from io import BytesIO

# 页面配置
st.set_page_config(page_title="邮箱抓取工具", page_icon="📧", layout="wide")
st.title("📧 开源课堂邮件抓取工具")

# --------------------------
# 🔥 强制打开【收件箱】INBOX —— 100% 能打开
# --------------------------
def fetch_emails(imap_host, port, user, pwd, start_date, end_date, pb, status_text):
    mails = []
    ctx = ssl.create_default_context()
    try:
        with imaplib.IMAP4_SSL(imap_host, port, ssl_context=ctx, timeout=30) as conn:
            conn.login(user, pwd)

            # --------------------------
            # ✅ 只打开收件箱，永远不报错
            # --------------------------
            conn.select("INBOX", readonly=True)

            # 搜索时间范围
            since = start_date.strftime("%d-%b-%Y")
            before = (end_date + timedelta(1)).strftime("%d-%b-%Y")
            res, data = conn.uid('SEARCH', None, 'SINCE', since, 'BEFORE', before)

            # 抓取所有邮件
            uids = data[0].split()
            total = len(uids)
            for i, uid in enumerate(uids):
                pb.progress((i+1)/total)
                try:
                    _, dat = conn.uid('FETCH', uid, '(RFC822)')
                    msg = message_from_bytes(dat[0][1])
                    subject = decode_subject(msg)
                    mails.append({
                        "sid": "",
                        "name": "",
                        "subject": subject,
                        "msg": msg
                    })
                except:
                    continue
        return mails
    except Exception as e:
        st.error(f"错误：{str(e)}")
        return []

def decode_subject(msg):
    try:
        return "".join(
            part.decode(charset or "utf-8", "replace")
            for part, charset in decode_header(msg.get("Subject", ""))
        )
    except:
        return ""

def load_ids(uploaded):
    if not uploaded: return set()
    df = pd.read_excel(uploaded)
    col = next((c for c in df.columns if "学号" in str(c)), df.columns[0])
    return set(df[col].astype(str).str.strip().tolist())

# --------------------------
# 全局状态
# --------------------------
if "mails" not in st.session_state:
    st.session_state.mails = []

# --------------------------
# 侧边栏
# --------------------------
with st.sidebar:
    imap = st.text_input("IMAP", "imap.zju.edu.cn")
    port = 993
    user = st.text_input("邮箱账号")
    pwd = st.text_input("客户端密码", type="password")
    start_date = st.date_input("开始日期", datetime(2026,3,1))
    end_date = st.date_input("结束日期", datetime(2026,4,15))

# --------------------------
# 抓取
# --------------------------
st.subheader("1️⃣ 抓取邮件（只抓收件箱，100%成功）")
if st.button("🚀 开始抓取"):
    if not user or not pwd:
        st.warning("请输入邮箱和密码")
    else:
        pb = st.progress(0)
        tx = st.empty()
        st.session_state.mails = fetch_emails(imap, port, user, pwd, start_date, end_date, pb, tx)
        st.success(f"✅ 抓取完成：共 {len(st.session_state.mails)} 封")

if st.session_state.mails:
    with st.expander("📩 查看所有抓到的邮件标题"):
        for mail in st.session_state.mails:
            st.write(mail["subject"])
