# -*- coding: utf-8 -*-
import streamlit as st
import pandas as pd
import re
import imaplib
import ssl
from email import message_from_bytes
from email.message import Message
from email.header import decode_header
from email.utils import parsedate_to_datetime
from datetime import datetime, timedelta
from docx import Document
import tempfile
import os
from io import BytesIO

# 页面配置
st.set_page_config(page_title="浙大开源课堂报名审核", page_icon="📧", layout="wide")
st.title("📧 浙江大学开源课堂报名审核系统")

# --------------------------
# 工具函数
# --------------------------
def decode_subject(msg):
    try:
        return "".join(
            part.decode(charset or "utf-8", "replace") if isinstance(part, bytes) else str(part)
            for part, charset in decode_header(msg.get("Subject", ""))
        )
    except:
        return "[解码失败]"

def parse_name_id(subject):
    s = re.sub(r"\s+", "", subject)
    match = re.search(r"([\u4e00-\u9fa5]{2,}).*?(\d{8,12})", s)
    return match.groups() if match else (None, None)

def parse_docx(filepath):
    res = {"is_supported": False, "reason_length": 0}
    try:
        doc = Document(filepath)
        for row in doc.tables[0].rows:
            for i, cell in enumerate(row.cells):
                t = cell.text.strip()
                if "是否为学生资助对象" in t and i+1 < len(row.cells):
                    res["is_supported"] = "是" in row.cells[i+1].text
                if "申请理由" in t:
                    res["reason_length"] = len(re.sub(r"\s+", "", cell.text))
    except:
        pass
    return res

def save_attachments(msg, save_dir):
    paths = []
    os.makedirs(save_dir, exist_ok=True)
    for part in msg.walk():
        if part.get("Content-Disposition") is None:
            continue
        fn = part.get_filename()
        if not fn:
            continue
        try:
            fn = decode_header(fn)[0][0]
            if isinstance(fn, bytes):
                fn = fn.decode()
            p = os.path.join(save_dir, fn)
            with open(p, "wb") as f:
                f.write(part.get_payload(decode=True))
            paths.append(p)
        except:
            continue
    return paths

def load_ids(uploaded):
    if not uploaded:
        return set()
    df = pd.read_excel(uploaded)
    col = next((c for c in df.columns if "学号" in str(c)), df.columns[0])
    return set(df[col].astype(str).str.strip().tolist())

# --------------------------
# 抓取：只从 INBOX 读取（永不报错）
# --------------------------
def fetch_emails(imap_host, port, user, pwd, start_date, end_date, pb, status_text):
    mails = []
    ctx = ssl.create_default_context()
    try:
        with imaplib.IMAP4_SSL(imap_host, port, ssl_context=ctx, timeout=20) as conn:
            conn.login(user, pwd)
            # ✅ 只使用 INBOX，100% 不报错
            conn.select("INBOX", readonly=True)

            since = start_date.strftime("%d-%b-%Y")
            before = (end_date + timedelta(1)).strftime("%d-%b-%Y")
            status_text.text("搜索中...")
            res, data = conn.uid('SEARCH', None, 'SINCE', since, 'BEFORE', before)
            if not data[0]:
                status_text.text("无邮件")
                return mails

            uids = data[0].split()
            total = len(uids)
            for i, uid in enumerate(uids):
                pb.progress((i+1)/total, text=f"{i+1}/{total}")
                try:
                    _, dat = conn.uid('FETCH', uid, '(RFC822)')
                    msg = message_from_bytes(dat[0][1])
                    subject = decode_subject(msg)
                    if "开源课堂" not in subject and "报名" not in subject:
                        continue
                    name, sid = parse_name_id(subject)
                    mails.append({"sid": sid, "name": name, "subject": subject, "msg": msg})
                except:
                    continue
        pb.empty()
    except Exception as e:
        st.error(f"错误：{str(e)}")
    return mails

# --------------------------
# 状态
# --------------------------
if "mails" not in st.session_state:
    st.session_state.mails = []
if "admitted" not in st.session_state:
    st.session_state.admitted = []
if "rejected" not in st.session_state:
    st.session_state.rejected = []

# --------------------------
# 侧边栏
# --------------------------
with st.sidebar:
    st.header("⚙️ 邮箱配置")
    imap = st.text_input("IMAP", "imap.zju.edu.cn")
    port = st.number_input("端口", 993)
    user = st.text_input("浙大邮箱")
    pwd = st.text_input("客户端密码", type="password")

    st.markdown("---")
    start_date = st.date_input("开始日期", datetime(2026,3,1))
    end_date = st.date_input("截止日期", datetime(2026,4,10))

    st.markdown("---")
    f_hongji = st.file_uploader("新鸿基", type="xlsx")
    f_last = st.file_uploader("去年已参加", type="xlsx")
    f_black = st.file_uploader("黑名单", type="xlsx")

hongji = load_ids(f_hongji)
last = load_ids(f_last)
black = load_ids(f_black)

# --------------------------
# 抓取
# --------------------------
st.subheader("1️⃣ 抓取邮件（从收件箱 INBOX）")
if st.button("🔍 开始抓取"):
    if not user or not pwd:
        st.warning("请填邮箱+密码")
    else:
        pb = st.progress(0)
        tx = st.empty()
        st.session_state.mails = fetch_emails(imap, port, user, pwd, start_date, end_date, pb, tx)
        st.success(f"✅ 抓取完成：{len(st.session_state.mails)} 封")

if st.session_state.mails:
    with st.expander("查看邮件"):
        st.dataframe([{"学号":m["sid"],"姓名":m["name"],"主题":m["subject"]} for m in st.session_state.mails])

# --------------------------
# 审核
# --------------------------
st.subheader("2️⃣ 自动审核")
if st.button("✅ 开始审核"):
    if not st.session_state.mails:
        st.warning("先抓取")
    else:
        admit, reject = [], []
        tmp = tempfile.mkdtemp()
        with st.spinner("审核中..."):
            for mail in st.session_state.mails:
                sid, name, msg = mail["sid"], mail["name"], mail["msg"]
                if not sid or not name:
                    reject.append({"学号":"未知","姓名":"未知","原因":"格式错"})
                    continue
                if sid in black:
                    reject.append({"学号":sid,"姓名":name,"原因":"黑名单"})
                    continue
                if sid in last:
                    reject.append({"学号":sid,"姓名":name,"原因":"去年已参加"})
                    continue
                if sid in hongji:
                    admit.append({"学号":sid,"姓名":name,"结果":"新鸿基直接录取"})
                    continue

                atts = save_attachments(msg, os.path.join(tmp, sid))
                docxs = [f for f in atts if f.endswith(".docx")]
                if not docxs:
                    reject.append({"学号":sid,"姓名":name,"原因":"无附件"})
                    continue

                info = parse_docx(docxs[0])
                if not info["is_supported"]:
                    reject.append({"学号":sid,"姓名":name,"原因":"非资助对象"})
                elif info["reason_length"] < 95:
                    reject.append({"学号":sid,"姓名":name,"原因":f"字数不足({info['reason_length']})"})
                else:
                    admit.append({"学号":sid,"姓名":name,"结果":"通过"})

        st.session_state.admitted, st.session_state.rejected = admit, reject
        st.success(f"录取 {len(admit)}｜拒绝 {len(reject)}")

# --------------------------
# 下载
# --------------------------
st.subheader("3️⃣ 导出")
if st.session_state.admitted or st.session_state.rejected:
    t1, t2 = st.tabs(["录取", "拒绝"])
    with t1:
        dfa = pd.DataFrame(st.session_state.admitted)
        st.dataframe(dfa)
        out = BytesIO()
        with pd.ExcelWriter(out, engine="openpyxl") as w:
            dfa.to_excel(w, index=False)
        st.download_button("下载录取", out.getvalue(), "录取.xlsx")
    with t2:
        dfr = pd.DataFrame(st.session_state.rejected)
        st.dataframe(dfr)
        out2 = BytesIO()
        with pd.ExcelWriter(out2, engine="openpyxl") as w:
            dfr.to_excel(w, index=False)
        st.download_button("下载拒绝", out2.getvalue(), "拒绝.xlsx")
