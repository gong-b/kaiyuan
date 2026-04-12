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
from pathlib import Path
from docx import Document
import tempfile
import os

# 页面配置
st.set_page_config(
    page_title="浙大开源课堂报名审核",
    page_icon="📧",
    layout="wide"
)
st.title("📧 浙江大学开源课堂报名自动审核系统")

# --------------------------
# 工具函数
# --------------------------
def decode_subject(msg: Message) -> str:
    decoded_parts = []
    for part, charset in decode_header(msg.get("Subject", "")):
        try:
            if charset:
                decoded = part.decode(charset, errors='replace')
            else:
                for enc in ['utf-8', 'gb18030', 'gbk', 'big5']:
                    try:
                        decoded = part.decode(enc)
                        break
                    except:
                        continue
                else:
                    decoded = part.decode('utf-8', errors='replace')
            decoded_parts.append(decoded)
        except:
            decoded_parts.append("[解码失败]")
    return "".join(decoded_parts)

def parse_name_id(subject: str) -> tuple[str | None, str | None]:
    s = re.sub(r"\s+", "", subject)
    pattern = re.compile(r"([\u4e00-\u9fa5]{2,}).*?(\d{8,12})")
    match = pattern.search(s)
    if match:
        return match.group(1), match.group(2)
    return None, None

def parse_docx(filepath: str) -> dict:
    result = {"is_supported": False, "reason_length": 0}
    try:
        doc = Document(filepath)
        if not doc.tables:
            return result
        table = doc.tables[0]
        reason_text = ""
        for row in table.rows:
            cells = row.cells
            for idx, cell in enumerate(cells):
                t = cell.text.strip()
                if "是否为学生资助对象" in t:
                    if idx + 1 < len(cells):
                        val = cells[idx+1].text.strip()
                        result["is_supported"] = ("是" in val) and ("不是" not in val)
                if "申请理由" in t:
                    reason_text += cell.text.strip()
        reason_text = re.sub(r"\s+", "", reason_text)
        result["reason_length"] = len(reason_text)
    except:
        pass
    return result

def fetch_emails(imap_host, port, user, pwd, start_date, end_date):
    mails = []
    ctx = ssl.create_default_context()
    try:
        with imaplib.IMAP4_SSL(imap_host, port, ssl_context=ctx) as conn:
            conn.login(user, pwd)
            conn.select("INBOX")
            
            # 构造日期范围：SINCE 开始日期 BEFORE 结束日期+1天
            since_str = start_date.strftime("%d-%b-%Y")
            before_date = end_date + timedelta(days=1)
            before_str = before_date.strftime("%d-%b-%Y")
            
            status, data = conn.uid('SEARCH', 'SINCE', since_str, 'BEFORE', before_str)
            if status != "OK" or not data[0]:
                return mails
            
            uids = data[0].split()
            for uid in uids:
                try:
                    st, msg_data = conn.uid('FETCH', uid, '(RFC822)')
                    if st != "OK" or not isinstance(msg_data[0][1], bytes):
                        continue
                        
                    msg = message_from_bytes(msg_data[0][1])
                    subject = decode_subject(msg)
                    recv_time = parsedate_to_datetime(msg.get("Date")) if msg.get("Date") else None
                    
                    # 二次精确过滤时间（防止IMAP搜索误差）
                    if recv_time:
                        mail_dt = recv_time.date()
                        if not (start_date <= mail_dt <= end_date):
                            continue
                    
                    name, sid = parse_name_id(subject)
                    mails.append({
                        "uid": uid.decode(),
                        "subject": subject,
                        "name": name,
                        "sid": sid,
                        "time": recv_time,
                        "msg": msg
                    })
                except:
                    continue
    except Exception as e:
        st.error(f"邮箱登录/抓取失败：{str(e)}")
    return mails

def save_attachments(msg, save_dir):
    paths = []
    os.makedirs(save_dir, exist_ok=True)
    for part in msg.walk():
        if part.get_content_maintype() == "multipart":
            continue
        if part.get("Content-Disposition") is None:
            continue
        fn = part.get_filename()
        if not fn:
            continue
        try:
            fn = decode_header(fn)[0][0]
            if isinstance(fn, bytes):
                fn = fn.decode("utf-8", "replace")
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
# 会话状态
# --------------------------
if "mails" not in st.session_state:
    st.session_state.mails = []
if "admitted" not in st.session_state:
    st.session_state.admitted = []
if "rejected" not in st.session_state:
    st.session_state.rejected = []

# --------------------------
# 侧边栏配置
# --------------------------
with st.sidebar:
    st.header("⚙️ 邮箱配置")
    imap = st.text_input("IMAP 服务器", "imap.zju.edu.cn")
    port = st.number_input("端口", value=993)
    user = st.text_input("浙大邮箱")
    pwd = st.text_input("客户端密码", type="password")

    st.markdown("---")
    st.header("📅 报名时间范围")
    start_date = st.date_input("开始日期", datetime(2025,3,1))
    end_date = st.date_input("截止日期", datetime(2025,3,31))

    st.markdown("---")
    st.header("📋 审核名单上传")
    f_hongji = st.file_uploader("新鸿基学生名单", type="xlsx")
    f_last = st.file_uploader("去年已参加名单", type="xlsx")
    f_black = st.file_uploader("黑名单", type="xlsx")

hongji_ids = load_ids(f_hongji)
last_ids = load_ids(f_last)
black_ids = load_ids(f_black)

# --------------------------
# 抓取邮件
# --------------------------
st.subheader("1️⃣ 抓取报名邮件")
st.caption(f"当前抓取范围：{start_date} ~ {end_date}")
if st.button("🔍 开始抓取邮件"):
    if not user or not pwd:
        st.warning("请填写完整的浙大邮箱和客户端密码")
    else:
        with st.spinner("正在抓取邮件..."):
            st.session_state.mails = fetch_emails(imap, port, user, pwd, start_date, end_date)
        st.success(f"✅ 抓取完成，共找到 {len(st.session_state.mails)} 封邮件")

if st.session_state.mails:
    with st.expander("查看抓取到的邮件列表"):
        df = pd.DataFrame([{
            "学号": x["sid"] or "未知",
            "姓名": x["name"] or "未知",
            "邮件主题": x["subject"]
        } for x in st.session_state.mails])
        st.dataframe(df, use_container_width=True)

# --------------------------
# 自动审核
# --------------------------
st.subheader("2️⃣ 自动审核报名")
if st.button("✅ 开始自动审核"):
    if not st.session_state.mails:
        st.warning("请先抓取邮件")
    else:
        admit = []
        reject = []
        tmp_dir = tempfile.mkdtemp()

        with st.spinner("正在审核中..."):
            for mail in st.session_state.mails:
                sid = mail["sid"]
                name = mail["name"]
                subject = mail["subject"]
                msg = mail["msg"]

                if not sid or not name:
                    reject.append({"学号": "未知", "姓名": "未知", "原因": "邮件主题格式错误"})
                    continue

                # 审核逻辑
                if sid in black_ids:
                    reject.append({"学号": sid, "姓名": name, "原因": "黑名单"})
                    continue
                if sid in last_ids:
                    reject.append({"学号": sid, "姓名": name, "原因": "去年已参加"})
                    continue
                if sid in hongji_ids:
                    admit.append({"学号": sid, "姓名": name, "审核结果": "新鸿基直接录取"})
                    continue

                # 附件解析
                att_path = os.path.join(tmp_dir, f"{sid}")
                attachments = save_attachments(msg, att_path)
                docx_files = [f for f in attachments if f.endswith(".docx")]

                if not docx_files:
                    reject.append({"学号": sid, "姓名": name, "原因": "未找到docx附件"})
                    continue

                doc_info = parse_docx(docx_files[0])
                if not doc_info["is_supported"]:
                    reject.append({"学号": sid, "姓名": name, "原因": "非学生资助对象"})
                elif doc_info["reason_length"] < 95:
                    reject.append({"学号": sid, "姓名": name, "原因": f"理由字数不足({doc_info['reason_length']}字)"})
                else:
                    admit.append({"学号": sid, "姓名": name, "审核结果": "审核通过"})

        st.session_state.admitted = admit
        st.session_state.rejected = reject
        st.success(f"🎯 审核完成：录取 {len(admit)} 人 | 拒绝 {len(reject)} 人")

# --------------------------
# 结果展示 + 下载
# --------------------------
st.subheader("3️⃣ 审核结果与导出")
if st.session_state.admitted or st.session_state.rejected:
    tab1, tab2 = st.tabs(["✅ 录取名单", "❌ 拒绝名单"])
    with tab1:
        df_admit = pd.DataFrame(st.session_state.admitted)
        st.dataframe(df_admit, use_container_width=True)
        st.download_button(
            label="📥 下载录取名单",
            data=df_admit.to_excel(index=False),
            file_name="录取名单.xlsx"
        )
    with tab2:
        df_reject = pd.DataFrame(st.session_state.rejected)
        st.dataframe(df_reject, use_container_width=True)
        st.download_button(
            label="📥 下载拒绝名单",
            data=df_reject.to_excel(index=False),
            file_name="拒绝名单.xlsx"
        )
else:
    st.info("点击「开始自动审核」查看结果")

st.markdown("---")
st.caption("审核规则：新鸿基直接录取 → 黑名单/去年参加过 → 拒绝 → 非资助对象 → 拒绝 → 理由<95字 → 拒绝")
