import streamlit as st
import logging
import tempfile
from datetime import datetime
from email.utils import parsedate_to_datetime
from pathlib import Path
from modules.config import Config
from modules.email_parser import EmailParser
from modules.email_client import SecureIMAPClient
from modules.excel_handler import ExcelHandler
from modules.docx_parser import DocxParser


st.set_page_config(page_title="开源课堂报名", layout="wide")
logging.basicConfig(level=logging.ERROR)

ep = EmailParser()
eh = ExcelHandler()
dp = DocxParser()

st.title("开源课堂报名审核")
st.divider()

c1, c2 = st.columns(2)
with c1:
    hongji = st.file_uploader("新鸿基名单 Excel", type="xlsx")
with c2:
    last = st.file_uploader("去年录取名单 Excel", type="xlsx")

st.subheader("📧 浙大邮箱")
ca, cb = st.columns(2)
with ca:
    user = st.text_input("邮箱账号")
    pwd = st.text_input("授权码", type="password")
with cb:
    folder = st.text_input("文件夹", value="开源课堂")
    s_date = st.date_input("开始日期", datetime(2026,3,1))
    e_date = st.date_input("截止日期", datetime(2026,5,1))

if st.button("🚀 开始审核", disabled=not (hongji and last and user and pwd)):
    with st.spinner("连接邮箱..."):
        H = eh.read_student_list(hongji)
        L = eh.read_student_list(last)
        ok = []
        no = []

        try:
            with SecureIMAPClient(user, pwd, folder) as client:
                mails = list(client.fetch_emails(s_date.strftime("%d-%b-%Y")))
                total = len(mails)
                bar = st.progress(0)

# app.py 核心循环部分
for idx, (uid, msg) in enumerate(mails):
    try:
        # 1. 基础检查：日期过滤
        d = parsedate_to_datetime(msg["Date"]).replace(tzinfo=None)
        if not (s_date <= d.date() <= e_date):
            continue

        # 2. 解析附件（核心信息来源）
        with tempfile.TemporaryDirectory() as tmp:
            docs = ep.extract_docx_attachments(msg, Path(tmp))
            
            if not docs:
                # 若无附件，提取主题显示，记录为拒绝
                subj = ep.parse_subject(msg)
                no.append({"学号": "未知", "姓名": "无附件", "原因": "缺失DOCX附件", "原主题": subj})
                continue

            # 3. 运行优化后的解析器
            info = dp.parse(str(docs[0]))
            f_name = info["name"]
            f_sid = info["sid"]

            if not f_sid:
                no.append({"学号": "未知", "姓名": f_name, "原因": "附件内无学号"})
                continue

            # 4. 黑白名单审核
            if f_sid in H:
                ok.append({"学号": f_sid, "姓名": f_name, "备注": "新鸿基(附件提取)"})
                continue

            if f_sid in L:
                no.append({"学号": f_sid, "姓名": f_name, "原因": "去年已录取"})
                continue

            # 5. 合规性审核
            if not info["is_supported"]:
                no.append({"学号": f_sid, "姓名": f_name, "原因": "非资助对象"})
            elif info["reason_length"] < Config.MIN_REASON_LENGTH:
                no.append({"学号": f_sid, "姓名": f_name, "原因": f"理由不足({info['reason_length']}字)"})
            else:
                ok.append({"学号": f_sid, "姓名": f_name, "备注": "审核通过"})

    except Exception as e:
        no.append({"学号": "?", "姓名": "?", "原因": f"解析异常: {str(e)[:20]}"})
    bar.progress((idx+1)/total)

            st.success(f"✅ 录取 {len(ok)} 人")
            st.dataframe(ok, use_container_width=True)
            st.download_button("下载录取名单", eh.to_excel_bytes(ok), "录取.xlsx")

            st.warning(f"❌ 拒绝 {len(no)} 人")
            st.dataframe(no, use_container_width=True)
            st.download_button("下载拒绝名单", eh.to_excel_bytes(no), "拒绝.xlsx")

        except Exception as ex:
            st.error(f"邮箱连接失败：{str(ex)}")
