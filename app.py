import streamlit as st
import logging
import tempfile
import imaplib
import re
from datetime import datetime
from email.utils import parsedate_to_datetime, parseaddr
from email.message import Message
from email import message_from_bytes
from pathlib import Path
from modules.config import Config
from modules.email_parser import EmailParser
from modules.email_client import SecureIMAPClient
from modules.excel_handler import ExcelHandler
from modules.file_parser import FileParser

st.set_page_config(page_title="开源课堂报名", layout="wide")
logging.basicConfig(level=logging.ERROR)

ep = EmailParser()
eh = ExcelHandler()
dp = FileParser()

st.title("开源课堂报名审核")
st.divider()

# ========== 1. 初始化 Session State (解决下载刷新问题) ==========
if "ok_final" not in st.session_state:
    st.session_state.ok_final = []
if "no_final" not in st.session_state:
    st.session_state.no_final = []
if "has_processed" not in st.session_state:
    st.session_state.has_processed = False

# ========== 2. 文件上传与配置 ==========
c1, c2, c3 = st.columns(3)
with c1:
    hongji = st.file_uploader("📋 新鸿基名单 Excel（可选）", type="xlsx")
with c2:
    last = st.file_uploader("📋 去年录取名单 Excel（可选）", type="xlsx")
with c3:
    blacklist = st.file_uploader("🚫 黑名单 Excel（可选）", type="xlsx")

st.subheader("📧 浙大邮箱配置")
ca, cb = st.columns(2)
with ca:
    user = st.text_input("邮箱账号")
    pwd = st.text_input("授权码", type="password")
with cb:
    folder = st.text_input("文件夹", value="开源课堂")
    s_date = st.date_input("开始日期", datetime(2026,3,1))
    e_date = st.date_input("截止日期", datetime(2026,5,1))
if st.button("🚀 开始审核", type="primary"):
    if not user or not pwd:
        st.error("请输入账号和授权码")
    else:
        try:
            # 读取参考名单
            hj_set = eh.read_student_list(hongji) if hongji else set()
            ls_set = eh.read_student_list(last) if last else set()
            bl_set = eh.read_student_list(blacklist) if blacklist else set()

            ok_list = []
            no_list = []
            student_records = {}

            with SecureIMAPClient(user, pwd, folder) as client:
                imap_date = s_date.strftime("%d-%b-%Y")
                status, data = client.conn.uid('SEARCH', 'ALL', 'SINCE', imap_date)
                
                if status == 'OK' and data[0]:
                    uids = data[0].split()
                    total = len(uids)
                    bar = st.progress(0, text=f"准备解析 {total} 封邮件...")

                    for idx, uid in enumerate(uids):
                        try:
                            res, msg_data = client.conn.uid('FETCH', uid, '(RFC822)')
                            msg = message_from_bytes(msg_data[0][1])
                            subj = ep.parse_subject(msg)
                            date_raw = parsedate_to_datetime(msg.get("Date"))
                            
                            with tempfile.TemporaryDirectory() as tmp:
                                tmp_path = Path(tmp)
                                docs = ep.extract_docx_attachments(msg, tmp_path) # 内部已兼容PDF
                                
                                if not docs: continue
                                
                                info = dp.parse(str(docs[0]))
                                sid = info.get("sid")
                                name = info.get("name")
                                apply_class = info.get("apply_class") or "未知班级"

                                # 审核逻辑
                                status_type = "accept"
                                reason = ""
                                
                                if not sid:
                                    status_type = "reject"
                                    reason = "无法解析学号"
                                elif sid in bl_set:
                                    status_type = "reject"
                                    reason = "黑名单用户"
                                elif sid in ls_set:
                                    status_type = "reject"
                                    reason = "往年已录取"
                                elif info.get("reason_length", 0) < Config.MIN_REASON_LENGTH:
                                    status_type = "reject"
                                    reason = f"字数不足({info.get('reason_length')})"

                                current_record = {
                                    "sid": sid, "name": name, "class": apply_class,
                                    "status": status_type, "reason": reason,
                                    "date": date_raw, "subject": subj,
                                    "remark": "新鸿基" if sid in hj_set else ""
                                }

                                # 同学号去重，保留最新的
                                if sid not in student_records or date_raw > student_records[sid]["date"]:
                                    student_records[sid] = current_record

                            bar.progress((idx + 1) / total, text=f"进度：{idx + 1}/{total}")
                        except Exception: continue

                    # 分类并存入状态
                    processed_ok = []
                    processed_no = []
                    for sid, record in student_records.items():
                        if record["status"] == "accept":
                            processed_ok.append({
                                "学号": record["sid"], "姓名": record["name"], 
                                "录取班级": record["class"], "备注": record["remark"], 
                                "报名时间": record["date"].strftime("%Y-%m-%d %H:%M")
                            })
                        else:
                            processed_no.append({
                                "学号": record["sid"], "姓名": record["name"], 
                                "报名班级": record["class"], "原因": record["reason"], 
                                "报名时间": record["date"].strftime("%Y-%m-%d %H:%M")
                            })

                    # ========== 3. 排序逻辑：先按班级排，再按时间排 ==========
                    processed_ok.sort(key=lambda x: (x["录取班级"], x["报名时间"]))
                    processed_no.sort(key=lambda x: (x["报名班级"], x["报名时间"]))

                    st.session_state.ok_final = processed_ok
                    st.session_state.no_final = processed_no
                    st.session_state.has_processed = True
                    st.rerun() # 强制刷新以显示结果

        except Exception as e:
            st.error(f"发生错误: {e}")

# ========== 4. 结果展示区域 (独立于按钮外，防止刷新消失) ==========
if st.session_state.has_processed:
    st.divider()
    col_left, col_right = st.columns(2)
    
    with col_left:
        st.success(f"录取人数：{len(st.session_state.ok_final)}")
        if st.session_state.ok_final:
            st.dataframe(st.session_state.ok_final, use_container_width=True)
            st.download_button(
                "📥 下载录取名单", 
                eh.to_excel_bytes(st.session_state.ok_final), 
                "录取名单.xlsx",
                key="btn_ok"
            )

    with col_right:
        st.warning(f"拒绝人数：{len(st.session_state.no_final)}")
        if st.session_state.no_final:
            st.dataframe(st.session_state.no_final, use_container_width=True)
            st.download_button(
                "📥 下载拒绝名单", 
                eh.to_excel_bytes(st.session_state.no_final), 
                "拒绝名单.xlsx",
                key="btn_no"
            )
