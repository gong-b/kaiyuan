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
from modules.docx_parser import DocxParser

st.set_page_config(page_title="开源课堂报名", layout="wide")
logging.basicConfig(level=logging.ERROR)

ep = EmailParser()
eh = ExcelHandler()
dp = DocxParser()

st.title("开源课堂报名审核")
st.divider()

# ========== 第一步：文件上传 ==========
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

if st.button("🚀 开始审核", disabled=not (user and pwd)):
    with st.spinner("正在连接邮箱并扫描附件..."):
        H = eh.read_student_list(hongji) if hongji else set()
        L = eh.read_student_list(last) if last else set()
        B = eh.read_student_list(blacklist) if blacklist else set()
        
        ok_final = []
        no_final = []
        student_records = {}
        student_admitted_class = {}

        try:
            with SecureIMAPClient(user, pwd, folder) as client:
                mails = list(client.fetch_emails(s_date.strftime("%d-%b-%Y")))
                total = len(mails)
                bar = st.progress(0, text="准备解析...")
                
                if total == 0:
                    st.info("ℹ️ 未找到指定日期范围内的邮件")
                else:
                    for idx, (uid, msg) in enumerate(mails):
                        try:
                            # 1. 基础过滤：过滤自己发送、回复及转发邮件
                            sender_email = parseaddr(msg.get("From", ""))[1]
                            if sender_email == user: continue
                            subj = ep.parse_subject(msg)
                            if any(prefix in subj[:5].upper() for prefix in ["RE:", "FW:", "回复:", "转发:"]): continue

                            # 2. 提取附件（核心修改：无附件则彻底不管）
                            with tempfile.TemporaryDirectory() as tmp:
                                tmp_path = Path(tmp)
                                docs = ep.extract_docx_attachments(msg, tmp_path)
                                
                                # 【逻辑修改】：如果没有 .docx 附件，直接跳过处理下一封，不记录任何信息
                                if not docs:
                                    continue 

                                # 3. 解析附件：此时信息 100% 来自 docx
                                info = dp.parse(str(docs[0]))
                                f_name = info.get("name", "未知")
                                f_sid = str(info.get("sid", "")).strip()
                                # 优先从附件解析班级（如“日语班”），若无则从主题简单匹配
                                apply_class = info.get("apply_class", "")
                                if not apply_class:
                                    class_match = re.search(r"([^+、\s]+班)", subj)
                                    apply_class = class_match.group(1).strip() if class_match else "未知班级"

                                # 4. 审核逻辑
                                current_record = None
                                if not f_sid:
                                    current_record = {
                                        "name": f_name, "sid": "缺失", "class": apply_class,
                                        "status": "reject", "reason": "报名表内未填写学号",
                                        "subject": subj, "date": datetime.now() # 附件无日期则取当前
                                    }
                                else:
                                    # 检查日期范围（附件有效才检查日期）
                                    try:
                                        d_utc = parsedate_to_datetime(msg["Date"])
                                        d_local = d_utc.astimezone()
                                        if not (s_date <= d_local.date() <= e_date): continue
                                    except: d_local = datetime.now()

                                    # 具体的自动化审核规则
                                    if f_sid in B:
                                        current_record = {"name": f_name, "sid": f_sid, "class": apply_class, "status": "reject", "reason": "黑名单人员", "subject": subj, "date": d_local}
                                    elif f_sid in H:
                                        current_record = {"name": f_name, "sid": f_sid, "class": apply_class, "status": "accept", "reason": "", "remark": "新鸿基录取", "date": d_local}
                                        student_admitted_class[f_sid] = apply_class
                                    elif f_sid in L:
                                        current_record = {"name": f_name, "sid": f_sid, "class": apply_class, "status": "reject", "reason": "去年已录取", "subject": subj, "date": d_local}
                                    elif not info.get("is_supported", False):
                                        current_record = {"name": f_name, "sid": f_sid, "class": apply_class, "status": "reject", "reason": "非资助对象", "subject": subj, "date": d_local}
                                    elif info.get("reason_length", 0) < Config.MIN_REASON_LENGTH:
                                        current_record = {"name": f_name, "sid": f_sid, "class": apply_class, "status": "reject", "reason": f"理由不足({info['reason_length']}字)", "subject": subj, "date": d_local}
                                    else:
                                        current_record = {"name": f_name, "sid": f_sid, "class": apply_class, "status": "accept", "reason": "", "remark": "审核通过", "date": d_local}
                                        student_admitted_class[f_sid] = apply_class

                                # 5. 去重逻辑
                                if current_record:
                                    sid_key = f_sid if f_sid and f_sid != "缺失" else f"NO_SID_{uid}"
                                    if sid_key not in student_records:
                                        student_records[sid_key] = current_record
                                    else:
                                        existing = student_records[sid_key]
                                        # 录取优先
                                        if existing["status"] == "reject" and current_record["status"] == "accept":
                                            student_records[sid_key] = current_record
                                        elif existing["status"] == current_record["status"]:
                                            if current_record["date"] > existing["date"]:
                                                student_records[sid_key] = current_record

                            bar.progress((idx+1)/total, text=f"解析中：{idx+1}/{total} 封")
                        except Exception as e:
                            logging.error(f"邮件 {uid} 处理失败: {e}")
                            continue

                    # 循环结束后生成报表
                    for sid, record in student_records.items():
                        if record["status"] == "accept":
                            ok_final.append({"学号": record["sid"], "姓名": record["name"], "录取班级": record["class"], "备注": record.get("remark", ""), "报名时间": record["date"].strftime("%Y-%m-%d %H:%M")})
                        else:
                            no_final.append({"学号": record["sid"], "姓名": record["name"], "报名班级": record["class"], "原因": record["reason"], "报名时间": record["date"].strftime("%Y-%m-%d %H:%M"), "原主题": record["subject"]})

                    # 展示结果
                    st.success(f"✅ 处理完成，找到有效申请 {len(student_records)} 份")
                    col_a, col_b = st.columns(2)
                    with col_a:
                        st.write(f"录取人数：{len(ok_final)}")
                        if ok_final: st.download_button("📥 下载录取名单", eh.to_excel_bytes(ok_final), "录取表.xlsx")
                    with col_b:
                        st.write(f"拒绝人数：{len(no_final)}")
                        if no_final: st.download_button("📥 下载拒绝名单", eh.to_excel_bytes(no_final), "拒绝表.xlsx")

        except Exception as ex:
            st.error(f"❌ 运行出错：{str(ex)}")
