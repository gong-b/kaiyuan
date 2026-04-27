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

# 初始化Session State
if "audit_result" not in st.session_state:
    st.session_state.audit_result = {
        "ok_final": [],
        "no_final": [],
        "total": 0
    }

st.set_page_config(page_title="开源课堂报名", layout="wide")
logging.basicConfig(level=logging.ERROR)

# 模块导入
try:
    from modules.config import Config
    from modules.email_parser import EmailParser
    from modules.email_client import SecureIMAPClient
    from modules.excel_handler import ExcelHandler
    from modules.file_parser import FileParser
except Exception as e:
    st.error(f"模块加载失败: {str(e)}")
    st.stop()

ep = EmailParser()
eh = ExcelHandler()
dp = FileParser()

st.title("开源课堂报名审核")
st.divider()

# 文件上传
c1, c2, c3 = st.columns(3)
with c1:
    hongji = st.file_uploader("📋 新鸿基名单 Excel", type="xlsx")
with c2:
    last = st.file_uploader("📋 去年录取名单 Excel", type="xlsx")
with c3:
    blacklist = st.file_uploader("🚫 黑名单 Excel", type="xlsx")

st.subheader("📧 邮箱配置")
ca, cb = st.columns(2)
with ca:
    user = st.text_input("邮箱账号")
    pwd = st.text_input("授权码", type="password")
with cb:
    folder = st.text_input("文件夹", value="开源课堂")
    s_date = st.date_input("开始日期", datetime(2026,3,1))
    e_date = st.date_input("截止日期", datetime(2026,5,1))

# 核心审核逻辑
if st.button("🚀 开始审核", disabled=not (user and pwd)):
    with st.spinner("正在解析..."):
        H = eh.read_student_list(hongji) if hongji else set()
        L = eh.read_student_list(last) if last else set()
        B = eh.read_student_list(blacklist) if blacklist else set()
        
        ok_final = []
        no_final = []
        student_records = {}

        try:
            with SecureIMAPClient(user, pwd, folder) as client:
                mails = list(client.fetch_emails(s_date.strftime("%d-%b-%Y")))
                total = len(mails)
                bar = st.progress(0, text="准备解析...")
                
                for idx, (uid, msg) in enumerate(mails):
                    try:
                        sender_email = parseaddr(msg.get("From", ""))[1]
                        if sender_email == user: continue
                        subj = ep.parse_subject(msg)
                        if any(prefix in subj[:5].upper() for prefix in ["RE:", "FW:", "回复:", "转发:"]): continue

                        with tempfile.TemporaryDirectory() as tmp:
                            docs = ep.extract_attachments(msg, Path(tmp))
                            if not docs: continue

                            # 解析附件
                            info = dp.parse(str(docs[0]))
                            f_name = info.get("name", "未知")
                            f_sid = str(info.get("sid", "")).strip()
                            apply_class = info.get("apply_class", "")
                            contact = info.get("contact", "")
                            reason_len = info.get("reason_length", 0)  # 理由字数

                            if not apply_class:
                                cm = re.search(r"([^+、\s]+班)", subj)
                                apply_class = cm.group(1) if cm else "未知班级"

                            # 时间
                            try:
                                d_utc = parsedate_to_datetime(msg["Date"])
                                d_local = d_utc.astimezone()
                            except:
                                d_local = datetime.now()

                            current_record = None
                            if not f_sid:
                                current_record = {
                                    "name":f_name,"sid":"缺失","class":apply_class,"status":"reject",
                                    "reason":"未填写学号","subject":subj,"date":d_local,
                                    "contact":contact,"reason_len":reason_len
                                }
                            else:
                                if f_sid in B:
                                    current_record = {"name":f_name,"sid":f_sid,"class":apply_class,"status":"reject","reason":"黑名单","date":d_local,"contact":contact,"reason_len":reason_len}
                                elif f_sid in H:
                                    current_record = {"name":f_name,"sid":f_sid,"class":apply_class,"status":"accept","remark":"新鸿基录取","date":d_local,"contact":contact,"reason_len":reason_len}
                                elif f_sid in L:
                                    current_record = {"name":f_name,"sid":f_sid,"class":apply_class,"status":"reject","reason":"去年已录取","date":d_local,"contact":contact,"reason_len":reason_len}
                                elif not info.get("is_supported", False):
                                    current_record = {"name":f_name,"sid":f_sid,"class":apply_class,"status":"reject","reason":"非资助对象","date":d_local,"contact":contact,"reason_len":reason_len}
                                elif reason_len < Config.MIN_REASON_LENGTH:
                                    current_record = {"name":f_name,"sid":f_sid,"class":apply_class,"status":"reject","reason":f"理由不足({reason_len}字)","date":d_local,"contact":contact,"reason_len":reason_len}
                                else:
                                    current_record = {"name":f_name,"sid":f_sid,"class":apply_class,"status":"accept","remark":"审核通过","date":d_local,"contact":contact,"reason_len":reason_len}

                            # 去重：保留最早报名
                            sid_key = f_sid if f_sid != "缺失" else f"NO_{uid}"
                            if sid_key not in student_records:
                                student_records[sid_key] = current_record
                            else:
                                ex = student_records[sid_key]
                                if ex["status"] == "reject" and current_record["status"] == "accept":
                                    student_records[sid_key] = current_record
                                elif ex["status"] == current_record["status"]:
                                    if current_record["date"] < ex["date"]:
                                        student_records[sid_key] = current_record

                        bar.progress((idx+1)/total, text=f"解析中：{idx+1}/{total}")
                    except Exception as e:
                        logging.error(f"邮件{uid}错误：{e}")
                        continue

                # 生成最终名单（新增：申请理由字数）
                for sid, r in student_records.items():
                    is_hj = "新鸿基录取" in r.get("remark", "")
                    if r["status"] == "accept":
                        ok_final.append({
                            "学号": r["sid"], "姓名": r["name"], "录取班级": r["class"],
                            "联系方式": r["contact"], "是否新鸿基": "是" if is_hj else "否",
                            "申请理由字数": r.get("reason_len", 0),  # 新增列
                            "报名时间": r["date"].strftime("%Y-%m-%d %H:%M"),
                            "备注": r.get("remark", ""), "date_obj": r["date"]
                        })
                    else:
                        no_final.append({
                            "学号": r["sid"], "姓名": r["name"], "报名班级": r["class"],
                            "联系方式": r["contact"], "是否新鸿基": "是" if is_hj else "否",
                            "申请理由字数": r.get("reason_len", 0),  # 新增列
                            "原因": r["reason"], "报名时间": r["date"].strftime("%Y-%m-%d %H:%M"),
                            "原主题": r["subject"], "date_obj": r["date"]
                        })

                st.session_state.audit_result = {
                    "ok_final": ok_final, "no_final": no_final, "total": len(student_records)
                }
                st.success(f"✅ 解析完成：有效申请 {len(student_records)} 份")

        except Exception as ex:
            st.error(f"❌ 出错：{str(ex)}")

# 排序展示
def group_and_sort(data, class_key):
    groups = {}
    for item in data:
        c = item[class_key]
        if c not in groups: groups[c] = []
        groups[c].append(item)
    for c in groups:
        groups[c].sort(key=lambda x: x["date_obj"])
    sorted_names = sorted(groups.keys())
    sorted_groups = {n: groups[n] for n in sorted_names}
    flat = []
    for n in sorted_names:
        flat.extend(groups[n])
    return sorted_groups, flat

# 展示结果
if st.session_state.audit_result["total"] > 0:
    ok = st.session_state.audit_result["ok_final"]
    no = st.session_state.audit_result["no_final"]
    ok_g, ok_all = group_and_sort(ok, "录取班级")
    no_g, no_all = group_and_sort(no, "报名班级")

    col1, col2 = st.columns(2)
    with col1:
        st.subheader(f"🎯 录取名单 {len(ok)} 人")
        for cls, students in ok_g.items():
            with st.expander(f"{cls}（{len(students)}人）"):
                show = [{k:v for k,v in s.items() if k != "date_obj"} for s in students]
                st.dataframe(show, use_container_width=True)
        dl_ok = [{k:v for k,v in s.items() if k != "date_obj"} for s in ok_all]
        st.download_button("📥 下载录取名单", eh.to_excel_bytes(dl_ok), "录取名单.xlsx", use_container_width=True)

    with col2:
        st.subheader(f"❌ 拒绝名单 {len(no)} 人")
        for cls, students in no_g.items():
            with st.expander(f"{cls}（{len(students)}人）"):
                show = [{k:v for k,v in s.items() if k != "date_obj"} for s in students]
                st.dataframe(show, use_container_width=True)
        dl_no = [{k:v for k,v in s.items() if k != "date_obj"} for s in no_all]
        st.download_button("📥 下载拒绝名单", eh.to_excel_bytes(dl_no), "拒绝名单.xlsx", use_container_width=True)
