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

# 初始化Session State，缓存审核结果
if "audit_result" not in st.session_state:
    st.session_state.audit_result = {
        "ok_final": [],
        "no_final": [],
        "total": 0
    }

st.set_page_config(page_title="开源课堂报名", layout="wide")
logging.basicConfig(level=logging.ERROR)

ep = EmailParser()
eh = ExcelHandler()
dp = FileParser()

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

# ========== 从主题提取姓名、学号、班级（通用所有邮件） ==========
def extract_info_from_subject(subject):
    name = "未知"
    sid = "未知"
    class_name = "未知班级"
    try:
        s = subject.strip()
        # 提取学号（10位数字）
        sid_match = re.search(r'(\d{10})', s)
        if sid_match:
            sid = sid_match.group(1)
        # 提取班级：XXX班
        class_match = re.search(r'([^\s]+班)', s)
        if class_match:
            class_name = class_match.group(1)
        # 提取姓名（非数字、开头部分）
        name_part = re.sub(r'\d+', '', s)
        name_part = re.sub(r'班.*', '', name_part)
        name_part = re.sub(r'[^\u4e00-\u9fa5]', '', name_part)
        if len(name_part) >= 2 and len(name_part) <= 5:
            name = name_part
    except:
        pass
    return name, sid, class_name

# ========== 核心审核逻辑 ==========
if st.button("🚀 开始审核", disabled=not (user and pwd)):
    with st.spinner("正在连接邮箱并扫描附件..."):
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
                        if sender_email == user:
                            continue
                        subj = ep.parse_subject(msg)
                        if any(prefix in subj[:5].upper() for prefix in ["RE:", "FW:", "回复:", "转发:"]):
                            continue

                        # 统一从主题提取信息
                        name_from_subj, sid_from_subj, class_from_subj = extract_info_from_subject(subj)

                        # 获取时间
                        try:
                            d_utc = parsedate_to_datetime(msg["Date"])
                            d_local = d_utc.astimezone()
                        except:
                            d_local = datetime.now()

                        # 提取附件
                        with tempfile.TemporaryDirectory() as tmp:
                            tmp_path = Path(tmp)
                            docs = ep.extract_attachments(msg, tmp_path)

                            # ==============================================
                            # 情况1：没有任何附件 → 拒绝：未上传附件
                            # ==============================================
                            if not docs:
                                record = {
                                    "name": name_from_subj,
                                    "sid": sid_from_subj,
                                    "class": class_from_subj,
                                    "status": "reject",
                                    "reason": "未上传附件",
                                    "subject": subj,
                                    "date": d_local,
                                    "contact": "",
                                    "reason_length": 0
                                }
                                key = sid_from_subj if sid_from_subj != "未知" else f"NO_{uid}"
                                if key not in student_records:
                                    student_records[key] = record
                                continue

                            # ==============================================
                            # 情况2：有附件 → 正常解析
                            # ==============================================
                            info = dp.parse(str(docs[0]))
                            f_name = info.get("name", name_from_subj)
                            f_sid = info.get("sid", sid_from_subj)
                            apply_class = info.get("apply_class", class_from_subj)
                            contact = info.get("contact", "")
                            reason_len = info.get("reason_length", 0)
                            is_supported = info.get("is_supported", False)

                            # 审核规则
                            if not f_sid:
                                reason = "学号缺失"
                                status = "reject"
                            elif f_sid in B:
                                reason = "黑名单人员"
                                status = "reject"
                            elif f_sid in H:
                                reason = ""
                                status = "accept"
                                remark = "新鸿基录取"
                            elif f_sid in L:
                                reason = "去年已录取"
                                status = "reject"
                            elif not is_supported:
                                reason = "非资助对象"
                                status = "reject"
                            elif reason_len < Config.MIN_REASON_LENGTH:
                                reason = f"理由不足({reason_len}字)"
                                status = "reject"
                            else:
                                reason = ""
                                status = "accept"
                                remark = "审核通过"

                            current_record = {
                                "name": f_name,
                                "sid": f_sid if f_sid else "缺失",
                                "class": apply_class,
                                "status": status,
                                "reason": reason,
                                "subject": subj,
                                "date": d_local,
                                "contact": contact,
                                "reason_length": reason_len,
                                "remark": remark if status == "accept" else ""
                            }

                            # 去重：保留最早的有效记录
                            sid_key = f_sid if f_sid else f"NO_{uid}"
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
                        # 解析出错 → 无法解析附件
                        name_from_subj, sid_from_subj, class_from_subj = extract_info_from_subject(subj)
                        err_record = {
                            "name": name_from_subj,
                            "sid": sid_from_subj,
                            "class": class_from_subj,
                            "status": "reject",
                            "reason": "无法解析附件",
                            "subject": subj,
                            "date": datetime.now(),
                            "contact": "",
                            "reason_length": 0
                        }
                        student_records[f"ERR_{uid}"] = err_record
                        logging.error(f"邮件{uid}错误：{str(e)}")
                        continue

                # 生成最终名单
                for rec in student_records.values():
                    is_hj = "新鸿基录取" in rec.get("remark", "")
                    if rec["status"] == "accept":
                        ok_final.append({
                            "学号": rec["sid"],
                            "姓名": rec["name"],
                            "录取班级": rec["class"],
                            "联系方式": rec["contact"],
                            "是否新鸿基": "是" if is_hj else "否",
                            "申请理由字数": rec["reason_length"],
                            "报名时间": rec["date"].strftime("%Y-%m-%d %H:%M"),
                            "备注": rec.get("remark", ""),
                            "date_obj": rec["date"]
                        })
                    else:
                        no_final.append({
                            "学号": rec["sid"],
                            "姓名": rec["name"],
                            "报名班级": rec["class"],
                            "原因": rec["reason"],
                            "联系方式": rec["contact"],
                            "是否新鸿基": "否",
                            "申请理由字数": rec["reason_length"],
                            "报名时间": rec["date"].strftime("%Y-%m-%d %H:%M"),
                            "原主题": rec["subject"],
                            "date_obj": rec["date"]
                        })

                st.session_state.audit_result = {
                    "ok_final": ok_final,
                    "no_final": no_final,
                    "total": len(student_records)
                }
                st.success(f"✅ 处理完成：共 {len(student_records)} 条记录")

        except Exception as ex:
            st.error(f"❌ 运行出错：{str(ex)}")

# ========== 分组排序 ==========
def group_and_sort(data, class_key):
    groups = {}
    for item in data:
        c = item[class_key]
        if c not in groups:
            groups[c] = []
        groups[c].append(item)
    for c in groups:
        groups[c].sort(key=lambda x: x["date_obj"])
    sorted_names = sorted(groups.keys())
    flat = []
    for name in sorted_names:
        flat.extend(groups[name])
    return {c: groups[c] for c in sorted_names}, flat

# ========== 展示结果 ==========
if st.session_state.audit_result["total"] > 0:
    ok_data = st.session_state.audit_result["ok_final"]
    no_data = st.session_state.audit_result["no_final"]

    ok_group, ok_all = group_and_sort(ok_data, "录取班级")
    no_group, no_all = group_and_sort(no_data, "报名班级")

    col1, col2 = st.columns(2)
    with col1:
        st.subheader(f"🎯 录取名单 {len(ok_data)} 人")
        for cls, students in ok_group.items():
            with st.expander(f"{cls}（{len(students)}人）"):
                show = [{k: v for k, v in s.items() if k != "date_obj"} for s in students]
                st.dataframe(show, use_container_width=True)
        dl_ok = [{k: v for k, v in s.items() if k != "date_obj"} for s in ok_all]
        st.download_button("📥 下载录取名单", eh.to_excel_bytes(dl_ok), "录取名单.xlsx", use_container_width=True)

    with col2:
        st.subheader(f"❌ 拒绝名单 {len(no_data)} 人")
        for cls, students in no_group.items():
            with st.expander(f"{cls}（{len(students)}人）"):
                show = [{k: v for k, v in s.items() if k != "date_obj"} for s in students]
                st.dataframe(show, use_container_width=True)
        dl_no = [{k: v for k, v in s.items() if k != "date_obj"} for s in no_all]
        st.download_button("📥 下载拒绝名单", eh.to_excel_bytes(dl_no), "拒绝名单.xlsx", use_container_width=True)
