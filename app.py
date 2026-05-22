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

# ========== 从主题提取姓名、学号、班级（仅作为兜底） ==========
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

# ========== 提取邮件正文中的链接 ==========
def extract_links_from_email(msg):
    """提取邮件正文中的所有http/https链接"""
    links = []
    # 正则匹配链接
    link_pattern = re.compile(r'https?://[^\s]+', re.IGNORECASE)
    
    def _extract_from_part(part):
        if part.get_content_maintype() == 'multipart':
            for subpart in part.get_payload():
                _extract_from_part(subpart)
        else:
            # 提取文本内容
            charset = part.get_content_charset() or 'utf-8'
            try:
                content = part.get_payload(decode=True).decode(charset, errors='replace')
                # 查找所有链接
                found_links = link_pattern.findall(content)
                links.extend(found_links)
            except:
                pass
    
    _extract_from_part(msg)
    # 去重并返回
    return list(set(links))

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
                            # 情况1：没有任何附件 → 检查是否有链接（兜底逻辑不变）
                            # ==============================================
                            if not docs:
                                # 提取邮件正文中的链接
                                links = extract_links_from_email(msg)
                                # 判定拒绝原因
                                if links:
                                    reason = f"仅含链接无附件（链接：{'; '.join(links[:3])}）"  # 最多显示3个链接
                                else:
                                    reason = "未上传附件且无链接"
                                
                                # 无附件时才使用主题提取的信息
                                name_from_subj, sid_from_subj, class_from_subj = extract_info_from_subject(subj)
                                record = {
                                    "name": name_from_subj,
                                    "sid": sid_from_subj,
                                    "class": class_from_subj,
                                    "status": "reject",
                                    "reason": reason,
                                    "subject": subj,
                                    "date": d_local,
                                    "contact": "",
                                    "reason_length": 0,
                                    "sender_email": sender_email  # 记录发件人邮箱
                                }
                                key = sid_from_subj if sid_from_subj != "未知" else f"NO_{uid}"
                                if key not in student_records:
                                    student_records[key] = record
                                continue

                            # ==============================================
                            # 情况2：有附件 → 优先使用附件解析的信息（核心修改）
                            # ==============================================
                            # 解析附件（优先来源）
                            attach_info = dp.parse(str(docs[0]))
                            
                            # 优先使用附件信息，为空时才用主题兜底
                            final_name = attach_info.get("name") or extract_info_from_subject(subj)[0]
                            final_sid = attach_info.get("sid") or extract_info_from_subject(subj)[1]
                            final_class = attach_info.get("apply_class") or extract_info_from_subject(subj)[2]
                            contact = attach_info.get("contact", "")
                            reason_len = attach_info.get("reason_length", 0)
                            is_supported = attach_info.get("is_supported", False)

                            # 审核规则（基于附件解析的信息）
                            if not final_sid:  # 使用附件解析的学号
                                reason = "学号缺失"
                                status = "reject"
                            elif final_sid in B:  # 黑名单校验
                                reason = "黑名单人员"
                                status = "reject"
                            elif final_sid in H:  # 新鸿基校验
                                reason = ""
                                status = "accept"
                                remark = "新鸿基录取"
                            elif final_sid in L:  # 去年录取校验
                                reason = "去年已录取"
                                status = "reject"
                            elif not is_supported:  # 资助对象校验
                                reason = "非资助对象"
                                status = "reject"
                            elif reason_len < Config.MIN_REASON_LENGTH:  # 理由字数校验
                                reason = f"理由不足({reason_len}字)"
                                status = "reject"
                            else:
                                reason = ""
                                status = "accept"
                                remark = "审核通过"

                            current_record = {
                                "name": final_name,          # 附件解析的姓名
                                "sid": final_sid,            # 附件解析的学号
                                "class": final_class,        # 附件解析的班级
                                "status": status,
                                "reason": reason,
                                "subject": subj,
                                "date": d_local,
                                "contact": contact,
                                "reason_length": reason_len,
                                "remark": remark if status == "accept" else "",
                                "sender_email": sender_email
                            }

                            # 去重：保留最早的有效记录
                            sid_key = final_sid if final_sid else f"NO_{uid}"  # 使用附件解析的学号作为key
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
                        # 解析出错 → 兜底使用主题信息
                        subj = ep.parse_subject(msg)
                        name_from_subj, sid_from_subj, class_from_subj = extract_info_from_subject(subj)
                        sender_email = parseaddr(msg.get("From", ""))[1]
                        err_record = {
                            "name": name_from_subj,
                            "sid": sid_from_subj,
                            "class": class_from_subj,
                            "status": "reject",
                            "reason": "无法解析附件",
                            "subject": subj,
                            "date": datetime.now(),
                            "contact": "",
                            "reason_length": 0,
                            "sender_email": sender_email
                        }
                        student_records[f"ERR_{uid}"] = err_record
                        logging.error(f"邮件{uid}错误：{str(e)}")
                        continue

                # 生成最终名单（逻辑不变）
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
                            "发件人邮箱": rec.get("sender_email", ""),
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
                            "发件人邮箱": rec.get("sender_email", ""),
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
