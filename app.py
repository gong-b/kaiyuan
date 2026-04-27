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

# ========== 核心审核逻辑 ==========
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

                            # 2. 提取附件（核心：无附件则跳过）
                            with tempfile.TemporaryDirectory() as tmp:
                                tmp_path = Path(tmp)
                                docs = ep.extract_attachments(msg, tmp_path)
                                if not docs:
                                    continue 

                                # 3. 解析附件信息
                                info = FileParser.parse(str(docs[0]))
                                f_name = info.get("name", "未知")
                                f_sid = str(info.get("sid", "")).strip()
                                apply_class = info.get("apply_class", "")
                                contact = info.get("contact", "")  # 新增：提取联系方式
                                reason_length = info.get("reason_length", 0)  # 新增：提取申请理由字数
                                if not apply_class:
                                    class_match = re.search(r"([^+、\s]+班)", subj)
                                    apply_class = class_match.group(1).strip() if class_match else "未知班级"

                                # 4. 审核逻辑
                                current_record = None
                                if not f_sid:
                                    current_record = {
                                        "name": f_name, "sid": "缺失", "class": apply_class,
                                        "status": "reject", "reason": "报名表内未填写学号",
                                        "subject": subj, "date": datetime.now(),
                                        "contact": contact,  # 新增：联系方式
                                        "reason_length": reason_length  # 新增：申请理由字数
                                    }
                                else:
                                    try:
                                        d_utc = parsedate_to_datetime(msg["Date"])
                                        d_local = d_utc.astimezone()
                                        if not (s_date <= d_local.date() <= e_date): continue
                                    except: d_local = datetime.now()

                                    # 自动化审核规则
                                    if f_sid in B:
                                        current_record = {"name": f_name, "sid": f_sid, "class": apply_class, "status": "reject", "reason": "黑名单人员", "subject": subj, "date": d_local, "contact": contact, "reason_length": reason_length}
                                    elif f_sid in H:
                                        current_record = {"name": f_name, "sid": f_sid, "class": apply_class, "status": "accept", "reason": "", "remark": "新鸿基录取", "date": d_local, "contact": contact, "reason_length": reason_length}
                                        student_admitted_class[f_sid] = apply_class
                                    elif f_sid in L:
                                        current_record = {"name": f_name, "sid": f_sid, "class": apply_class, "status": "reject", "reason": "去年已录取", "subject": subj, "date": d_local, "contact": contact, "reason_length": reason_length}
                                    elif not info.get("is_supported", False):
                                        current_record = {"name": f_name, "sid": f_sid, "class": apply_class, "status": "reject", "reason": "非资助对象", "subject": subj, "date": d_local, "contact": contact, "reason_length": reason_length}
                                    elif info.get("reason_length", 0) < Config.MIN_REASON_LENGTH:
                                        current_record = {"name": f_name, "sid": f_sid, "class": apply_class, "status": "reject", "reason": f"理由不足({info['reason_length']}字)", "subject": subj, "date": d_local, "contact": contact, "reason_length": reason_length}
                                    else:
                                        current_record = {"name": f_name, "sid": f_sid, "class": apply_class, "status": "accept", "reason": "", "remark": "审核通过", "date": d_local, "contact": contact, "reason_length": reason_length}
                                        student_admitted_class[f_sid] = apply_class

                                # 5. 去重逻辑（修改核心：保留最早的符合要求的报名）
                                if current_record:
                                    sid_key = f_sid if f_sid and f_sid != "缺失" else f"NO_SID_{uid}"
                                    if sid_key not in student_records:
                                        student_records[sid_key] = current_record
                                    else:
                                        existing = student_records[sid_key]
                                        # 规则调整：
                                        # 1. 拒绝→录取：替换（录取优先级更高）
                                        # 2. 同状态：保留更早的记录（原逻辑是保留更新的，现在反过来）
                                        if existing["status"] == "reject" and current_record["status"] == "accept":
                                            student_records[sid_key] = current_record
                                        elif existing["status"] == current_record["status"]:
                                            if current_record["date"] < existing["date"]:  # 取更早的记录
                                                student_records[sid_key] = current_record

                            bar.progress((idx+1)/total, text=f"解析中：{idx+1}/{total} 封")
                        except Exception as e:
                            logging.error(f"邮件 {uid} 处理失败: {e}")
                            continue

                    # 生成最终名单（保留日期对象用于排序）
                    for sid, record in student_records.items():
                        # 新增：判断是否是新鸿基
                        is_hongji = "新鸿基录取" in record.get("remark", "")
                        if record["status"] == "accept":
                            ok_final.append({
                                "学号": record["sid"], 
                                "姓名": record["name"], 
                                "录取班级": record["class"], 
                                "备注": record.get("remark", ""), 
                                "报名时间": record["date"].strftime("%Y-%m-%d %H:%M"),
                                "date_obj": record["date"],
                                "class_sort": record["class"],
                                "联系方式": record.get("contact", ""),  # 新增：联系方式
                                "是否新鸿基": "是" if is_hongji else "否",  # 新增：是否新鸿基
                                "申请理由字数": record.get("reason_length", 0)  # 新增：申请理由字数
                            })
                        else:
                            no_final.append({
                                "学号": record["sid"], 
                                "姓名": record["name"], 
                                "报名班级": record["class"], 
                                "原因": record["reason"], 
                                "报名时间": record["date"].strftime("%Y-%m-%d %H:%M"),
                                "原主题": record["subject"],
                                "date_obj": record["date"],
                                "class_sort": record["class"],
                                "联系方式": record.get("contact", ""),  # 新增：联系方式
                                "是否新鸿基": "是" if is_hongji else "否",  # 新增：是否新鸿基
                                "申请理由字数": record.get("reason_length", 0)  # 新增：申请理由字数
                            })

                    # 缓存结果
                    st.session_state.audit_result = {
                        "ok_final": ok_final,
                        "no_final": no_final,
                        "total": len(student_records)
                    }
                    st.success(f"✅ 处理完成，找到有效申请 {len(student_records)} 份")

        except Exception as ex:
            st.error(f"❌ 运行出错：{str(ex)}")

# ========== 修复后：先按班级分类，再按时间排序 ==========
def group_and_sort(data, class_key):
    """
    第一步：按班级分组；第二步：组内按报名时间升序排序；第三步：整体按班级名称排序
    :param data: 原始名单数据
    :param class_key: 班级字段名（录取名单用"录取班级"，拒绝名单用"报名班级"）
    :return: 1. 分组排序后的字典 2. 整体按班级+时间排序的列表（用于下载）
    """
    # 第一步：按班级分组
    class_groups = {}
    for student in data:
        cls_name = student[class_key]
        if cls_name not in class_groups:
            class_groups[cls_name] = []
        class_groups[cls_name].append(student)
    
    # 第二步：每个班级内按报名时间升序排序（先报名在前）
    for cls_name in class_groups:
        class_groups[cls_name].sort(key=lambda x: x["date_obj"], reverse=False)
    
    # 第三步：按班级名称排序（保证整体展示顺序是按班级来）
    sorted_class_names = sorted(class_groups.keys())
    sorted_groups = {cls: class_groups[cls] for cls in sorted_class_names}

    # 生成整体排序的列表（用于下载：先班级，后时间）
    total_sorted_list = []
    for cls in sorted_class_names:
        total_sorted_list.extend(class_groups[cls])
    
    return sorted_groups, total_sorted_list

# ========== 结果展示与下载（修复核心） ==========
if st.session_state.audit_result["total"] > 0:
    ok_final = st.session_state.audit_result["ok_final"]
    no_final = st.session_state.audit_result["no_final"]

    # 录取名单：先按班级分组，再按时间排序
    ok_grouped, ok_total_sorted = group_and_sort(ok_final, class_key="录取班级")
    # 拒绝名单：先按班级分组，再按时间排序
    no_grouped, no_total_sorted = group_and_sort(no_final, class_key="报名班级")

    # 分栏展示
    col_a, col_b = st.columns(2)
    with col_a:
        st.subheader(f"🎯 录取名单（总计 {len(ok_final)} 人）")
        # 按班级分组展示（先班级，组内时间）
        for cls_name, students in ok_grouped.items():
            with st.expander(f"{cls_name}（{len(students)} 人）"):
                # 移除排序用的字段，只展示有用信息
                display_data = [{k: v for k, v in s.items() if k not in ["date_obj", "class_sort"]} for s in students]
                st.dataframe(display_data, use_container_width=True)
        
        # 下载的列表：先按班级排序，再按时间排序
        ok_download = [{k: v for k, v in s.items() if k not in ["date_obj", "class_sort"]} for s in ok_total_sorted]
        st.download_button("📥 下载录取名单", eh.to_excel_bytes(ok_download), "录取表.xlsx", use_container_width=True)

    with col_b:
        st.subheader(f"❌ 拒绝名单（总计 {len(no_final)} 人）")
        # 按班级分组展示（先班级，组内时间）
        for cls_name, students in no_grouped.items():
            with st.expander(f"{cls_name}（{len(students)} 人）"):
                display_data = [{k: v for k, v in s.items() if k not in ["date_obj", "class_sort"]} for s in students]
                st.dataframe(display_data, use_container_width=True)
        
        # 下载的列表：先按班级排序，再按时间排序
        no_download = [{k: v for k, v in s.items() if k not in ["date_obj", "class_sort"]} for s in no_total_sorted]
        st.download_button("📥 下载拒绝名单", eh.to_excel_bytes(no_download), "拒绝表.xlsx", use_container_width=True)
