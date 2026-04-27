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
# 注意：请确保 config/email_parser/email_client/excel_handler 模块存在
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
        # 读取辅助名单
        H = eh.read_student_list(hongji) if hongji else set()
        L = eh.read_student_list(last) if last else set()
        B = eh.read_student_list(blacklist) if blacklist else set()
        
        ok_final = []
        no_final = []
        student_records = {}  # 主键：学号 | 无学号时：NO_SID_UID
        student_first_apply = {}  # 新增：记录学生首次报名的班级和时间（解决多班级录取逻辑）

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
                                f_phone = info.get("phone", "")  # 新增：提取手机号
                                apply_class = info.get("apply_class", "")
                                if not apply_class:
                                    class_match = re.search(r"([^+、\s]+班)", subj)
                                    apply_class = class_match.group(1).strip() if class_match else "未知班级"

                                # 4. 处理邮件时间
                                try:
                                    d_utc = parsedate_to_datetime(msg["Date"])
                                    d_local = d_utc.astimezone()
                                    if not (s_date <= d_local.date() <= e_date): continue
                                except: 
                                    d_local = datetime.now()

                                # 5. 审核逻辑
                                current_record = None
                                if not f_sid:
                                    current_record = {
                                        "name": f_name, 
                                        "sid": "缺失", 
                                        "class": apply_class,
                                        "phone": f_phone,  # 新增：手机号
                                        "is_hongji": "否", # 新增：是否新鸿基
                                        "status": "reject", 
                                        "reason": "报名表内未填写学号",
                                        "subject": subj, 
                                        "date": d_local
                                    }
                                else:
                                    # 标记是否新鸿基
                                    is_hongji = "是" if f_sid in H else "否"

                                    # 自动化审核规则
                                    if f_sid in B:
                                        current_record = {
                                            "name": f_name, 
                                            "sid": f_sid, 
                                            "class": apply_class,
                                            "phone": f_phone,
                                            "is_hongji": is_hongji,
                                            "status": "reject", 
                                            "reason": "黑名单人员", 
                                            "subject": subj, 
                                            "date": d_local
                                        }
                                    elif f_sid in H:
                                        current_record = {
                                            "name": f_name, 
                                            "sid": f_sid, 
                                            "class": apply_class,
                                            "phone": f_phone,
                                            "is_hongji": is_hongji,
                                            "status": "accept", 
                                            "reason": "", 
                                            "remark": "新鸿基录取", 
                                            "date": d_local
                                        }
                                    elif f_sid in L:
                                        current_record = {
                                            "name": f_name, 
                                            "sid": f_sid, 
                                            "class": apply_class,
                                            "phone": f_phone,
                                            "is_hongji": is_hongji,
                                            "status": "reject", 
                                            "reason": "去年已录取", 
                                            "subject": subj, 
                                            "date": d_local
                                        }
                                    elif not info.get("is_supported", False):
                                        current_record = {
                                            "name": f_name, 
                                            "sid": f_sid, 
                                            "class": apply_class,
                                            "phone": f_phone,
                                            "is_hongji": is_hongji,
                                            "status": "reject", 
                                            "reason": "非资助对象", 
                                            "subject": subj, 
                                            "date": d_local
                                        }
                                    elif info.get("reason_length", 0) < Config.MIN_REASON_LENGTH:
                                        current_record = {
                                            "name": f_name, 
                                            "sid": f_sid, 
                                            "class": apply_class,
                                            "phone": f_phone,
                                            "is_hongji": is_hongji,
                                            "status": "reject", 
                                            "reason": f"理由不足({info['reason_length']}字)", 
                                            "subject": subj, 
                                            "date": d_local
                                        }
                                    else:
                                        current_record = {
                                            "name": f_name, 
                                            "sid": f_sid, 
                                            "class": apply_class,
                                            "phone": f_phone,
                                            "is_hongji": is_hongji,
                                            "status": "accept", 
                                            "reason": "", 
                                            "remark": "审核通过", 
                                            "date": d_local
                                        }

                                # 6. 修复多班级报名逻辑：保留首次报名的班级（早报名优先）
                                if current_record:
                                    sid_key = f_sid if f_sid and f_sid != "缺失" else f"NO_SID_{uid}"
                                    
                                    # 首次报名：直接记录
                                    if sid_key not in student_records:
                                        student_records[sid_key] = current_record
                                        # 记录首次报名信息（仅针对录取状态）
                                        if current_record["status"] == "accept":
                                            student_first_apply[sid_key] = {
                                                "class": apply_class,
                                                "date": d_local
                                            }
                                    else:
                                        existing = student_records[sid_key]
                                        # 规则1：已拒绝 → 新记录是录取 → 更新（但班级保留首次报名的）
                                        if existing["status"] == "reject" and current_record["status"] == "accept":
                                            # 检查是否有首次报名记录，有则用首次班级
                                            if sid_key in student_first_apply:
                                                current_record["class"] = student_first_apply[sid_key]["class"]
                                            student_records[sid_key] = current_record
                                        # 规则2：同状态 → 保留更早的记录（拒绝/录取都保留早的）
                                        elif existing["status"] == current_record["status"]:
                                            if current_record["date"] < existing["date"]:
                                                student_records[sid_key] = current_record
                                                # 更新首次报名记录
                                                if current_record["status"] == "accept":
                                                    student_first_apply[sid_key] = {
                                                        "class": apply_class,
                                                        "date": d_local
                                                    }

                            bar.progress((idx+1)/total, text=f"解析中：{idx+1}/{total} 封")
                        except Exception as e:
                            logging.error(f"邮件 {uid} 处理失败: {e}")
                            continue

                    # 生成最终名单（保留日期对象用于排序）
                    for sid, record in student_records.items():
                        if record["status"] == "accept":
                            ok_final.append({
                                "学号": record["sid"], 
                                "姓名": record["name"], 
                                "录取班级": record["class"], 
                                "联系方式": record["phone"],  # 新增：手机号
                                "是否新鸿基": record["is_hongji"],  # 新增：是否新鸿基
                                "备注": record.get("remark", ""), 
                                "报名时间": record["date"].strftime("%Y-%m-%d %H:%M"),
                                "date_obj": record["date"],
                                "class_sort": record["class"]
                            })
                        else:
                            no_final.append({
                                "学号": record["sid"], 
                                "姓名": record["name"], 
                                "报名班级": record["class"], 
                                "联系方式": record["phone"],  # 新增：手机号
                                "是否新鸿基": record["is_hongji"],  # 新增：是否新鸿基
                                "原因": record["reason"], 
                                "报名时间": record["date"].strftime("%Y-%m-%d %H:%M"),
                                "原主题": record["subject"],
                                "date_obj": record["date"],
                                "class_sort": record["class"]
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

# ========== 分组排序函数 ==========
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

# ========== 结果展示与下载 ==========
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
