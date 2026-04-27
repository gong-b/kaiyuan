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
import pandas as pd
from openpyxl import Workbook

# 模拟配置模块（补充缺失的Config类）
class Config:
    MIN_REASON_LENGTH = 50  # 申请理由最小长度要求

# 模拟Excel处理模块（补充缺失的ExcelHandler）
class ExcelHandler:
    @staticmethod
    def read_student_list(file):
        """读取Excel中的学号列表（假设第一列是学号）"""
        if not file:
            return set()
        try:
            df = pd.read_excel(file)
            return set(df.iloc[:, 0].astype(str).str.strip())
        except Exception as e:
            st.error(f"读取Excel失败: {e}")
            return set()
    
    @staticmethod
    def to_excel_bytes(data):
        """将列表数据转为Excel字节流供下载"""
        wb = Workbook()
        ws = wb.active
        if data:
            # 写入表头
            headers = list(data[0].keys())
            ws.append(headers)
            # 写入数据
            for row in data:
                ws.append([row[h] for h in headers])
        # 保存到字节流
        import io
        output = io.BytesIO()
        wb.save(output)
        output.seek(0)
        return output

# 模拟邮件解析模块（补充缺失的EmailParser）
class EmailParser:
    @staticmethod
    def parse_subject(msg):
        """解析邮件主题"""
        subject = msg.get("Subject", "")
        # 处理编码问题
        try:
            from email.header import decode_header
            decoded = decode_header(subject)
            return "".join([str(t[0], t[1] or 'utf-8') if isinstance(t[0], bytes) else str(t[0]) for t in decoded])
        except:
            return subject
    
    @staticmethod
    def extract_attachments(msg, save_dir):
        """提取邮件中的附件（docx/pdf）"""
        attachments = []
        for part in msg.walk():
            if part.get_content_maintype() == 'multipart':
                continue
            if part.get('Content-Disposition') is None:
                continue
            # 获取附件文件名
            filename = part.get_filename()
            if not filename:
                continue
            # 解码文件名
            from email.header import decode_header
            filename = decode_header(filename)
            filename = "".join([str(t[0], t[1] or 'utf-8') if isinstance(t[0], bytes) else str(t[0]) for t in filename])
            # 仅保留docx/pdf格式
            if filename.lower().endswith(('.docx', '.pdf')):
                save_path = save_dir / filename
                with open(save_path, 'wb') as f:
                    f.write(part.get_payload(decode=True))
                attachments.append(save_path)
        return attachments

# 模拟安全IMAP客户端（补充缺失的SecureIMAPClient）
class SecureIMAPClient:
    def __init__(self, user, pwd, folder):
        self.user = user
        self.pwd = pwd
        self.folder = folder
        self.client = None
    
    def __enter__(self):
        """连接邮箱"""
        try:
            self.client = imaplib.IMAP4_SSL("imap.zju.edu.cn")  # 浙大邮箱IMAP地址
            self.client.login(self.user, self.pwd)
            self.client.select(self.folder, readonly=True)
            return self
        except Exception as e:
            st.error(f"邮箱连接失败: {e}")
            raise
    
    def __exit__(self, exc_type, exc_val, exc_tb):
        """关闭连接"""
        if self.client:
            self.client.close()
            self.client.logout()
    
    def fetch_emails(self, since_date):
        """获取指定日期后的邮件"""
        # 搜索邮件：SINCE "DD-MMM-YYYY" 格式（如 01-Mar-2026）
        status, data = self.client.search(None, f'SINCE "{since_date}"')
        if status != 'OK':
            return []
        email_uids = data[0].split()
        emails = []
        for uid in email_uids:
            status, data = self.client.fetch(uid, '(RFC822)')
            if status == 'OK':
                msg = message_from_bytes(data[0][1])
                emails.append((uid, msg))
        return emails

# 导入文件解析器
from file_parser import FileParser

# ========== 初始化与页面配置 ==========
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
                            if sender_email == user:
                                continue
                            subj = ep.parse_subject(msg)
                            if any(prefix in subj[:5].upper() for prefix in ["RE:", "FW:", "回复:", "转发:"]):
                                continue

                            # 2. 提取附件（无附件则跳过）
                            with tempfile.TemporaryDirectory() as tmp:
                                tmp_path = Path(tmp)
                                docs = ep.extract_attachments(msg, tmp_path)
                                if not docs:
                                    continue 

                                # 3. 解析附件信息
                                info = FileParser.parse(str(docs[0]))
                                f_name = info.get("name", "未知").strip()
                                f_sid = str(info.get("sid", "")).strip()
                                apply_class = info.get("apply_class", "")
                                # 附件中未提取到班级则从主题补充
                                if not apply_class:
                                    class_match = re.search(r"([^+、\s]+班)", subj)
                                    apply_class = class_match.group(1).strip() if class_match else "未知班级"

                                # 4. 审核逻辑
                                current_record = None
                                if not f_sid:
                                    current_record = {
                                        "name": f_name, 
                                        "sid": "缺失", 
                                        "class": apply_class,
                                        "status": "reject", 
                                        "reason": "报名表内未填写学号",
                                        "subject": subj, 
                                        "date": datetime.now()
                                    }
                                else:
                                    # 校验邮件日期是否在范围内
                                    try:
                                        d_utc = parsedate_to_datetime(msg["Date"])
                                        d_local = d_utc.astimezone()
                                        if not (s_date <= d_local.date() <= e_date):
                                            continue
                                    except:
                                        d_local = datetime.now()

                                    # 自动化审核规则
                                    if f_sid in B:
                                        current_record = {
                                            "name": f_name, "sid": f_sid, "class": apply_class,
                                            "status": "reject", "reason": "黑名单人员", 
                                            "subject": subj, "date": d_local
                                        }
                                    elif f_sid in H:
                                        current_record = {
                                            "name": f_name, "sid": f_sid, "class": apply_class,
                                            "status": "accept", "reason": "", "remark": "新鸿基录取", 
                                            "date": d_local
                                        }
                                        student_admitted_class[f_sid] = apply_class
                                    elif f_sid in L:
                                        current_record = {
                                            "name": f_name, "sid": f_sid, "class": apply_class,
                                            "status": "reject", "reason": "去年已录取", 
                                            "subject": subj, "date": d_local
                                        }
                                    elif not info.get("is_supported", False):
                                        current_record = {
                                            "name": f_name, "sid": f_sid, "class": apply_class,
                                            "status": "reject", "reason": "非资助对象", 
                                            "subject": subj, "date": d_local
                                        }
                                    elif info.get("reason_length", 0) < Config.MIN_REASON_LENGTH:
                                        current_record = {
                                            "name": f_name, "sid": f_sid, "class": apply_class,
                                            "status": "reject", 
                                            "reason": f"理由不足({info['reason_length']}字)", 
                                            "subject": subj, "date": d_local
                                        }
                                    else:
                                        current_record = {
                                            "name": f_name, "sid": f_sid, "class": apply_class,
                                            "status": "accept", "reason": "", "remark": "审核通过", 
                                            "date": d_local
                                        }
                                        student_admitted_class[f_sid] = apply_class

                                # 5. 去重逻辑：保留更优状态/最新记录
                                if current_record:
                                    sid_key = f_sid if f_sid and f_sid != "缺失" else f"NO_SID_{uid}"
                                    if sid_key not in student_records:
                                        student_records[sid_key] = current_record
                                    else:
                                        existing = student_records[sid_key]
                                        # 拒绝→通过 则更新
                                        if existing["status"] == "reject" and current_record["status"] == "accept":
                                            student_records[sid_key] = current_record
                                        # 同状态则保留最新记录
                                        elif existing["status"] == current_record["status"]:
                                            if current_record["date"] > existing["date"]:
                                                student_records[sid_key] = current_record

                            bar.progress((idx+1)/total, text=f"解析中：{idx+1}/{total} 封")
                        except Exception as e:
                            logging.error(f"邮件 {uid} 处理失败: {e}")
                            continue

                    # 整理最终名单
                    for sid, record in student_records.items():
                        if record["status"] == "accept":
                            ok_final.append({
                                "学号": record["sid"], 
                                "姓名": record["name"], 
                                "录取班级": record["class"], 
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

# ========== 分组排序工具函数 ==========
def group_and_sort(data, class_key):
    """
    第一步：按班级分组；第二步：组内按报名时间升序；第三步：整体按班级名称排序
    :param data: 原始名单数据
    :param class_key: 班级字段名（录取名单用"录取班级"，拒绝名单用"报名班级"）
    :return: 1. 分组排序后的字典 2. 整体排序的列表（用于下载）
    """
    # 按班级分组
    class_groups = {}
    for student in data:
        cls_name = student[class_key]
        if cls_name not in class_groups:
            class_groups[cls_name] = []
        class_groups[cls_name].append(student)
    
    # 组内按报名时间升序
    for cls_name in class_groups:
        class_groups[cls_name].sort(key=lambda x: x["date_obj"], reverse=False)
    
    # 整体按班级名称排序
    sorted_class_names = sorted(class_groups.keys())
    sorted_groups = {cls: class_groups[cls] for cls in sorted_class_names}

    # 生成整体排序的列表（用于下载）
    total_sorted_list = []
    for cls in sorted_class_names:
        total_sorted_list.extend(class_groups[cls])
    
    return sorted_groups, total_sorted_list

# ========== 结果展示与下载 ==========
if st.session_state.audit_result["total"] > 0:
    ok_final = st.session_state.audit_result["ok_final"]
    no_final = st.session_state.audit_result["no_final"]

    # 录取/拒绝名单分组排序
    ok_grouped, ok_total_sorted = group_and_sort(ok_final, class_key="录取班级")
    no_grouped, no_total_sorted = group_and_sort(no_final, class_key="报名班级")

    # 分栏展示
    col_a, col_b = st.columns(2)
    with col_a:
        st.subheader(f"🎯 录取名单（总计 {len(ok_final)} 人）")
        # 按班级分组展示
        for cls_name, students in ok_grouped.items():
            with st.expander(f"{cls_name}（{len(students)} 人）"):
                # 移除排序用的临时字段
                display_data = [{k: v for k, v in s.items() if k not in ["date_obj", "class_sort"]} for s in students]
                st.dataframe(display_data, use_container_width=True)
        
        # 下载录取名单
        ok_download = [{k: v for k, v in s.items() if k not in ["date_obj", "class_sort"]} for s in ok_total_sorted]
        st.download_button(
            "📥 下载录取名单", 
            eh.to_excel_bytes(ok_download), 
            "录取表.xlsx", 
            use_container_width=True
        )

    with col_b:
        st.subheader(f"❌ 拒绝名单（总计 {len(no_final)} 人）")
        # 按班级分组展示
        for cls_name, students in no_grouped.items():
            with st.expander(f"{cls_name}（{len(students)} 人）"):
                display_data = [{k: v for k, v in s.items() if k not in ["date_obj", "class_sort"]} for s in students]
                st.dataframe(display_data, use_container_width=True)
        
        # 下载拒绝名单
        no_download = [{k: v for k, v in s.items() if k not in ["date_obj", "class_sort"]} for s in no_total_sorted]
        st.download_button(
            "📥 下载拒绝名单", 
            eh.to_excel_bytes(no_download), 
            "拒绝表.xlsx", 
            use_container_width=True
        )
