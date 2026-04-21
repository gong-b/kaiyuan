import streamlit as st
import logging
import tempfile
import imaplib  # 必须导入，解决之前的NameError
import re  # 新增：用于主题提取学号
from datetime import datetime
from email.utils import parsedate_to_datetime, parseaddr
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

# ========== 第一步：修改文件上传为可选（3个文件都可选） ==========
c1, c2, c3 = st.columns(3)  # 新增一列放黑名单
with c1:
    hongji = st.file_uploader("📋 新鸿基名单 Excel（可选）", type="xlsx")
with c2:
    last = st.file_uploader("📋 去年录取名单 Excel（可选）", type="xlsx")
with c3:
    blacklist = st.file_uploader("🚫 黑名单 Excel（可选）", type="xlsx")  # 新增黑名单上传

st.subheader("📧 浙大邮箱")
ca, cb = st.columns(2)
with ca:
    user = st.text_input("邮箱账号")
    pwd = st.text_input("授权码", type="password")
with cb:
    folder = st.text_input("文件夹", value="其他文件夹/开源课堂")
    s_date = st.date_input("开始日期", datetime(2026,3,1))
    e_date = st.date_input("截止日期", datetime(2026,5,1))

# ========== 第二步：修改按钮禁用条件（仅需邮箱账号+授权码） ==========
if st.button("🚀 开始审核", disabled=not (user and pwd)):
    with st.spinner("连接邮箱..."):
        # ========== 第三步：处理可选文件（无文件则为空集合） ==========
        H = eh.read_student_list(hongji) if hongji else set()  # 新鸿基名单（可选）
        L = eh.read_student_list(last) if last else set()      # 去年录取（可选）
        B = eh.read_student_list(blacklist) if blacklist else set()  # 黑名单（新增）
        ok = []
        no = []

        try:
            with SecureIMAPClient(user, pwd, folder) as client:
                mails = list(client.fetch_emails(s_date.strftime("%d-%b-%Y")))
                total = len(mails)
                bar = st.progress(0, text="准备解析...")
                if total == 0:
                    st.info("ℹ️ 未找到指定日期范围内的邮件")
                else:
                   # 循环外新增：关键字典
student_records = {}  # 存储每个学号的最优记录
student_admitted_class = {}  # 存储：学号 → 首次录取的班级（核心！）

for idx, (uid, msg) in enumerate(mails):
    try:
        # 1. 原有过滤逻辑（跳过自己发送/回复/非目标日期邮件）
        sender_email = parseaddr(msg.get("From", ""))[1]
        if sender_email == user:
            bar.progress((idx+1)/total, text=f"解析中：{idx+1}/{total} 封（跳过自己发送的邮件）")
            continue

        subj = ep.parse_subject(msg)
        if any(prefix in subj[:5] for prefix in ["RE:", "FW:", "回复:", "转发:"]):
            bar.progress((idx+1)/total, text=f"解析中：{idx+1}/{total} 封（跳过回复/转发邮件）")
            continue

        d_utc = parsedate_to_datetime(msg["Date"])
        d_local = d_utc.astimezone()
        if not (s_date <= d_local.date() <= e_date):
            bar.progress((idx+1)/total, text=f"解析中：{idx+1}/{total} 封（跳过非目标日期邮件）")
            continue

        # 2. 提取原始邮件（兼容会话模式）
        raw_msg = msg
        if msg.is_multipart():
            for part in msg.walk():
                if part.get_content_type() == "message/rfc822":
                    raw_msg = message_from_bytes(part.get_payload(decode=True))
                    break

        # 3. 解析附件 + 提取报名班级（核心新增）
        with tempfile.TemporaryDirectory() as tmp:
            docs = ep.extract_docx_attachments(raw_msg, Path(tmp))
            f_name = "未知"
            f_sid = ""
            apply_class = "未知班级"  # 报名班级

            # 3.1 附件提取失败：从主题提取学号/姓名/班级
            if not docs:
                # 正则匹配：姓名+学号+班级（如“张三+123456+小提琴班报名申请”）
                pattern = re.search(r"([^+]+)\+(\d{8,10})\+(.*?班)", subj)
                if pattern:
                    f_name = pattern.group(1).strip()
                    f_sid = pattern.group(2).strip()
                    apply_class = pattern.group(3).strip()  # 提取班级（如“小提琴班”）
                # 临时记录：无附件
                current_record = {
                    "name": f_name,
                    "sid": f_sid,
                    "class": apply_class,
                    "status": "reject",
                    "reason": "缺失DOCX附件",
                    "subject": subj,
                    "date": d_local
                }
            else:
                # 3.2 附件提取成功：解析信息+班级
                info = dp.parse(str(docs[0]))
                f_name = info["name"]
                f_sid = info["sid"]
                
                # 优先从附件提取班级，无则从主题提取
                apply_class = info.get("apply_class", "")  # 需确保docx_parser返回apply_class（见下方说明）
                if not apply_class:
                    class_match = re.search(r"(.*?班)", subj)
                    apply_class = class_match.group(1).strip() if class_match else "未知班级"

                if not f_sid:
                    current_record = {
                        "name": f_name,
                        "sid": "",
                        "class": apply_class,
                        "status": "reject",
                        "reason": "附件内无学号",
                        "subject": subj,
                        "date": d_local
                    }
                else:
                    # ========== 核心：按“首次录取班级”审核 ==========
                    # 第一步：判断是否已录取过其他班级
                    if f_sid in student_admitted_class:
                        admitted_class = student_admitted_class[f_sid]
                        # 情况1：已录取，且本次报名同一班级→保留录取（去重）
                        if apply_class == admitted_class:
                            current_record = {
                                "name": f_name,
                                "sid": f_sid,
                                "class": apply_class,
                                "status": "accept",
                                "reason": "",
                                "remark": f"审核通过（已录取{admitted_class}）",
                                "date": d_local
                            }
                        # 情况2：已录取，本次报名其他班级→直接拒绝
                        else:
                            current_record = {
                                "name": f_name,
                                "sid": f_sid,
                                "class": apply_class,
                                "status": "reject",
                                "reason": f"重复报名（已录取{admitted_class}，本次报名{apply_class}）",
                                "subject": subj,
                                "date": d_local
                            }
                    # 第二步：未录取过，按原有规则审核
                    else:
                        if f_sid in B:  # 黑名单
                            current_record = {"name": f_name, "sid": f_sid, "class": apply_class, "status": "reject", "reason": "黑名单人员", "subject": subj, "date": d_local}
                        elif f_sid in H:  # 新鸿基
                            current_record = {"name": f_name, "sid": f_sid, "class": apply_class, "status": "accept", "reason": "", "remark": f"新鸿基(录取{apply_class})", "date": d_local}
                            student_admitted_class[f_sid] = apply_class  # 记录首次录取班级
                        elif f_sid in L:  # 去年录取
                            current_record = {"name": f_name, "sid": f_sid, "class": apply_class, "status": "reject", "reason": "去年已录取", "subject": subj, "date": d_local}
                        elif not info["is_supported"]:  # 非资助对象
                            current_record = {"name": f_name, "sid": f_sid, "class": apply_class, "status": "reject", "reason": "非资助对象", "subject": subj, "date": d_local}
                        elif info["reason_length"] < Config.MIN_REASON_LENGTH:  # 理由不足
                            current_record = {"name": f_name, "sid": f_sid, "class": apply_class, "status": "reject", "reason": f"理由不足({info['reason_length']}字)", "subject": subj, "date": d_local}
                        else:  # 审核通过
                            current_record = {"name": f_name, "sid": f_sid, "class": apply_class, "status": "accept", "reason": "", "remark": f"审核通过（录取{apply_class}）", "date": d_local}
                            student_admitted_class[f_sid] = apply_class  # 记录首次录取班级

        # ========== 去重逻辑：保留最优记录 ==========
        if f_sid and f_sid != "未知":
            if f_sid not in student_records:
                student_records[f_sid] = current_record
            else:
                existing = student_records[f_sid]
                # 录取优先 + 同状态取最新
                if existing["status"] == "reject" and current_record["status"] == "accept":
                    student_records[f_sid] = current_record
                elif existing["status"] == current_record["status"]:
                    if current_record["date"] > existing["date"]:
                        student_records[f_sid] = current_record
        else:
            # 无有效学号：直接加入拒绝列表
            no.append({
                "学号": f_sid if f_sid else "未知",
                "姓名": f_name,
                "报名班级": apply_class,
                "原因": current_record["reason"],
                "原主题": subj
            })

        bar.progress((idx+1)/total, text=f"解析中：{idx+1}/{total} 封")

    except Exception as e:
        err_msg = f"解析异常: {str(e)[:20]}"
        no.append({"学号": "?", "姓名": "?", "原因": err_msg})
        bar.progress((idx+1)/total, text=f"解析中：{idx+1}/{total} 封")

# ========== 生成最终录取/拒绝列表（含班级信息） ==========
ok = []
no = []
for sid, record in student_records.items():
    if record["status"] == "accept":
        ok.append({
            "学号": sid,
            "姓名": record["name"],
            "录取班级": record["class"],
            "备注": record.get("remark", ""),
            "报名时间": record["date"].strftime("%Y-%m-%d %H:%M")
        })
    else:
        no.append({
            "学号": sid,
            "姓名": record["name"],
            "报名班级": record["class"],
            "原因": record["reason"],
            "报名时间": record["date"].strftime("%Y-%m-%d %H:%M"),
            "原主题": record["subject"]
        })
                    # 结果展示
                    st.success(f"✅ 录取 {len(ok)} 人")
                    st.dataframe(ok, use_container_width=True)
                    if ok:
                        st.download_button("📥 下载录取名单", eh.to_excel_bytes(ok), "录取.xlsx")
                    else:
                        st.info("暂无录取人员")

                    st.warning(f"❌ 拒绝 {len(no)} 人")
                    st.dataframe(no, use_container_width=True)
                    if no:
                        st.download_button("📥 下载拒绝名单", eh.to_excel_bytes(no), "拒绝.xlsx")
                    else:
                        st.info("暂无拒绝人员")

        except imaplib.IMAP4.error as ex:
            st.error(f"❌ 邮箱登录/操作失败：{str(ex)}（请检查账号、授权码或文件夹名）")
        except ValueError as ex:
            st.error(f"❌ 文件夹名异常：{str(ex)}")
        except Exception as ex:
            st.error(f"❌ 未知错误：{str(ex)}")
