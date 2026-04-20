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
                    for idx, (uid, msg) in enumerate(mails):
                        try:
                            # 1. 过滤：跳过自己发送的邮件
                            sender_email = parseaddr(msg.get("From", ""))[1]
                            if sender_email == user:
                                bar.progress((idx+1)/total, text=f"解析中：{idx+1}/{total} 封（跳过自己发送的邮件）")
                                continue

                            # 2. 过滤：跳过回复/转发的邮件
                            subj = ep.parse_subject(msg)
                            if any(prefix in subj[:5] for prefix in ["RE:", "FW:", "回复:", "转发:"]):
                                bar.progress((idx+1)/total, text=f"解析中：{idx+1}/{total} 封（跳过回复/转发邮件）")
                                continue

                            # 3. 日期过滤（修复时区问题）
                            d_utc = parsedate_to_datetime(msg["Date"])
                            d_local = d_utc.astimezone()
                            if not (s_date <= d_local.date() <= e_date):
                                bar.progress((idx+1)/total, text=f"解析中：{idx+1}/{total} 封（跳过非目标日期邮件）")
                                continue

                            # 4. 提取原始邮件（兼容会话模式）
                            raw_msg = msg
                            if msg.is_multipart():
                                for part in msg.walk():
                                    if part.get_content_type() == "message/rfc822":
                                        raw_msg = message_from_bytes(part.get_payload(decode=True))
                                        break

                            # 5. 解析附件
                            with tempfile.TemporaryDirectory() as tmp:
                                docs = ep.extract_docx_attachments(raw_msg, Path(tmp))
                                
                                # 初始化学号/姓名（用于附件提取失败时从主题补充）
                                f_name = "未知"
                                f_sid = ""

                                # 5.1 附件提取失败：从主题提取学号/姓名
                                if not docs:
                                    # 匹配主题里的“姓名+学号”格式（如“张三+123456+报名申请”）
                                    name_sid_match = re.search(r"([^+]+)\+(\d{8,10})\+", subj)
                                    if name_sid_match:
                                        f_name = name_sid_match.group(1).strip()
                                        f_sid = name_sid_match.group(2).strip()
                                    # 加入拒绝列表（无附件）
                                    no.append({
                                        "学号": f_sid if f_sid else "未知",
                                        "姓名": f_name,
                                        "原因": "缺失DOCX附件",
                                        "原主题": subj
                                    })
                                    bar.progress((idx+1)/total, text=f"解析中：{idx+1}/{total} 封（无附件）")
                                    continue

                                # 5.2 附件提取成功：解析学号/姓名
                                info = dp.parse(str(docs[0]))
                                f_name = info["name"]
                                f_sid = info["sid"]

                                if not f_sid:
                                    no.append({
                                        "学号": "未知",
                                        "姓名": f_name,
                                        "原因": "附件内无学号"
                                    })
                                    bar.progress((idx+1)/total, text=f"解析中：{idx+1}/{total} 封（无学号）")
                                    continue

                                # ========== 第四步：新增黑名单逻辑（优先级最高） ==========
                                # 先判断黑名单→再判断新鸿基→再判断去年录取→最后合规性
                                if f_sid in B:
                                    no.append({
                                        "学号": f_sid,
                                        "姓名": f_name,
                                        "原因": "黑名单人员"
                                    })
                                elif f_sid in H:  # 新鸿基（可选）
                                    ok.append({
                                        "学号": f_sid,
                                        "姓名": f_name,
                                        "备注": "新鸿基(附件提取)"
                                    })
                                elif f_sid in L:  # 去年录取（可选）
                                    no.append({
                                        "学号": f_sid,
                                        "姓名": f_name,
                                        "原因": "去年已录取"
                                    })
                                elif not info["is_supported"]:
                                    no.append({
                                        "学号": f_sid,
                                        "姓名": f_name,
                                        "原因": "非资助对象"
                                    })
                                elif info["reason_length"] < Config.MIN_REASON_LENGTH:
                                    no.append({
                                        "学号": f_sid,
                                        "姓名": f_name,
                                        "原因": f"理由不足({info['reason_length']}字)"
                                    })
                                else:
                                    ok.append({
                                        "学号": f_sid,
                                        "姓名": f_name,
                                        "备注": "审核通过"
                                    })

                        except Exception as e:
                            err_msg = f"解析异常: {str(e)[:20]}"
                            no.append({"学号": "?", "姓名": "?", "原因": err_msg})
                        # 更新进度条
                        bar.progress((idx+1)/total, text=f"解析中：{idx+1}/{total} 封")

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
