import streamlit as st
import logging
import tempfile
import imaplib  
from datetime import datetime
from email.utils import parsedate_to_datetime
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

c1, c2 = st.columns(2)
with c1:
    hongji = st.file_uploader("新鸿基名单 Excel", type="xlsx")
with c2:
    last = st.file_uploader("去年录取名单 Excel", type="xlsx")

st.subheader("📧 浙大邮箱")
ca, cb = st.columns(2)
with ca:
    user = st.text_input("邮箱账号")
    pwd = st.text_input("授权码", type="password")
with cb:
    folder = st.text_input("文件夹", value="开源课堂")
    s_date = st.date_input("开始日期", datetime(2026,3,1))
    e_date = st.date_input("截止日期", datetime(2026,5,1))

if st.button("🚀 开始审核", disabled=not (hongji and last and user and pwd)):
    with st.spinner("连接邮箱..."):
        H = eh.read_student_list(hongji)
        L = eh.read_student_list(last)
        ok = []
        no = []

        # ========== 修复：缩小外层try范围，仅包裹邮箱连接/获取逻辑 ==========
        try:
            with SecureIMAPClient(user, pwd, folder) as client:
                # 仅用SINCE过滤起始日期，结束日期通过代码过滤（避免IMAP范围限制）
                mails = list(client.fetch_emails(s_date.strftime("%d-%b-%Y")))
                total = len(mails)
                # ========== 修复：进度条初始化+防除零+文本提示 ==========
                bar = st.progress(0, text="准备解析...")
                if total == 0:
                    st.info("ℹ️ 未找到指定日期范围内的邮件")
                else:
                    # ========== 修复：核心循环缩进（放入with client内） ==========
                    for idx, (uid, msg) in enumerate(mails):
                        try:
                            # ========== 修复：日期过滤（时区处理+边界） ==========
                            # 解析日期并转换为本地时区，避免UTC偏移导致日期错误
                            d_utc = parsedate_to_datetime(msg["Date"])
                            d_local = d_utc.astimezone()  # 转为本地时区
                            if not (s_date <= d_local.date() <= e_date):
                                bar.progress((idx+1)/total, text=f"解析中：{idx+1}/{total} 封（跳过非目标日期邮件）")
                                continue

                            # 解析附件（核心信息来源）
                            with tempfile.TemporaryDirectory() as tmp:
                                docs = ep.extract_docx_attachments(msg, Path(tmp))
                                
                                if not docs:
                                    # 若无附件，提取主题显示，记录为拒绝
                                    subj = ep.parse_subject(msg)
                                    no.append({"学号": "未知", "姓名": "无附件", "原因": "缺失DOCX附件", "原主题": subj})
                                    bar.progress((idx+1)/total, text=f"解析中：{idx+1}/{total} 封（无附件）")
                                    continue

                                # 运行优化后的解析器
                                info = dp.parse(str(docs[0]))
                                f_name = info["name"]
                                f_sid = info["sid"]

                                if not f_sid:
                                    no.append({"学号": "未知", "姓名": f_name, "原因": "附件内无学号"})
                                    bar.progress((idx+1)/total, text=f"解析中：{idx+1}/{total} 封（无学号）")
                                    continue

                                # 黑白名单审核
                                if f_sid in H:
                                    ok.append({"学号": f_sid, "姓名": f_name, "备注": "新鸿基(附件提取)"})
                                elif f_sid in L:
                                    no.append({"学号": f_sid, "姓名": f_name, "原因": "去年已录取"})
                                # 合规性审核
                                elif not info["is_supported"]:
                                    no.append({"学号": f_sid, "姓名": f_name, "原因": "非资助对象"})
                                elif info["reason_length"] < Config.MIN_REASON_LENGTH:
                                    no.append({"学号": f_sid, "姓名": f_name, "原因": f"理由不足({info['reason_length']}字)"})
                                else:
                                    ok.append({"学号": f_sid, "姓名": f_name, "备注": "审核通过"})

                        # ========== 修复：单封邮件解析异常仅影响当前邮件，不中断整体流程 ==========
                        except Exception as e:
                            err_msg = f"解析异常: {str(e)[:20]}"
                            no.append({"学号": "?", "姓名": "?", "原因": err_msg})
                        # ========== 修复：进度条实时更新+明确文本提示 ==========
                        bar.progress((idx+1)/total, text=f"解析中：{idx+1}/{total} 封")

                    # ========== 修复：结果展示逻辑移到循环外，且在try内（仅邮箱连接成功后执行） ==========
                    st.success(f"✅ 录取 {len(ok)} 人")
                    st.dataframe(ok, use_container_width=True)
                    # 空数据保护（可选保留，原需求未提但补充更友好）
                    if ok:
                        st.download_button("下载录取名单", eh.to_excel_bytes(ok), "录取.xlsx")
                    else:
                        st.info("暂无录取人员")

                    st.warning(f"❌ 拒绝 {len(no)} 人")
                    st.dataframe(no, use_container_width=True)
                    if no:
                        st.download_button("下载拒绝名单", eh.to_excel_bytes(no), "拒绝.xlsx")
                    else:
                        st.info("暂无拒绝人员")

        # ========== 修复：异常捕获分类，精准提示 ==========
        except imaplib.IMAP4.error as ex:
            st.error(f"❌ 邮箱登录/操作失败：{str(ex)}（请检查账号、授权码或文件夹名）")
        except ValueError as ex:
            st.error(f"❌ 文件夹名异常：{str(ex)}")
        except Exception as ex:
            st.error(f"❌ 未知错误：{str(ex)}")
