import streamlit as st
import pandas as pd
import os
from email_client import EmailClient
from docx_parser import DocxParser
from email.utils import parsedate_to_datetime
from io import BytesIO

# 彻底关闭所有警告
import logging
logging.getLogger("streamlit").setLevel(logging.ERROR)
logging.getLogger().setLevel(logging.CRITICAL)
import warnings
warnings.filterwarnings("ignore")

# 页面配置（必须放在最开头）
st.set_page_config(page_title="书法班报名筛选系统", layout="wide")
st.title("🎓 书法班报名自动筛选系统")

# 输入区域
col1, col2 = st.columns(2)
with col1:
    email_account = st.text_input("浙大邮箱", placeholder="zzbgs@zju.edu.cn")
    password = st.text_input("客户端专用密码", type="password")
with col2:
    start_date = st.text_input("开始日期 (格式：YYYY-MM-DD)", value="2025-10-02")
    end_date = st.text_input("结束日期 (格式：YYYY-MM-DD)", value="2025-10-10")

# 上传名单
st.subheader("📂 上传名单")
col3, col4, col5 = st.columns(3)
with col3:
    xhj_file = st.file_uploader("新鸿基名单", type="xlsx")
with col4:
    black_file = st.file_uploader("黑名单", type="xlsx")
with col5:
    last_file = st.file_uploader("去年已参加名单", type="xlsx")

# 开始筛选按钮
if st.button("✅ 开始筛选", type="primary", use_container_width=True):
    # 校验输入
    if not all([email_account, password, start_date, end_date, xhj_file, black_file, last_file]):
        st.error("❌ 请填写完整信息并上传所有名单！")
        st.stop()

    # 内存读取名单，不写本地文件
    def get_ids_from_memory(uploaded_file):
        try:
            df = pd.read_excel(BytesIO(uploaded_file.getvalue()), dtype=str)
            return set(df.iloc[:, 0].dropna().str.strip())
        except Exception as e:
            st.error(f"读取名单失败: {e}")
            return set()

    xhj_ids = get_ids_from_memory(xhj_file)
    black_ids = get_ids_from_memory(black_file)
    last_ids = get_ids_from_memory(last_file)

    # 配置环境变量
    os.environ["EMAIL_USER"] = email_account
    os.environ["EMAIL_PASS"] = password
    os.environ["START_DATE"] = start_date
    os.environ["END_DATE"] = end_date

    # 收取邮件
    with st.spinner("📩 正在收取邮件..."):
        client = EmailClient()
        mails = client.fetch_mails()
    st.success(f"✅ 共收取邮件：{len(mails)}封")

    # 筛选逻辑（完全保留你的业务逻辑）
    accept_list = []
    reject_list = []
    processed = set()

    for mail in mails:
        sid = mail.get("student_id", "")
        name = mail.get("name", "")
        rt = mail.get("receive_time", "")
        attach_io = mail.get("attach_io")

        g, c, is_sub, cnt, err = "未知", "未知", False, 0, ""

        try:
            t = parsedate_to_datetime(rt)
        except:
            t = None

        # 内存解析附件
        if attach_io:
            try:
                p = DocxParser(attach_io)
                g = p.get_grade()
                c = p.get_apply_class()
                is_sub = p.get_subsidy()
                cnt = p.get_reason_count()
            except Exception as e:
                err = "文件解析失败"
        else:
            err = "无Word附件"

        # 筛选规则
        if sid in black_ids:
            err = "黑名单"
        elif sid in last_ids:
            err = "本年已参加"
        elif not err and not is_sub:
            err = "非资助对象"
        elif not err and cnt < 100:
            err = f"字数不足({cnt})"

        # 去重
        if not err:
            if sid in processed:
                err = "重复报名"
            else:
                processed.add(sid)

        row = [sid, name, g, cnt, "是" if is_sub else "否", c, t]
        if err:
            reject_list.append(row + [err])
        else:
            accept_list.append(row)

    # 排序：先班级 → 再时间正序
    accept_list.sort(key=lambda x: (x[5], x[6] if x[6] else ""))
    reject_list.sort(key=lambda x: (x[5], x[6] if x[6] else ""))

    # 生成表格数据
    cols1 = ["学号", "姓名", "年级", "申请理由字数", "是否资助", "报名班级"]
    cols2 = ["学号", "姓名", "年级", "申请理由字数", "是否资助", "报名班级", "拒绝原因"]
    df_a = pd.DataFrame([x[:6] for x in accept_list], columns=cols1)
    df_r = pd.DataFrame([x[:7] for x in reject_list], columns=cols2)

    # 结果展示（核心修复：用 Markdown 表格替代 DataFrame，彻底解决前端报错）
    st.success("✅ 筛选完成！")
    col1, col2 = st.columns(2)
    col1.info(f"🎯 最终录取：{len(df_a)} 人")
    col2.error(f"❌ 最终拒绝：{len(df_r)} 人")

    # 录取名单（纯 Markdown 渲染，0 依赖前端动态组件）
    st.subheader("✅ 录取名单")
    st.markdown(df_a.to_markdown(index=False, numalign="left", stralign="left"), unsafe_allow_html=True)

    # 拒绝名单
    st.subheader("❌ 拒绝名单")
    st.markdown(df_r.to_markdown(index=False, numalign="left", stralign="left"), unsafe_allow_html=True)

    # 内存生成 Excel，不写本地文件
    buf_a = BytesIO()
    buf_r = BytesIO()
    df_a.to_excel(buf_a, index=False)
    df_r.to_excel(buf_r, index=False)
    buf_a.seek(0)
    buf_r.seek(0)

    # 下载按钮
    st.subheader("📥 下载名单")
    col_a, col_b = st.columns(2)
    with col_a:
        st.download_button("📥 下载录取名单.xlsx", buf_a, "录取名单.xlsx", type="primary")
    with col_b:
        st.download_button("📥 下载拒绝名单.xlsx", buf_r, "拒绝名单.xlsx", type="primary")
