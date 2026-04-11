import streamlit as st
import logging
import os
import pandas as pd
from email_client import EmailClient
from docx_parser import DocxParser
from email.utils import parsedate_to_datetime
from io import BytesIO

# ============= 彻底关闭所有警告，解决刷屏问题 =============
logging.getLogger("streamlit").setLevel(logging.ERROR)
logging.getLogger().setLevel(logging.CRITICAL)
import warnings
warnings.filterwarnings("ignore")
# ============================================================

# 页面标题
st.title("🎓 书法班报名自动筛选系统")

# 输入区域
email = st.text_input("浙大邮箱")
pwd = st.text_input("客户端专用密码", type="password")
start_date = st.text_input("开始日期 (例：2025-10-02)")
end_date = st.text_input("结束日期 (例：2025-10-10)")

st.subheader("📂 上传名单")
xhj_file = st.file_uploader("新鸿基名单", type="xlsx")
black_file = st.file_uploader("黑名单", type="xlsx")
last_file = st.file_uploader("去年已参加", type="xlsx")

# 开始筛选
if st.button("✅ 开始筛选"):
    if not all([email, pwd, start_date, end_date, xhj_file, black_file, last_file]):
        st.warning("请填写完整信息！")
        st.stop()

    # 内存读取文件，彻底解决云端写盘问题
    def get_ids(uploaded_file):
        try:
            df = pd.read_excel(BytesIO(uploaded_file.getvalue()), dtype=str)
            return set(df.iloc[:, 0].dropna().str.strip())
        except Exception as e:
            st.error(f"读取名单失败: {e}")
            return set()

    xhj_ids = get_ids(xhj_file)
    black_ids = get_ids(black_file)
    last_ids = get_ids(last_file)

    # 配置环境变量
    os.environ["EMAIL_USER"] = email
    os.environ["EMAIL_PASS"] = pwd
    os.environ["START_DATE"] = start_date
    os.environ["END_DATE"] = end_date

    # 收取邮件
    client = EmailClient()
    mails = client.fetch_mails()
    st.success(f"📩 共收取邮件：{len(mails)}封")

    # ===================== 你的原始逻辑，100% 不变 =====================
    accept_list = []
    reject_list = []
    processed_students = set()

    for mail in mails:
        sid = mail.get("student_id", "")
        name = mail.get("name", "")
        receive_time = mail.get("receive_time", "")
        attach_path = mail.get("attachment_path", "")

        grade = "未知"
        apply_class = "未知班级"
        is_subsidy = False
        reason_count = 0
        reject_reason = ""
        parse_success = False
        real_datetime = None

        try:
            real_datetime = parsedate_to_datetime(receive_time)
        except:
            real_datetime = None

        if attach_path:
            try:
                parser = DocxParser(attach_path)
                grade = parser.get_grade()
                apply_class = parser.get_apply_class()
                is_subsidy = parser.get_subsidy()
                reason_count = parser.get_reason_count()
                parse_success = True
            except Exception as e:
                reject_reason = "文件解析失败"

        if sid in black_ids:
            reject_reason = "黑名单"
        elif sid in last_ids:
            reject_reason = "本年已参加"
        elif sid in xhj_ids:
            reject_reason = ""
        elif not attach_path:
            reject_reason = "无Word附件"
        elif not parse_success:
            reject_reason = "文件解析失败"
        else:
            if not is_subsidy:
                reject_reason = "非资助对象"
            elif reason_count < 100:
                reject_reason = f"理由字数不足({reason_count}/100)"
            else:
                reject_reason = ""

        if not reject_reason:
            if sid in processed_students:
                reject_reason = "重复报名，仅录取最先报名的班级"
            else:
                processed_students.add(sid)

        base_row = [sid, name, grade, reason_count, "是" if is_subsidy else "否", apply_class, real_datetime]

        if reject_reason:
            reject_list.append([*base_row, reject_reason])
        else:
            accept_list.append(base_row)

    # 排序：先班级 → 再时间正序（完全不变）
    accept_list.sort(key=lambda x: (x[5], x[6] if x[6] else ""))
    reject_list.sort(key=lambda x: (x[5], x[6] if x[6] else ""))

    accept_final = [[x[0], x[1], x[2], x[3], x[4], x[5]] for x in accept_list]
    reject_final = [[x[0], x[1], x[2], x[3], x[4], x[5], x[7]] for x in reject_list]

    accept_cols = ["学号", "姓名", "年级", "申请理由字数", "是否资助", "报名班级"]
    reject_cols = ["学号", "姓名", "年级", "申请理由字数", "是否资助", "报名班级", "拒绝原因"]

    df_accept = pd.DataFrame(accept_final, columns=accept_cols)
    df_reject = pd.DataFrame(reject_final, columns=reject_cols)

    df_accept.to_excel("录取名单.xlsx", index=False)
    df_reject.to_excel("拒绝名单.xlsx", index=False)
    # =================================================================

    # ============= 核心修复：用 markdown 表格替代 DataFrame，彻底解决前端报错 =============
    st.success("✅ 筛选完成！")
    col1, col2 = st.columns(2)
    col1.info(f"🎯 最终录取：{len(accept_final)} 人")
    col2.error(f"❌ 最终拒绝：{len(reject_final)} 人")

    # 录取名单（纯 markdown 渲染，0 依赖前端组件）
    st.subheader("✅ 录取名单")
    st.markdown(df_accept.to_markdown(index=False, numalign="left", stralign="left"), unsafe_allow_html=True)

    # 拒绝名单
    st.subheader("❌ 拒绝名单")
    st.markdown(df_reject.to_markdown(index=False, numalign="left", stralign="left"), unsafe_allow_html=True)

    # 下载按钮
    st.subheader("📥 下载名单")
    with open("录取名单.xlsx", "rb") as f:
        st.download_button("📥 下载录取名单.xlsx", f, "录取名单.xlsx")
    with open("拒绝名单.xlsx", "rb") as f:
        st.download_button("📥 下载拒绝名单.xlsx", f, "拒绝名单.xlsx")
