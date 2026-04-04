import logging
import os
import shutil
import zipfile
import pandas as pd
from email.utils import parsedate_to_datetime
import streamlit as st
from email_client import EmailClient
from docx_parser import DocxParser

# 关闭所有警告日志
logging.basicConfig(level=logging.ERROR)
import warnings
warnings.filterwarnings("ignore")

def main():
    st.set_page_config(page_title="书法班报名自动筛选系统", page_icon="🎓", layout="wide")
    st.title("🎓 书法班报名自动筛选系统")

    # 邮箱登录
    st.subheader("📩 邮箱登录")
    email = st.text_input("浙大邮箱", value="zzbgs@zju.edu.cn")
    pwd = st.text_input("客户端专用密码", type="password")

    # 筛选日期
    st.subheader("⏰ 筛选日期")
    d1 = st.date_input("开始日期", value=pd.to_datetime("2025-10-02"))
    d2 = st.date_input("结束日期", value=pd.to_datetime("2025-10-10"))

    # 上传名单
    st.subheader("📂 上传名单")
    f1 = st.file_uploader("新鸿基名单", type="xlsx")
    f2 = st.file_uploader("黑名单", type="xlsx")
    f3 = st.file_uploader("去年已参加", type="xlsx")

    if not st.button("✅ 开始筛选"):
        st.stop()

    # 校验输入
    if not all([email, pwd, f1, f2, f3]):
        st.warning("请填写完整信息并上传所有名单！")
        st.stop()

    # 保存上传的名单
    with open("新鸿基名单.xlsx", "wb") as f:
        f.write(f1.getbuffer())
    with open("黑名单.xlsx", "wb") as f:
        f.write(f2.getbuffer())
    with open("副本去年报名名单.xlsx", "wb") as f:
        f.write(f3.getbuffer())

    # 清理并重建附件目录
    if os.path.exists("data"):
        shutil.rmtree("data")
    if os.path.exists("attachments"):
        shutil.rmtree("attachments")
    os.makedirs("data", exist_ok=True)
    os.makedirs("attachments", exist_ok=True)

    # 读取名单
    def get_ids(path):
        try:
            df = pd.read_excel(path, dtype=str)
            return set(df.iloc[:, 0].dropna().astype(str).str.strip())
        except Exception as e:
            st.error(f"读取{path}失败: {e}")
            return set()

    xhj_ids = get_ids("新鸿基名单.xlsx")
    black_ids = get_ids("黑名单.xlsx")
    last_ids = get_ids("副本去年报名名单.xlsx")

    # 配置邮箱环境变量
    os.environ["EMAIL_USER"] = email
    os.environ["EMAIL_PASS"] = pwd
    os.environ["START_DATE"] = str(d1)
    os.environ["END_DATE"] = str(d2)

    # 收取邮件
    with st.spinner("📩 正在收取邮件..."):
        client = EmailClient()
        mails = client.fetch_mails()
    st.success(f"✅ 共收取邮件：{len(mails)}封")

    accept_list = []
    reject_list = []
    processed_students = set()
    all_attachments = []

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

        # 解析邮件时间
        try:
            real_datetime = parsedate_to_datetime(receive_time)
        except:
            real_datetime = None

        # 解析附件
        if attach_path and os.path.exists(attach_path):
            try:
                parser = DocxParser(attach_path)
                grade = parser.get_grade()
                apply_class = parser.get_apply_class()
                # 统一班级名（补全"班"字）
                if apply_class != "未知班级" and not apply_class.endswith("班"):
                    apply_class = f"{apply_class}班"
                # 统一年级格式（2023→2023级，23→23级，大一→大一）
                if grade.isdigit() and len(grade) == 4:
                    grade = f"{grade}级"
                elif grade.isdigit() and len(grade) == 2:
                    grade = f"{grade}级"
                is_subsidy = parser.get_subsidy()
                reason_count = parser.get_reason_count()
                parse_success = True
            except Exception as e:
                reject_reason = "文件解析失败"

        # 审核规则
        if sid in black_ids:
            reject_reason = "黑名单"
        elif sid in last_ids:
            reject_reason = "本年已参加"
        elif sid in xhj_ids:
            reject_reason = ""
        elif not attach_path or not os.path.exists(attach_path):
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

        # 一人多报：只录取最先报名的班级
        if not reject_reason:
            if sid in processed_students:
                reject_reason = "重复报名，仅录取最先报名的班级"
            else:
                processed_students.add(sid)

        # 复制附件到统一目录，避免重名
        if attach_path and os.path.exists(attach_path):
            try:
                ext = os.path.splitext(attach_path)[1]
                base_name = os.path.basename(attach_path)
                # 用时间戳避免重名
                new_name = f"{sid}_{name}_{int(real_datetime.timestamp()) if real_datetime else 0}{ext}"
                new_path = os.path.join("attachments", new_name)
                shutil.copy(attach_path, new_path)
                all_attachments.append(new_path)
            except Exception as e:
                st.warning(f"复制附件失败: {e}")

        # 组装数据
        base_row = [sid, name, grade, reason_count, "是" if is_subsidy else "否", apply_class, real_datetime]
        if reject_reason:
            reject_list.append([*base_row, reject_reason])
        else:
            accept_list.append(base_row)

    # 排序：先按班级分组，班内按报名时间正序
    accept_list.sort(key=lambda x: (x[5], x[6] if x[6] else ""))
    reject_list.sort(key=lambda x: (x[5], x[6] if x[6] else ""))

    # 去掉时间字段，生成最终导出数据
    accept_final = [[x[0],x[1],x[2],x[3],x[4],x[5]] for x in accept_list]
    reject_final = [[x[0],x[1],x[2],x[3],x[4],x[5],x[7]] for x in reject_list]

    # 表头
    accept_cols = ["学号", "姓名", "年级", "申请理由字数", "是否资助", "报名班级"]
    reject_cols = ["学号", "姓名", "年级", "申请理由字数", "是否资助", "报名班级", "拒绝原因"]

    # 生成DataFrame
    df_accept = pd.DataFrame(accept_final, columns=accept_cols)
    df_reject = pd.DataFrame(reject_final, columns=reject_cols)

    # 导出总表
    df_accept.to_excel("录取名单.xlsx", index=False)
    df_reject.to_excel("拒绝名单.xlsx", index=False)

    # 按班级分班导出
    if not df_accept.empty:
        for cls_name, group in df_accept.groupby("报名班级"):
            group.to_excel(f"录取_{cls_name}.xlsx", index=False)
    if not df_reject.empty:
        for cls_name, group in df_reject.groupby("报名班级"):
            group.to_excel(f"拒绝_{cls_name}.xlsx", index=False)

    # 打包所有附件为ZIP（去重）
    zip_path = "所有报名附件.zip"
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
        seen_names = set()
        for f in all_attachments:
            if os.path.exists(f):
                file_name = os.path.basename(f)
                if file_name not in seen_names:
                    zf.write(f, arcname=file_name)
                    seen_names.add(file_name)

    # 页面展示结果
    st.subheader("🎯 筛选结果")
    st.success(f"✅ 录取：{len(accept_final)} 人 | ❌ 拒绝：{len(reject_final)} 人")

    # 展示录取名单
    st.subheader("✅ 录取名单")
    st.dataframe(df_accept, use_container_width=True)
    with open("录取名单.xlsx", "rb") as f:
        st.download_button("📥 下载录取名单.xlsx", f, file_name="录取名单.xlsx")

    # 展示拒绝名单
    st.subheader("❌ 拒绝名单")
    st.dataframe(df_reject, use_container_width=True)
    with open("拒绝名单.xlsx", "rb") as f:
        st.download_button("📥 下载拒绝名单.xlsx", f, file_name="拒绝名单.xlsx")

    # 分班文件下载区
    st.subheader("📁 分班名单下载")
    for f in os.listdir("."):
        if f.startswith("录取_") and f.endswith(".xlsx"):
            with open(f, "rb") as fp:
                st.download_button(f"📥 下载{f}", fp, file_name=f)
        elif f.startswith("拒绝_") and f.endswith(".xlsx"):
            with open(f, "rb") as fp:
                st.download_button(f"📥 下载{f}", fp, file_name=f)

    # 🔥 附件打包下载区（核心修复！）
    st.markdown("---")
    st.subheader("📎 所有学生报名附件（打包下载）")
    if os.path.exists(zip_path) and os.path.getsize(zip_path) > 0:
        with open(zip_path, "rb") as f:
            st.download_button(
                label="📦 下载 所有报名附件.zip",
                data=f,
                file_name="所有报名附件.zip",
                mime="application/zip"
            )
        st.success(f"✅ 共打包 {len(all_attachments)} 个附件")
    else:
        st.warning("⚠️ 未找到有效附件，可能所有邮件都无附件")

if __name__ == "__main__":
    main()
