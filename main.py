import logging
import os
import shutil
from email_client import EmailClient
from docx_parser import DocxParser
import pandas as pd
from email.utils import parsedate_to_datetime
import streamlit as st
import zipfile

logging.basicConfig(level=logging.INFO, format="%(message)s")

def main():
    logging.info("✅ 开始筛选")

    # 清理并重建附件目录
    if os.path.exists("attachments"):
        shutil.rmtree("attachments")
    os.makedirs("attachments", exist_ok=True)

    def get_ids(path):
        try:
            df = pd.read_excel(path, dtype=str)
            return set(df.iloc[:, 0].dropna().str.strip())
        except Exception as e:
            print(f"⚠️ 读取名单 {path} 失败: {e}")
            return set()

    xhj_ids = get_ids("新鸿基名单.xlsx")
    black_ids = get_ids("黑名单.xlsx")
    last_ids = get_ids("去年名单.xlsx")

    client = EmailClient()
    mails = client.fetch_mails()
    logging.info(f"📩 共收取邮件：{len(mails)}封")

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

        # 筛选规则
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

        # 一人多报去重
        if not reject_reason:
            if sid in processed_students:
                reject_reason = "重复报名，仅录取最先报名的班级"
            else:
                processed_students.add(sid)

        # 复制附件到统一目录
        if attach_path and os.path.exists(attach_path):
            try:
                # 重命名附件，方便查找
                ext = os.path.splitext(attach_path)[1]
                new_name = f"{sid}_{name}{ext}"
                new_path = os.path.join("attachments", new_name)
                shutil.copy(attach_path, new_path)
                all_attachments.append(new_path)
            except:
                pass

        base_row = [sid, name, grade, reason_count, "是" if is_subsidy else "否", apply_class, real_datetime]

        if reject_reason:
            reject_list.append([*base_row, reject_reason])
        else:
            accept_list.append(base_row)

    # 排序：先班级，班内按时间正序
    accept_list.sort(key=lambda x: (x[5], x[6] if x[6] else ""))
    reject_list.sort(key=lambda x: (x[5], x[6] if x[6] else ""))

    accept_final = [[x[0],x[1],x[2],x[3],x[4],x[5]] for x in accept_list]
    reject_final = [[x[0],x[1],x[2],x[3],x[4],x[5],x[7]] for x in reject_final]

    accept_cols = ["学号", "姓名", "年级", "申请理由字数", "是否资助", "报名班级"]
    reject_cols = ["学号", "姓名", "年级", "申请理由字数", "是否资助", "报名班级", "拒绝原因"]

    files = []

    # 导出Excel
    df_accept = pd.DataFrame(accept_final, columns=accept_cols)
    df_accept.to_excel("录取名单.xlsx", index=False)
    files.append("录取名单.xlsx")

    df_reject = pd.DataFrame(reject_final, columns=reject_cols)
    df_reject.to_excel("拒绝名单.xlsx", index=False)
    files.append("拒绝名单.xlsx")

    # 分班导出
    if not df_accept.empty:
        for cls, group in df_accept.groupby("报名班级"):
            fname = f"录取_{cls}.xlsx"
            group.to_excel(fname, index=False)
            files.append(fname)

    if not df_reject.empty:
        for cls, group in df_reject.groupby("报名班级"):
            fname = f"拒绝_{cls}.xlsx"
            group.to_excel(fname, index=False)
            files.append(fname)

    # ======================
    # 🔥 打包所有附件为 ZIP
    # ======================
    zip_path = "所有报名附件.zip"
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
        for f in all_attachments:
            if os.path.exists(f):
                zf.write(f, arcname=os.path.basename(f))

    # ======================
    # 页面底部展示下载
    # ======================
    st.markdown("---")
    st.subheader("📁 全部结果下载")

    # Excel文件
    for f in files:
        try:
            with open(f, "rb") as fp:
                st.download_button(f"📥 下载 {f}", fp, file_name=f)
        except:
            pass

    # 🔥 所有附件打包
    st.markdown("---")
    st.subheader("📎 所有学生报名附件（打包下载）")
    try:
        with open(zip_path, "rb") as fp:
            st.download_button(
                label="📦 下载 所有报名附件.zip",
                data=fp,
                file_name="所有报名附件.zip"
            )
    except:
        st.warning("暂无附件")

    logging.info(f"🎯 录取：{len(accept_final)} 人")
    logging.info(f"❌ 拒绝：{len(reject_final)}")

if __name__ == "__main__":
    main()
