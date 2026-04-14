import streamlit as st
import subprocess
import sys
import os
import pandas as pd
from datetime import datetime
from pathlib import Path
import shutil

# 页面基础配置
st.set_page_config(page_title="书法班筛选", page_icon="🎓", layout="wide")
st.title("🎓 书法班报名自动筛选系统")

# ---------------------- 目录与路径初始化 ----------------------
DATA_DIR = Path("data")
DATA_DIR.mkdir(exist_ok=True, parents=True)

# ---------------------- 前端交互区域 ----------------------
st.subheader("📩 邮箱配置")
email = st.text_input("浙大IMAP邮箱", placeholder="例如：zzbgs@zju.edu.cn")
pwd = st.text_input("邮箱客户端专用密码", type="password", help="不是邮箱登录密码，需在邮箱设置中开启IMAP并生成专用密码")

st.subheader("⏰ 筛选时间范围")
col1, col2 = st.columns(2)
with col1:
    start_date = st.date_input("开始日期", value=datetime(2025, 3, 1))
with col2:
    end_date = st.date_input("结束日期", value=datetime.now())

st.subheader("📂 基础名单上传")
col3, col4, col5 = st.columns(3)
with col3:
    new_hongji_file = st.file_uploader("新鸿基推荐名单", type="xlsx", help="包含学号列的Excel文件")
with col4:
    blacklist_file = st.file_uploader("黑名单（可选）", type="xlsx", help="无需处理的学号名单")
with col5:
    last_year_file = st.file_uploader("去年已录取名单", type="xlsx", help="避免重复录取的学号名单")

# ---------------------- 核心处理逻辑 ----------------------
if st.button("▶️ 开始筛选", type="primary"):
    # 1. 基础校验
    required_fields = [email, pwd, start_date, end_date, new_hongji_file, last_year_file]
    if not all(required_fields):
        st.warning("⚠️ 请填写邮箱、密码、开始/结束日期，并上传新鸿基名单、去年录取名单！")
        st.stop()

    # 2. 保存上传的Excel文件到data目录
    try:
        new_hongji_path = DATA_DIR / "2024-2025学年秋冬学期新鸿基推荐学生名单.xlsx"
        with open(new_hongji_path, "wb") as f:
            f.write(new_hongji_file.getbuffer())

        last_year_path = DATA_DIR / "24秋冬学期开源课堂人员名单.xlsx"
        with open(last_year_path, "wb") as f:
            f.write(last_year_file.getbuffer())

        if blacklist_file:
            blacklist_path = DATA_DIR / "blacklist.xlsx"
            with open(blacklist_path, "wb") as f:
                f.write(blacklist_file.getbuffer())
        st.success("✅ 文件上传完成！")
    except Exception as e:
        st.error(f"❌ 文件保存失败：{str(e)}")
        st.stop()

    # 3. 设置环境变量
    os.environ["IMAP_HOST"] = "imap.zju.edu.cn"
    os.environ["IMAP_PORT"] = "993"
    os.environ["EMAIL_USER"] = email
    os.environ["EMAIL_PASSWORD"] = pwd
    os.environ["START_DATE"] = start_date.strftime("%d-%b-%Y")
    os.environ["END_DATE"] = end_date.strftime("%d-%b-%Y")

    # 4. 执行main.py并捕获输出
    with st.spinner("🔍 正在从 [开源课堂] 文件夹筛选邮件和处理数据..."):
        try:
            result = subprocess.run(
                [sys.executable, "main.py"],
                capture_output=True,
                encoding="utf-8",
                errors="replace",
                timeout=300
            )

            st.subheader("📜 运行日志")
            log_content = result.stdout + "\n" + result.stderr
            st.code(log_content, language="text")

            # 5. 展示并提供下载结果文件
            st.subheader("📊 筛选结果")
            admitted_path = DATA_DIR / "admitted_students.xlsx"
            rejected_path = DATA_DIR / "rejected_students.xlsx"
            
            col6, col7 = st.columns(2)
            with col6:
                if os.path.exists(admitted_path):
                    df_admitted = pd.read_excel(admitted_path)
                    st.write(f"✅ 录取名单（共{len(df_admitted)}人）")
                    st.dataframe(df_admitted, use_container_width=True)
                    with open(admitted_path, "rb") as f:
                        st.download_button(
                            label="📥 下载录取名单",
                            data=f,
                            file_name=f"书法班录取名单_{datetime.now().strftime('%Y%m%d')}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                else:
                    st.warning("暂无录取名单")

            with col7:
                if os.path.exists(rejected_path):
                    df_rejected = pd.read_excel(rejected_path)
                    st.write(f"❌ 拒绝名单（共{len(df_rejected)}人）")
                    st.dataframe(df_rejected, use_container_width=True)
                    with open(rejected_path, "rb") as f:
                        st.download_button(
                            label="📥 下载拒绝名单",
                            data=f,
                            file_name=f"书法班拒绝名单_{datetime.now().strftime('%Y%m%d')}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                else:
                    st.warning("暂无拒绝名单")

        except subprocess.TimeoutExpired:
            st.error("❌ 处理超时（超过5分钟），请检查邮件数量或网络状态！")
        except Exception as e:
            st.error(f"❌ 执行失败：{str(e)}")

# ---------------------- 辅助说明 ----------------------
st.divider()
st.subheader("ℹ️ 使用说明")
st.markdown("""
1. **邮箱配置**：需使用浙大IMAP邮箱，密码为邮箱客户端专用密码（非登录密码）。
2. **时间范围**：开始日期建议设置为报名起始日，结束日期为当前日。
3. **名单格式**：上传的Excel文件需包含「学号」列。
4. **提取位置**：程序已默认配置为从邮箱的 **「开源课堂」** 文件夹中提取邮件，请确保邮件已正确归类至此。
5. **结果说明**：
    - 录取名单：包含新鸿基直接录取 + 候补录取（按邮件接收时间排序）。
    - 拒绝名单：包含格式错误、去年已录取、缺少附件、理由不足、名额已满等情况。
""")

# ---------------------- 清理临时文件 ----------------------
if st.button("🗑️ 清理临时文件"):
    try:
        for file in DATA_DIR.glob("*.xlsx"):
            file.unlink()
        for file in DATA_DIR.glob("*.log"):
            file.unlink()
        shutil.rmtree(DATA_DIR / "pdfs", ignore_errors=True)
        st.success("✅ 临时文件清理完成！")
    except Exception as e:
        st.error(f"❌ 清理失败：{str(e)}")
