import streamlit as st
import subprocess
import sys
import os
import pandas as pd
from datetime import datetime
from pathlib import Path
import shutil
import locale
from config import (
    DATA_DIR, NEW_HONGJI_FILE, LAST_YEAR_FILE, 
    BLACKLIST_FILE, ADMITTED_FILE, REJECTED_FILE
)

# ---------------------- 基础配置 ----------------------
# 设置中文环境（兼容不同系统）
try:
    locale.setlocale(locale.LC_ALL, 'zh_CN.UTF-8')
except:
    try:
        locale.setlocale(locale.LC_ALL, 'Chinese')
    except:
        st.warning("⚠️ 无法设置中文环境，部分文本可能显示异常")

# 页面样式配置
st.set_page_config(
    page_title="书法班报名自动筛选系统",
    page_icon="🎓",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# ---------------------- 页面UI ----------------------
st.title("🎓 书法班报名自动筛选系统")
st.caption("适配文件：新鸿基名单.xlsx | 副本去年报名名单.xlsx | 黑名单.xlsx")

# 1. 邮箱配置区域
st.subheader("📩 浙大邮箱配置")
col_email, col_pwd = st.columns(2)
with col_email:
    email = st.text_input(
        "IMAP邮箱账号",
        placeholder="例如：xxx@zju.edu.cn",
        value=os.environ.get("EMAIL_USER", ""),
        help="必须使用浙大邮箱，需先开启IMAP/SMTP服务"
    )
with col_pwd:
    pwd = st.text_input(
        "客户端专用密码",
        type="password",
        help="⚠️ 不是邮箱登录密码！需在浙大邮箱→设置→账户→生成客户端专用密码"
    )

# 2. 筛选时间范围
st.subheader("⏰ 报名时间范围")
col_start, col_end = st.columns(2)
with col_start:
    start_date = st.date_input(
        "开始日期",
        value=datetime(2025, 3, 1),
        format="YYYY-MM-DD",
        help="仅处理该日期及之后收到的报名邮件"
    )
with col_end:
    end_date = st.date_input(
        "结束日期",
        value=datetime.now(),
        format="YYYY-MM-DD",
        help="仅处理该日期之前收到的报名邮件"
    )

# 3. 录取名额配置
st.subheader("🎯 录取配置")
quota = st.number_input(
    "总录取名额",
    min_value=1,
    max_value=100,
    value=25,
    step=1,
    help="包含新鸿基推荐学生名额，非新鸿基学生按邮件时间先到先得"
)

# 4. 文件上传区域（明确提示文件名）
st.subheader("📂 上传基础名单（需与文件名一致）")
col_nhj, col_last, col_black = st.columns(3)

# 新鸿基名单上传（对应NEW_HONGJI_FILE）
with col_nhj:
    new_hongji_file = st.file_uploader(
        "1. 新鸿基名单.xlsx",
        type=["xlsx", "xls"],
        help="必填！包含新鸿基推荐学生的学号（支持列名：学号/学生学号/ID等）"
    )

# 去年报名名单上传（对应LAST_YEAR_FILE）
with col_last:
    last_year_file = st.file_uploader(
        "2. 副本去年报名名单.xlsx",
        type=["xlsx", "xls"],
        help="必填！包含去年已录取学生的学号，避免重复录取"
    )

# 黑名单上传（对应BLACKLIST_FILE）
with col_black:
    blacklist_file = st.file_uploader(
        "3. 黑名单.xlsx（可选）",
        type=["xlsx", "xls"],
        help="可选！包含需跳过的学生学号（如违规学生）"
    )

# ---------------------- 核心执行逻辑 ----------------------
if st.button("▶️ 开始筛选", type="primary", use_container_width=True):
    # 1. 前置校验：必填项检查
    required_check = {
        "浙大邮箱账号": email.strip() != "",
        "客户端专用密码": pwd.strip() != "",
        "新鸿基名单": new_hongji_file is not None,
        "去年报名名单": last_year_file is not None
    }
    missing_fields = [k for k, v in required_check.items() if not v]
    if missing_fields:
        st.error(f"❌ 请补全必填项：{', '.join(missing_fields)}")
        st.stop()

    # 2. 保存上传文件（路径与config完全一致）
    try:
        st.info("📤 正在保存上传文件...")
        
        # 保存新鸿基名单
        with open(NEW_HONGJI_FILE, "wb") as f:
            f.write(new_hongji_file.getbuffer())
        st.success(f"✅ 新鸿基名单保存成功：{NEW_HONGJI_FILE.name}")
        
        # 保存去年报名名单
        with open(LAST_YEAR_FILE, "wb") as f:
            f.write(last_year_file.getbuffer())
        st.success(f"✅ 去年报名名单保存成功：{LAST_YEAR_FILE.name}")
        
        # 保存黑名单（可选）
        if blacklist_file:
            with open(BLACKLIST_FILE, "wb") as f:
                f.write(blacklist_file.getbuffer())
            st.success(f"✅ 黑名单保存成功：{BLACKLIST_FILE.name}")

    except Exception as e:
        st.error(f"❌ 文件保存失败：{str(e)}")
        st.code(f"详细错误：{str(e)}", language="text")
        st.stop()

    # 3. 设置环境变量（传递给main.py）
    os.environ["IMAP_HOST"] = "imap.zju.edu.cn"
    os.environ["IMAP_PORT"] = "993"
    os.environ["EMAIL_USER"] = email.strip()
    os.environ["EMAIL_PASSWORD"] = pwd.strip()
    os.environ["START_DATE"] = start_date.strftime("%d-%b-%Y")  # IMAP兼容格式（如01-Mar-2025）
    os.environ["END_DATE"] = end_date.strftime("%d-%b-%Y")
    os.environ["ADMISSION_QUOTA"] = str(quota)

    # 4. 执行main.py并捕获输出
    with st.spinner("🔍 正在筛选邮件和处理数据（最长等待10分钟）..."):
        try:
            # 切换到脚本所在目录（避免路径错误）
            original_cwd = os.getcwd()
            os.chdir(Path(__file__).parent)

            # 执行主脚本（超时10分钟，捕获所有输出）
            result = subprocess.run(
                [sys.executable, "main.py"],
                capture_output=True,
                encoding="utf-8",
                errors="replace",
                timeout=600  # 10分钟超时
            )

            # 恢复原工作目录
            os.chdir(original_cwd)

            # ---------------------- 日志展示（高亮错误） ----------------------
            st.subheader("📜 运行日志（错误标红/警告标橙）")
            log_content = result.stdout + "\n" + result.stderr
            log_lines = [line.strip() for line in log_content.split("\n") if line.strip()]

            # 日志高亮处理
            highlighted_log = []
            for line in log_lines:
                if any(level in line for level in ["CRITICAL", "ERROR"]):
                    highlighted_log.append(f"<span style='color:#dc2626; font-weight:bold;'>{line}</span>")
                elif "WARNING" in line:
                    highlighted_log.append(f"<span style='color:#f59e0b;'>{line}</span>")
                elif "INFO" in line:
                    highlighted_log.append(f"<span style='color:#059669;'>{line}</span>")
                else:
                    highlighted_log.append(line)

            # 折叠展示日志（错误时默认展开）
            with st.expander("查看完整日志", expanded=result.returncode != 0):
                st.markdown("<br>".join(highlighted_log), unsafe_allow_html=True)

            # ---------------------- 执行状态判断 ----------------------
            if result.returncode == 0:
                st.success("✅ 筛选流程执行完成！")
            else:
                st.error(f"❌ 筛选流程出错（返回码：{result.returncode}）")
                # 提取关键错误信息
                error_lines = [line for line in log_lines if "ERROR" in line or "CRITICAL" in line]
                if error_lines:
                    st.subheader("🔴 关键错误信息（前10条）")
                    for idx, line in enumerate(error_lines[:10], 1):
                        st.code(f"{idx}. {line}")

            # ---------------------- 结果展示与下载 ----------------------
            st.subheader("📊 筛选结果")
            col_admit, col_reject = st.columns(2)

            # 展示录取名单
            with col_admit:
                if os.path.exists(ADMITTED_FILE) and os.path.getsize(ADMITTED_FILE) > 0:
                    try:
                        df_admit = pd.read_excel(ADMITTED_FILE, engine="openpyxl")
                        st.write(f"✅ 录取名单（共{len(df_admit)}人）")
                        st.dataframe(df_admit, use_container_width=True, hide_index=True)
                        # 下载按钮（带时间戳，避免覆盖）
                        with open(ADMITTED_FILE, "rb") as f:
                            st.download_button(
                                label="📥 下载录取名单",
                                data=f,
                                file_name=f"书法班录取名单_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                            )
                    except Exception as e:
                        st.warning(f"⚠️ 读取录取名单失败：{str(e)}")
                else:
                    st.info("📭 暂无录取名单（可能无符合条件的学生）")

            # 展示拒绝名单
            with col_reject:
                if os.path.exists(REJECTED_FILE) and os.path.getsize(REJECTED_FILE) > 0:
                    try:
                        df_reject = pd.read_excel(REJECTED_FILE, engine="openpyxl")
                        st.write(f"❌ 拒绝名单（共{len(df_reject)}人）")
                        st.dataframe(df_reject, use_container_width=True, hide_index=True)
                        # 下载按钮
                        with open(REJECTED_FILE, "rb") as f:
                            st.download_button(
                                label="📥 下载拒绝名单",
                                data=f,
                                file_name=f"书法班拒绝名单_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                            )
                    except Exception as e:
                        st.warning(f"⚠️ 读取拒绝名单失败：{str(e)}")
                else:
                    st.info("📭 暂无拒绝名单（可能所有学生均符合条件）")

        except subprocess.TimeoutExpired:
            st.error("""❌ 处理超时（超过10分钟）！请检查：
            1. 邮件数量是否过多（建议缩小日期范围测试）
            2. 浙大邮箱IMAP连接是否稳定
            3. 网络是否正常（Streamlit → 浙大邮箱）""")
        except PermissionError:
            st.error("❌ 执行权限不足！请确保：\n1. main.py有执行权限\n2. data目录有读写权限")
        except Exception as e:
            st.error(f"❌ 执行失败：{str(e)}")
            st.code(f"详细错误：{str(e)}", language="text")

# ---------------------- 辅助功能 ----------------------
st.divider()
st.subheader("ℹ️ 辅助工具")

# 清理数据按钮
col_clean, col_help = st.columns([1, 3])
with col_clean:
    if st.button("🗑️ 清理所有数据", type="secondary"):
        try:
            # 清理Excel文件（名单+结果）
            for file in DATA_DIR.glob("*.xlsx"):
                file.unlink(missing_ok=True)
            # 清理日志文件
            for file in DATA_DIR.glob("*.log"):
                file.unlink(missing_ok=True)
            # 清理附件目录
            shutil.rmtree(DATA_DIR / "attachments", ignore_errors=True)
            st.success("✅ 所有数据已清理完成！")
        except Exception as e:
            st.error(f"❌ 清理失败：{str(e)}")

# 使用说明
with col_help:
    with st.expander("📖 详细使用说明", expanded=False):
        st.markdown("""
        ### 一、核心筛选规则
        1. **新鸿基学生**：直接录取（占用总名额），无需参与候补
        2. **去年已录取**：自动拒绝，不参与本次筛选
        3. **黑名单学生**：直接跳过，不进入筛选流程
        4. **普通学生**：需同时满足：
           - DOCX申请材料中标注“是学生资助对象”
           - 申请理由≥95个中文字符
           - 按邮件接收时间排序，名额满则拒绝

        ### 二、常见问题解决
        1. **邮箱登录失败**：
           - 检查IMAP服务是否开启（浙大邮箱→设置→账户→IMAP/SMTP）
           - 确认使用“客户端专用密码”（非登录密码）
        2. **Excel读取失败**：
           - 检查文件名是否与上传框提示一致（如“新鸿基名单.xlsx”）
           - 确认学号列名在支持列表中（学号/学生学号/ID等）
           - 避免文件有合并单元格、密码保护
        3. **无邮件结果**：
           - 检查日期范围是否包含报名邮件时间
           - 确认邮件主题含“书法班”关键词
        """)

# 页脚信息
st.divider()
st.caption(f"系统更新时间：{datetime.now().strftime('%Y-%m-%d %H:%M:%S')} | 适配文件：新鸿基名单.xlsx | 副本去年报名名单.xlsx | 黑名单.xlsx")
