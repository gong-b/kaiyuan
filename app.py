import streamlit as st
import subprocess
import sys
import os
import pandas as pd
from datetime import datetime
from pathlib import Path
import shutil
import locale

# ---------------------- 基础配置 ----------------------
# 设置中文环境（兼容不同系统）
try:
    locale.setlocale(locale.LC_ALL, 'zh_CN.UTF-8')
except:
    try:
        locale.setlocale(locale.LC_ALL, 'Chinese')
    except:
        st.warning("⚠️ 无法设置中文环境，部分显示可能异常")

# 页面配置
st.set_page_config(
    page_title="书法班报名自动筛选系统",
    page_icon="🎓",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# 目录初始化
BASE_DIR = Path(__file__).parent
DATA_DIR = BASE_DIR / "data"
DATA_DIR.mkdir(exist_ok=True, parents=True)

# ---------------------- 页面UI ----------------------
st.title("🎓 书法班报名自动筛选系统")

# 1. 邮箱配置区域
st.subheader("📩 邮箱配置")
col_email, col_pwd = st.columns(2)
with col_email:
    email = st.text_input(
        "浙大IMAP邮箱",
        placeholder="例如：xxx@zju.edu.cn",
        value=os.environ.get("EMAIL_USER", ""),
        help="需使用浙大邮箱，且已开启IMAP/SMTP服务"
    )
with col_pwd:
    pwd = st.text_input(
        "邮箱客户端专用密码",
        type="password",
        help="⚠️ 不是邮箱登录密码！需在浙大邮箱设置中开启IMAP后生成专用密码"
    )

# 2. 时间范围配置
st.subheader("⏰ 筛选时间范围")
col_start, col_end = st.columns(2)
with col_start:
    start_date = st.date_input(
        "开始日期",
        value=datetime(2025, 3, 1),
        format="YYYY-MM-DD",
        help="仅处理该日期及之后收到的邮件"
    )
with col_end:
    end_date = st.date_input(
        "结束日期",
        value=datetime.now(),
        format="YYYY-MM-DD",
        help="仅处理该日期之前收到的邮件"
    )

# 3. 录取配置
st.subheader("🎯 录取配置")
quota = st.number_input(
    "录取总名额",
    min_value=1,
    max_value=100,
    value=25,
    step=1,
    help="包含新鸿基直接录取的名额，非新鸿基学生按邮件时间先到先得"
)

# 4. 基础名单上传
st.subheader("📂 基础名单上传")
col_nhj, col_black, col_last = st.columns(3)
with col_nhj:
    new_hongji_file = st.file_uploader(
        "新鸿基推荐名单 *",
        type=["xlsx", "xls"],
        help="必填！包含学号列的Excel文件，新鸿基学生直接录取"
    )
with col_black:
    blacklist_file = st.file_uploader(
        "黑名单（可选）",
        type=["xlsx", "xls"],
        help="可选！包含需跳过的学号，不会参与筛选"
    )
with col_last:
    last_year_file = st.file_uploader(
        "去年已录取名单 *",
        type=["xlsx", "xls"],
        help="必填！包含去年录取的学号，避免重复录取"
    )

# ---------------------- 核心执行逻辑 ----------------------
if st.button("▶️ 开始筛选", type="primary"):
    # 前置校验：必填项检查
    required_check = {
        "邮箱账号": email.strip() != "",
        "客户端密码": pwd.strip() != "",
        "新鸿基名单": new_hongji_file is not None,
        "去年录取名单": last_year_file is not None
    }
    missing_fields = [k for k, v in required_check.items() if not v]
    if missing_fields:
        st.error(f"❌ 请补全必填项：{', '.join(missing_fields)}")
        st.stop()

    # 保存上传的Excel文件到data目录
    try:
        st.info("📤 正在保存上传的名单文件...")
        
        # 保存新鸿基名单
        new_hongji_path = DATA_DIR / "2024-2025学年秋冬学期新鸿基推荐学生名单.xlsx"
        with open(new_hongji_path, "wb") as f:
            f.write(new_hongji_file.getbuffer())
        
        # 保存去年录取名单
        last_year_path = DATA_DIR / "24秋冬学期开源课堂人员名单.xlsx"
        with open(last_year_path, "wb") as f:
            f.write(last_year_file.getbuffer())
        
        # 保存黑名单（可选）
        if blacklist_file:
            blacklist_path = DATA_DIR / "blacklist.xlsx"
            with open(blacklist_path, "wb") as f:
                f.write(blacklist_file.getbuffer())
        
        st.success("✅ 文件上传完成！开始连接邮箱处理数据...")
    except Exception as e:
        st.error(f"❌ 文件保存失败：{str(e)}")
        st.code(f"详细错误：{str(e)}", language="text")
        st.stop()

    # 设置环境变量（传递给main.py）
    os.environ["IMAP_HOST"] = "imap.zju.edu.cn"
    os.environ["IMAP_PORT"] = "993"
    os.environ["EMAIL_USER"] = email.strip()
    os.environ["EMAIL_PASSWORD"] = pwd.strip()
    os.environ["START_DATE"] = start_date.strftime("%d-%b-%Y")  # 转换为IMAP兼容格式
    os.environ["END_DATE"] = end_date.strftime("%d-%b-%Y")
    os.environ["ADMISSION_QUOTA"] = str(quota)

    # 执行main.py并捕获输出
    with st.spinner("🔍 正在筛选邮件和处理数据（最长等待10分钟）..."):
        try:
            # 切换工作目录到脚本所在目录
            original_cwd = os.getcwd()
            os.chdir(BASE_DIR)
            
            # 执行主脚本（超时10分钟）
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
            st.subheader("📜 运行日志")
            log_content = result.stdout + "\n" + result.stderr
            log_lines = [line.strip() for line in log_content.split("\n") if line.strip()]
            
            # 高亮日志（错误标红，警告标橙）
            highlighted_log = []
            for line in log_lines:
                if any(level in line for level in ["CRITICAL", "ERROR"]):
                    highlighted_log.append(f"<span style='color: #dc2626; font-weight: bold;'>{line}</span>")
                elif "WARNING" in line:
                    highlighted_log.append(f"<span style='color: #f59e0b;'>{line}</span>")
                elif "INFO" in line:
                    highlighted_log.append(f"<span style='color: #059669;'>{line}</span>")
                else:
                    highlighted_log.append(line)
            
            # 折叠展示日志（默认展开错误）
            with st.expander("查看完整日志（错误标红/警告标橙）", expanded=result.returncode != 0):
                st.markdown("<br>".join(highlighted_log), unsafe_allow_html=True)

            # ---------------------- 执行状态判断 ----------------------
            if result.returncode == 0:
                st.success("✅ 筛选完成！")
            else:
                st.error(f"❌ 筛选过程出错（返回码：{result.returncode}）")
                # 提取关键错误信息展示
                error_lines = [line for line in log_lines if "ERROR" in line or "CRITICAL" in line]
                if error_lines:
                    st.subheader("🔴 关键错误信息（前10条）")
                    for idx, line in enumerate(error_lines[:10], 1):
                        st.code(f"{idx}. {line}")

            # ---------------------- 结果展示与下载 ----------------------
            st.subheader("📊 筛选结果")
            admitted_path = DATA_DIR / "admitted_students.xlsx"
            rejected_path = DATA_DIR / "rejected_students.xlsx"

            col_admit, col_reject = st.columns(2)
            # 录取名单
            with col_admit:
                if os.path.exists(admitted_path) and os.path.getsize(admitted_path) > 0:
                    try:
                        df_admitted = pd.read_excel(admitted_path, engine="openpyxl")
                        st.write(f"✅ 录取名单（共{len(df_admitted)}人）")
                        st.dataframe(df_admitted, use_container_width=True, hide_index=True)
                        # 下载按钮
                        with open(admitted_path, "rb") as f:
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

            # 拒绝名单
            with col_reject:
                if os.path.exists(rejected_path) and os.path.getsize(rejected_path) > 0:
                    try:
                        df_rejected = pd.read_excel(rejected_path, engine="openpyxl")
                        st.write(f"❌ 拒绝名单（共{len(df_rejected)}人）")
                        st.dataframe(df_rejected, use_container_width=True, hide_index=True)
                        # 下载按钮
                        with open(rejected_path, "rb") as f:
                            st.download_button(
                                label="📥 下载拒绝名单",
                                data=f,
                                file_name=f"书法班拒绝名单_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                            )
                    except Exception as e:
                        st.warning(f"⚠️ 读取拒绝名单失败：{str(e)}")
                else:
                    st.info("📭 暂无拒绝名单（可能所有学生都符合条件）")

        except subprocess.TimeoutExpired:
            st.error("""❌ 处理超时（超过10分钟）！请检查：
            1. 邮件数量是否过多（建议缩小日期范围测试）
            2. 浙大邮箱IMAP连接是否稳定
            3. 网络是否正常（Streamlit服务器→浙大邮箱）""")
        except PermissionError:
            st.error("❌ 执行权限不足！请检查：\n1. 是否有main.py的执行权限\n2. 是否有data目录的写入权限")
        except Exception as e:
            st.error(f"❌ 执行失败：{str(e)}")
            st.code(f"详细错误堆栈：{str(e)}", language="text")

# ---------------------- 辅助功能 ----------------------
st.divider()
st.subheader("ℹ️ 辅助功能")

# 清理数据按钮
col_clean, col_help = st.columns([1, 3])
with col_clean:
    if st.button("🗑️ 清理所有数据", type="secondary"):
        try:
            # 清理Excel文件
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
        ### 核心规则
        1. **新鸿基学生**：直接录取，占用总名额
        2. **去年已录取**：自动拒绝，不参与筛选
        3. **黑名单学生**：直接跳过，不参与筛选
        4. **普通学生**：需满足
           - 是学生资助对象（DOCX中标记为“是”）
           - 申请理由≥95个中文字符
           - 按邮件接收时间排序，名额有限先到先得

        ### 常见问题
        1. 登录失败：检查IMAP开关是否开启、客户端密码是否正确
        2. 无邮件结果：检查日期范围是否正确、邮件主题是否含“书法班”
        3. Excel读取失败：确保文件含“学号”列，格式为.xlsx（推荐）
        4. DOCX解析失败：确保文件有表格，且包含“资助对象”“申请理由”字段
        """)

# 页脚
st.divider()
st.caption(f"系统更新时间：{datetime.now().strftime('%Y-%m-%d %H:%M:%S')} | 适配浙大邮箱IMAP协议")
