import streamlit as st
import subprocess
import sys
import os
import pandas as pd
from datetime import datetime
from pathlib import Path
import shutil
import locale

# 设置中文环境
try:
    locale.setlocale(locale.LC_ALL, 'zh_CN.UTF-8')
except:
    pass

# 页面基础配置
st.set_page_config(page_title="书法班筛选", page_icon="🎓", layout="wide")
st.title("🎓 书法班报名自动筛选系统")

# ---------------------- 目录与路径初始化 ----------------------
BASE_DIR = Path(__file__).parent
DATA_DIR = BASE_DIR / "data"
DATA_DIR.mkdir(exist_ok=True, parents=True)

# ---------------------- 前端交互区域 ----------------------
st.subheader("📩 邮箱配置")
col_email, col_pwd = st.columns(2)
with col_email:
    email = st.text_input("浙大IMAP邮箱", placeholder="例如：zzbgs@zju.edu.cn", value=os.environ.get("EMAIL_USER", ""))
with col_pwd:
    pwd = st.text_input("邮箱客户端专用密码", type="password", help="不是邮箱登录密码，需在邮箱设置中开启IMAP并生成专用密码")

st.subheader("⏰ 筛选时间范围")
col1, col2 = st.columns(2)
with col1:
    start_date = st.date_input("开始日期", value=datetime(2025, 3, 1), format="YYYY-MM-DD")
with col2:
    end_date = st.date_input("结束日期", value=datetime.now(), format="YYYY-MM-DD")

st.subheader("🎯 录取配置")
quota = st.number_input("录取总名额", min_value=1, max_value=100, value=25, help="包含新鸿基直接录取名额")

st.subheader("📂 基础名单上传")
col3, col4, col5 = st.columns(3)
with col3:
    new_hongji_file = st.file_uploader("新鸿基推荐名单", type=["xlsx", "xls"], help="包含学号列的Excel文件")
with col4:
    blacklist_file = st.file_uploader("黑名单（可选）", type=["xlsx", "xls"], help="无需处理的学号名单")
with col5:
    last_year_file = st.file_uploader("去年已录取名单", type=["xlsx", "xls"], help="避免重复录取的学号名单")

# ---------------------- 核心处理逻辑 ----------------------
if st.button("▶️ 开始筛选", type="primary"):
    # 1. 基础校验
    required_fields = [email, pwd, start_date, end_date, new_hongji_file, last_year_file]
    if not all(required_fields):
        st.warning("⚠️ 请填写邮箱、密码、开始/结束日期，并上传新鸿基名单、去年录取名单！")
        st.stop()

    # 2. 保存上传的Excel文件到data目录
    try:
        # 新鸿基名单
        new_hongji_path = DATA_DIR / "2024-2025学年秋冬学期新鸿基推荐学生名单.xlsx"
        with open(new_hongji_path, "wb") as f:
            f.write(new_hongji_file.getbuffer())
        
        # 去年录取名单
        last_year_path = DATA_DIR / "24秋冬学期开源课堂人员名单.xlsx"
        with open(last_year_path, "wb") as f:
            f.write(last_year_file.getbuffer())
        
        # 黑名单
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
    os.environ["ADMISSION_QUOTA"] = str(quota)

    # 4. 执行main.py并捕获输出
    with st.spinner("🔍 正在筛选邮件和处理数据..."):
        try:
            # 切换到脚本目录
            original_cwd = os.getcwd()
            os.chdir(BASE_DIR)
            
            # 运行主处理脚本
            result = subprocess.run(
                [sys.executable, "main.py"],
                capture_output=True,
                encoding="utf-8",
                errors="replace",
                timeout=600  # 延长超时到10分钟
            )

            # 恢复工作目录
            os.chdir(original_cwd)

            # 显示运行日志
            st.subheader("📜 运行日志")
            log_content = result.stdout + "\n" + result.stderr
            # 日志折叠显示
            with st.expander("查看完整日志", expanded=False):
                st.code(log_content, language="text")

            # 显示执行状态
            if result.returncode == 0:
                st.success("✅ 筛选完成！")
            else:
                st.error(f"❌ 筛选过程出错（返回码：{result.returncode}）")

            # 5. 展示并提供下载结果文件
            st.subheader("📊 筛选结果")
            admitted_path = DATA_DIR / "admitted_students.xlsx"
            rejected_path = DATA_DIR / "rejected_students.xlsx"

            col6, col7 = st.columns(2)
            # 录取名单
            with col6:
                if os.path.exists(admitted_path) and os.path.getsize(admitted_path) > 0:
                    df_admitted = pd.read_excel(admitted_path, engine="openpyxl")
                    st.write(f"✅ 录取名单（共{len(df_admitted)}人）")
                    st.dataframe(df_admitted, use_container_width=True)
                    # 下载按钮
                    with open(admitted_path, "rb") as f:
                        st.download_button(
                            label="📥 下载录取名单",
                            data=f,
                            file_name=f"书法班录取名单_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                else:
                    st.warning("暂无录取名单")

            # 拒绝名单
            with col7:
                if os.path.exists(rejected_path) and os.path.getsize(rejected_path) > 0:
                    df_rejected = pd.read_excel(rejected_path, engine="openpyxl")
                    st.write(f"❌ 拒绝名单（共{len(df_rejected)}人）")
                    st.dataframe(df_rejected, use_container_width=True)
                    # 下载按钮
                    with open(rejected_path, "rb") as f:
                        st.download_button(
                            label="📥 下载拒绝名单",
                            data=f,
                            file_name=f"书法班拒绝名单_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                else:
                    st.warning("暂无拒绝名单")

        except subprocess.TimeoutExpired:
            st.error("❌ 处理超时（超过10分钟），请检查邮件数量或网络状态！")
        except Exception as e:
            st.error(f"❌ 执行失败：{str(e)}", exc_info=True)

# ---------------------- 辅助说明 ----------------------
st.divider()
st.subheader("ℹ️ 使用说明")
st.markdown("""
### 操作步骤
1. **邮箱配置**：
   - 使用浙大IMAP邮箱，需先在邮箱设置中开启IMAP/SMTP服务
   - 密码为「客户端专用密码」，非邮箱登录密码

2. **时间范围**：
   - 开始日期：建议设置为报名起始日
   - 结束日期：建议设置为报名截止日

3. **名单格式要求**：
   - Excel文件（.xlsx/.xls）
   - 包含「学号」列（列名含“学号”即可，不区分大小写）
   - Sheet名称不限，自动识别

4. **录取规则**：
   - 新鸿基推荐学生直接录取
   - 去年已录取学生自动拒绝
   - 非新鸿基学生需满足：
     - 是学生资助对象
     - 申请理由≥95个中文字符
     - 按邮件接收时间排序，名额有限先到先得

5. **结果说明**：
   - 录取名单：新鸿基直接录取 + 候补录取
   - 拒绝名单：格式错误、去年已录取、非资助对象、理由不足、名额已满等
""")

# ---------------------- 清理临时文件 ----------------------
st.divider()
col_clean, col_blank = st.columns([1, 4])
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
            shutil.rmtree(DATA_DIR / "pdfs", ignore_errors=True)
            st.success("✅ 所有数据已清理！")
        except Exception as e:
            st.error(f"❌ 清理失败：{str(e)}")
