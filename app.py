import streamlit as st
import logging
import tempfile
from pathlib import Path
from email import message_from_bytes
from modules import (
    Config, EmailParser, ExcelHandler, 
    DocxParser, PdfProcessor
)

# ===================== 页面配置 =====================
st.set_page_config(
    page_title="书法班报名审核系统",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# ===================== 日志配置 =====================
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# ===================== 初始化处理器 =====================
email_parser = EmailParser()
excel_handler = ExcelHandler()
docx_parser = DocxParser()
# PDF处理器（本地运行时配置wkhtmltopdf路径，Cloud版自动跳过）
pdf_processor = PdfProcessor(wkhtmltopdf_path=r"D:\program\wkhtmltopdf\bin\wkhtmltopdf.exe")

# ===================== 页面UI =====================
st.title("📝 书法班报名审核系统")
st.divider()

# 第一步：上传基础名单
col1, col2 = st.columns(2)
with col1:
    new_hongji_file = st.file_uploader(
        "📤 新鸿基推荐学生名单（Excel）",
        type=["xlsx"],
        key="new_hongji"
    )
with col2:
    last_year_file = st.file_uploader(
        "📤 去年录取名单（Excel）",
        type=["xlsx"],
        key="last_year"
    )

# 第二步：上传邮件文件
# 日期范围选择
start_date = st.date_input("开始日期", value=pd.to_datetime("2026-03-01"))
end_date = st.date_input("截止日期", value=pd.to_datetime("2026-04-01"))
st.subheader("📧 上传报名邮件")
email_files = st.file_uploader(
    "上传EML格式邮件（可多选）",
    type=["eml"],
    accept_multiple_files=True,
    key="email_files"
)

# 第三步：触发审核
if st.button("🚀 开始审核", type="primary", disabled=not (new_hongji_file and last_year_file and email_files)):
    with st.spinner("正在处理邮件..."):
        # 1. 读取基础名单
        new_hongji_ids = excel_handler.read_student_list(new_hongji_file)
        last_year_ids = excel_handler.read_student_list(last_year_file)
        
        # 2. 初始化结果
        admitted = []  # 录取列表
        rejected = []  # 拒绝列表
        candidates = []  # 候选列表
        
        # 3. 处理每封邮件
        progress_bar = st.progress(0)
        total_emails = len(email_files)
        
        for idx, email_file in enumerate(email_files):
            try:
                # 读取邮件内容
                raw_email = email_file.read()
                msg = message_from_bytes(raw_email)
                
           # 校验邮件日期
           # 校验邮件日期（使用界面选择的日期）
recv_date = parsedate_to_datetime(msg.get("Date"))
recv_date = recv_date.replace(tzinfo=None)  # 去掉时区

if not (start_date <= recv_date.date() <= end_date):
    rejected.append({
        "学号": "未知",
        "姓名": "未知",
        "原主题": subject,
        "原因": f"不在所选日期范围内 {start_date} ~ {end_date}"
    })
    continue
                
                # 解析主题（姓名+学号）
                subject = email_parser.parse_subject(msg)
                name, student_id = email_parser.extract_name_id(subject)
                
                if not name or not student_id:
                    rejected.append({
                        "学号": "未知",
                        "姓名": "未知",
                        "原主题": subject,
                        "原因": "主题格式错误（示例：薛孜324011234书法班报名申请）"
                    })
                    continue
                
                # 新鸿基直接录取
                if student_id in new_hongji_ids:
                    admitted.append({
                        "学号": student_id,
                        "姓名": name,
                        "备注": "新鸿基"
                    })
                    # 本地运行时生成PDF（Cloud版自动跳过）
                    pdf_processor.save_email_pdf(msg, student_id, name, Config.PDF_DIR)
                    continue
                
                # 去年已录取：拒绝
                if student_id in last_year_ids:
                    rejected.append({
                        "学号": student_id,
                        "姓名": name,
                        "原因": "去年已录取"
                    })
                    continue
                
                # 提取DOCX附件
                with tempfile.TemporaryDirectory() as tmpdir:
                    tmp_path = Path(tmpdir)
                    docx_files = email_parser.extract_docx_attachments(msg, tmp_path)
                    
                    # 无DOCX附件：拒绝
                    if not docx_files:
                        rejected.append({
                            "学号": student_id,
                            "姓名": name,
                            "原因": "缺少DOCX申请附件"
                        })
                        continue
                    
                    # 解析DOCX
                    docx_info = docx_parser.parse(str(docx_files[0]))
                    
                    # 非资助对象：拒绝
                    if not docx_info["is_supported"]:
                        rejected.append({
                            "学号": student_id,
                            "姓名": name,
                            "原因": "非学生资助对象"
                        })
                        continue
                    
                    # 申请理由不足：拒绝
                    if docx_info["reason_length"] < Config.MIN_REASON_LENGTH:
                        rejected.append({
                            "学号": student_id,
                            "姓名": name,
                            "原因": f"申请理由不足（仅{docx_info['reason_length']}字，需≥{Config.MIN_REASON_LENGTH}字）"
                        })
                        continue
                    
                    # 符合条件：加入候选
                    candidates.append({
                        "学号": student_id,
                        "姓名": name,
                        "备注": "候选"
                    })
            
            except Exception as e:
                logger.error(f"处理邮件失败: {str(e)}")
                rejected.append({
                    "学号": "未知",
                    "姓名": "未知",
                    "原主题": "解析失败",
                    "原因": f"系统错误：{str(e)}"
                })
            
            # 更新进度条
            progress_bar.progress((idx + 1) / total_emails)
        
        # 4. 处理候选名单（补足录取名额）
        remaining_quota = Config.ADMIT_QUOTA - len(admitted)
        if remaining_quota > 0 and candidates:
            st.info(f"📊 新鸿基录取{len(admitted)}人，剩余名额{remaining_quota}个，从候选中补充")
            # 按学号排序候选
            sorted_candidates = sorted(candidates, key=lambda x: x["学号"])
            # 补充录取
            admitted.extend(sorted_candidates[:remaining_quota])
            # 剩余候选加入拒绝
            for candidate in sorted_candidates[remaining_quota:]:
                rejected.append({
                    "学号": candidate["学号"],
                    "姓名": candidate["姓名"],
                    "原因": "名额已满"
                })
        
        # 5. 展示结果
        st.divider()
        col_admit, col_reject = st.columns(2)
        
        with col_admit:
            st.success(f"✅ 录取名单（共{len(admitted)}人）")
            st.dataframe(admitted, use_container_width=True)
            # 下载录取名单
            csv_data = excel_handler.to_csv_bytes(admitted)
            st.download_button(
                "📥 下载录取名单",
                data=csv_data,
                file_name="录取名单.csv",
                mime="text/csv"
            )
        
        with col_reject:
            st.warning(f"❌ 拒绝名单（共{len(rejected)}人）")
            st.dataframe(rejected, use_container_width=True)
            # 下载拒绝名单
            csv_data = excel_handler.to_csv_bytes(rejected)
            st.download_button(
                "📥 下载拒绝名单",
                data=csv_data,
                file_name="拒绝名单.csv",
                mime="text/csv"
            )
        
        st.success("🎉 审核完成！")

# ===================== 使用说明 =====================
