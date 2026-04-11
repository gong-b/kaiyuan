import logging
import re
import streamlit as st
from datetime import datetime, timezone
from email.header import decode_header
from email.message import Message
from email.utils import parsedate_to_datetime
from config import *
from email_client import SecureIMAPClient
from email_processor import EmailProcessor
from docx_parser import parse_docx
from excel_handler import read_student_list, save_results

# ==================== 日志配置 ====================
class SafeLogFilter(logging.Filter):
    def filter(self, record: logging.LogRecord):
        try:
            record.msg = str(record.msg).encode('utf-8', errors='replace').decode('utf-8')
        except:
            pass
        return True

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s",
    handlers=[
        logging.FileHandler(DATA_DIR / "processing.log", encoding="utf-8"),
        logging.StreamHandler()
    ]
)

logger = logging.getLogger(__name__)
logger.addFilter(SafeLogFilter())

# ==================== 正则配置 ====================
SUBJECT_PATTERN = re.compile(
    r"^\s*"
    r"([()（）\[\]【】\{\}｛｝])?"
    r"([\u4e00-\u9fa5]{2,})"  # 姓名（至少2个中文字符）
    r"\s*[+＋-]?\s*"
    r"(\d{8,12})"  # 学号（8-12位数字）
    r"\s*[+＋-]?\s*"
    r"书法班报名申请"
    r"([)）\]\】\}\｝])?"
    r"\s*$",
    re.UNICODE
)

# ==================== 核心函数 ====================
def parse_subject_pattern(subject: str) -> tuple[str, str] | tuple[None, None]:
    """解析邮件主题，提取姓名和学号"""
    if not subject:
        return None, None
    clean_subject = re.sub(r"\s+", "", subject)
    match = SUBJECT_PATTERN.search(clean_subject)
    if match:
        return match.group(2).strip(), match.group(3).strip()
    return None, None

def parse_subject(msg: Message) -> str:
    """安全解码邮件主题"""
    decoded_parts = []
    try:
        for part, charset in decode_header(msg.get("Subject", "")):
            if isinstance(part, bytes):
                # 优先尝试常见编码
                for encoding in [charset, 'utf-8', 'gb18030', 'big5']:
                    if not encoding:
                        continue
                    try:
                        decoded_parts.append(part.decode(encoding))
                        break
                    except:
                        continue
                else:
                    decoded_parts.append(part.decode('utf-8', errors='replace'))
            else:
                decoded_parts.append(str(part))
    except Exception as e:
        logger.warning(f"主题解码异常: {str(e)}")
        decoded_parts.append("[解码失败]")
    return "".join(decoded_parts)

def process_emails() -> tuple[list[dict], list[dict]]:
    """处理邮件主逻辑"""
    email_processor = EmailProcessor()
    new_hongji = read_student_list(str(NEW_HONGJI_FILE))
    last_year = read_student_list(str(LAST_YEAR_FILE))
    
    admitted = []
    rejected = []
    candidates = []
    
    try:
        with SecureIMAPClient() as client:
            # Streamlit 进度条
            st.info("开始读取邮件...")
            email_count = 0
            progress_bar = st.progress(0)
            
            for uid, msg in client.fetch_emails():
                email_count += 1
                progress_bar.progress(email_count / 100)  # 简易进度
                
                # 解析邮件时间
                try:
                    date_str = msg.get("Date")
                    recv_date = parsedate_to_datetime(date_str) if date_str else None
                    if recv_date:
                        recv_date = recv_date.astimezone(timezone.utc)
                except Exception as e:
                    logger.error(f"日期解析失败: {e}")
                    rejected.append({
                        "学号": "未知",
                        "姓名": "未知",
                        "原主题": parse_subject(msg),
                        "原因": "邮件日期解析失败"
                    })
                    continue
                
                # 时间过滤（2025-03-01之后）
                cutoff_date = datetime(2025, 3, 1, tzinfo=timezone.utc)
                if recv_date and recv_date < cutoff_date:
                    logger.warning(f"邮件{uid}时间不符合要求: {recv_date}")
                    rejected.append({
                        "学号": "未知",
                        "姓名": "未知",
                        "原主题": parse_subject(msg),
                        "原因": f"邮件接收时间过早（{recv_date}）"
                    })
                    continue
                
                # 解析主题
                subject = parse_subject(msg)
                name, student_id = parse_subject_pattern(subject)
                
                if not student_id or not name:
                    rejected.append({
                        "学号": "未知",
                        "姓名": "未知",
                        "原主题": subject,
                        "原因": "主题格式错误（示例：薛孜324011234书法班报名申请）"
                    })
                    continue
                
                # 新鸿基直接录取
                if student_id in new_hongji:
                    admitted.append({"学号": student_id, "姓名": name, "备注": "新鸿基"})
                    email_processor.save_email_pdf(msg, student_id, name)
                    continue
                
                # 去年已录取
                if student_id in last_year:
                    rejected.append({
                        "学号": student_id,
                        "姓名": name,
                        "原因": "去年已录取"
                    })
                    continue
                
                # 保存附件
                attachments = email_processor.save_attachments(msg, student_id, name)
                docx_files = [a for a in attachments if a.suffix == ".docx"]
                
                if not docx_files:
                    rejected.append({
                        "学号": student_id,
                        "姓名": name,
                        "原因": "缺少DOCX格式申请附件"
                    })
                    continue
                
                # 解析DOCX
                try:
                    docx_info = parse_docx(str(docx_files[0]))
                    if not docx_info["is_supported"]:
                        rejected.append({
                            "学号": student_id,
                            "姓名": name,
                            "原因": "非学生资助对象"
                        })
                    elif docx_info["reason_length"] < 95:
                        rejected.append({
                            "学号": student_id,
                            "姓名": name,
                            "原因": f"申请理由字数不足（仅{docx_info['reason_length']}字，需≥95字）"
                        })
                    else:
                        candidates.append((student_id, name, recv_date))
                except Exception as e:
                    rejected.append({
                        "学号": student_id,
                        "姓名": name,
                        "原因": f"附件解析失败: {str(e)[:50]}..."
                    })
        
        # 处理候补名单
        remaining = 25 - len(admitted)
        if remaining > 0 and candidates:
            # 按接收时间排序（早到早得）
            candidates.sort(key=lambda x: x[2])
            # 录取候补
            for student_id, name, _ in candidates[:remaining]:
                admitted.append({"学号": student_id, "姓名": name, "备注": "非新鸿基"})
            # 拒绝剩余候补
            for student_id, name, _ in candidates[remaining:]:
                rejected.append({
                    "学号": student_id,
                    "姓名": name,
                    "原因": "名额已满（候补未录取）"
                })
        
        # 保存结果
        save_results(admitted, rejected)
        logger.info(f"处理完成 - 录取{len(admitted)}人，拒绝{len(rejected)}人")
        
        return admitted, rejected
        
    except Exception as e:
        logger.error(f"处理过程异常: {str(e)}")
        st.error(f"处理失败: {str(e)}")
        return admitted, rejected

# ==================== Streamlit 页面 ====================
def main():
    """Streamlit主页面"""
    # 页面配置
    st.set_page_config(
        page_title=STREAMLIT_PAGE_TITLE,
        page_icon=STREAMLIT_PAGE_ICON,
        layout="wide"
    )
    
    # 侧边栏
    with st.sidebar:
        st.title("📝 书法班报名审核")
        st.divider()
        # 邮箱配置（Streamlit Secrets）
        st.text_input("IMAP服务器", value=IMAP_HOST, key="imap_host")
        st.number_input("IMAP端口", value=IMAP_PORT, key="imap_port")
        st.text_input("邮箱账号", value=EMAIL_USER, key="email_user")
        st.text_input("邮箱密码", type="password", value=EMAIL_PASSWORD, key="email_password")
        st.divider()
        run_button = st.button("开始审核", type="primary")
    
    # 主内容区
    st.header("书法班报名审核系统")
    st.divider()
    
    if run_button:
        # 更新配置
        global IMAP_HOST, IMAP_PORT, EMAIL_USER, EMAIL_PASSWORD
        IMAP_HOST = st.session_state.imap_host
        IMAP_PORT = st.session_state.imap_port
        EMAIL_USER = st.session_state.email_user
        EMAIL_PASSWORD = st.session_state.email_password
        
        # 执行审核
        with st.spinner("正在处理邮件和附件..."):
            admitted, rejected = process_emails()
        
        # 显示结果
        col1, col2 = st.columns(2)
        
        with col1:
            st.subheader("✅ 录取名单")
            if admitted:
                st.dataframe(admitted, use_container_width=True)
                # 下载按钮
                df_admitted = pd.DataFrame(admitted)
                csv_admitted = df_admitted.to_csv(index=False, encoding="utf-8-sig")
                st.download_button(
                    "下载录取名单",
                    csv_admitted,
                    "admitted.csv",
                    "text/csv",
                    key="download_admitted"
                )
            else:
                st.info("暂无录取人员")
        
        with col2:
            st.subheader("❌ 拒绝名单")
            if rejected:
                st.dataframe(rejected, use_container_width=True)
                # 下载按钮
                df_rejected = pd.DataFrame(rejected)
                csv_rejected = df_rejected.to_csv(index=False, encoding="utf-8-sig")
                st.download_button(
                    "下载拒绝名单",
                    csv_rejected,
                    "rejected.csv",
                    "text/csv",
                    key="download_rejected"
                )
            else:
                st.info("暂无拒绝人员")
        
        # 统计信息
        st.divider()
        st.info(f"本次审核完成：共录取 {len(admitted)} 人，拒绝 {len(rejected)} 人")
    
    else:
        st.info("请在侧边栏配置邮箱信息，然后点击【开始审核】按钮")

if __name__ == "__main__":
    main()
