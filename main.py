import logging
import re
import os
from datetime import datetime, timezone
from email.header import decode_header
from email.message import Message
from email.utils import parsedate_to_datetime
from config import *
from email_client import SecureIMAPClient
from email_processor import EmailProcessor
from docx_parser import parse_docx
from excel_handler import read_student_list, save_results

# 日志配置（标准化）
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(name)s - %(levelname)s - %(module)s:%(lineno)d - %(message)s",
    handlers=[
        logging.FileHandler(DATA_DIR / "processing.log", encoding="utf-8"),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

# 正则表达式
SUBJECT_PATTERN = re.compile(
    r"^\s*[()（）\[\]【】\{\}｛｝]*([\u4e00-\u9fa5]{2,4})\s*[+＋-—\s]*(\d{8,12})\s*[+＋-—\s]*书法班报名申请[()（）\[\]【】\{\}｛｝]*\s*$",
    re.IGNORECASE
)

def parse_subject_pattern(subject: str) -> tuple[str, str] | tuple[None, None]:
    """解析主题：捕获正则匹配异常"""
    if not subject:
        return None, None
    try:
        clean_subject = re.sub(r"\s+", "", subject)
        match = SUBJECT_PATTERN.match(clean_subject)
        if match:
            return match.group(1).strip(), match.group(2).strip()
        return None, None
    except Exception as e:
        logger.error(f"主题正则匹配异常（主题：{subject[:50]}）: {e}", exc_info=True)
        return None, None

def main():
    """主流程：分层捕获异常，出错有兜底"""
    # 初始化结果容器（即使中间出错，也能保存已有结果）
    admitted: list[dict[str, str]] = []
    rejected: list[dict[str, str]] = []
    candidates: list[tuple[str, str, datetime]] = []

    try:
        # 1. 初始化处理器（捕获初始化异常）
        try:
            email_processor = EmailProcessor()
        except Exception as e:
            logger.error(f"邮件处理器初始化失败: {e}", exc_info=True)
            raise RuntimeError("邮件处理器初始化失败，请检查依赖配置") from e

        # 2. 读取基础名单（捕获Excel读取异常，单独处理每个文件）
        logger.info("读取基础学生名单...")
        list_config = {
            "新鸿基推荐名单": (NEW_HONGJI_FILE, True),  # 必选
            "去年录取名单": (LAST_YEAR_FILE, True),     # 必选
            "黑名单": (BLACKLIST_FILE, False)           # 可选
        }
        list_data = {}
        
        for list_name, (file_path, required) in list_config.items():
            try:
                data = read_student_list(str(file_path))
                list_data[list_name] = data
                # 必选文件为空则终止
                if required and not data:
                    raise RuntimeError(f"{list_name}为空或文件解析失败: {file_path}")
                logger.info(f"{list_name}读取完成: {len(data)}个有效学号")
            except Exception as e:
                if required:
                    raise RuntimeError(f"读取{list_name}失败（必选文件）: {e}") from e
                else:
                    logger.warning(f"读取{list_name}失败（可选文件，已跳过）: {e}")
                    list_data[list_name] = set()

        new_hongji = list_data["新鸿基推荐名单"]
        last_year = list_data["去年录取名单"]
        blacklist = list_data["黑名单"]

        # 3. 解析日期范围（捕获日期解析异常）
        try:
            start_date_str = os.environ.get("START_DATE", "01-Mar-2025")
            end_date_str = os.environ.get("END_DATE", datetime.now().strftime("%d-%b-%Y"))
            start_date = datetime.strptime(start_date_str, "%d-%b-%Y").replace(tzinfo=timezone.utc)
            end_date = datetime.strptime(end_date_str, "%d-%b-%Y").replace(tzinfo=timezone.utc)
            logger.info(f"处理日期范围: {start_date_str} 至 {end_date_str}")
        except ValueError as e:
            raise RuntimeError(f"日期解析失败（格式需为dd-Mon-yyyy，如01-Mar-2025）: {e}") from e

        # 4. 处理邮件（核心逻辑，单封邮件出错不终止）
        logger.info("开始处理邮件...")
        email_count = 0
        error_count = 0
        
        with SecureIMAPClient() as client:
            for uid, msg in client.fetch_emails():
                email_count += 1
                try:
                    # 解析接收时间
                    recv_date = None
                    date_str = msg.get("Date")
                    if date_str:
                        recv_date = parsedate_to_datetime(date_str)
                        if recv_date.tzinfo is None:
                            recv_date = recv_date.replace(tzinfo=timezone.utc)
                        else:
                            recv_date = recv_date.astimezone(timezone.utc)
                    
                    # 日期过滤
                    if not recv_date or not (start_date <= recv_date <= end_date):
                        logger.debug(f"邮件{uid}时间不在范围内，跳过")
                        continue

                    # 解析主题
                    subject = client._get_msg_subject(msg)
                    name, student_id = parse_subject_pattern(subject)
                    
                    # 主题格式校验
                    if not student_id or not name:
                        rejected.append({
                            "学号": "未知",
                            "姓名": "未知",
                            "原主题": subject[:100],  # 截断过长主题
                            "原因": "主题格式错误（示例：薛孜324011234书法班报名申请）"
                        })
                        continue

                    # 黑名单过滤
                    if student_id in blacklist:
                        rejected.append({"学号": student_id, "姓名": name, "原因": "黑名单用户"})
                        continue

                    # 新鸿基直接录取
                    if student_id in new_hongji:
                        admitted.append({"学号": student_id, "姓名": name, "备注": "新鸿基"})
                        continue

                    # 去年已录取
                    if student_id in last_year:
                        rejected.append({"学号": student_id, "姓名": name, "原因": "去年已录取"})
                        continue

                    # 处理附件（捕获附件处理异常）
                    try:
                        attachments = email_processor.save_attachments(msg, student_id, name)
                        docx_files = [a for a in attachments if a.suffix.lower() == ".docx"]
                    except Exception as e:
                        logger.error(f"邮件{uid}附件处理失败: {e}", exc_info=True)
                        rejected.append({"学号": student_id, "姓名": name, "原因": f"附件处理失败: {str(e)[:50]}"})
                        continue

                    # 附件校验
                    if not docx_files:
                        rejected.append({"学号": student_id, "姓名": name, "原因": "缺少DOCX格式申请附件"})
                        continue

                    # 解析DOCX（捕获DOCX解析异常）
                    try:
                        docx_info = parse_docx(str(docx_files[0]))
                    except Exception as e:
                        logger.error(f"邮件{uid}DOCX解析失败: {e}", exc_info=True)
                        rejected.append({"学号": student_id, "姓名": name, "原因": f"申请材料解析失败: {str(e)[:50]}"})
                        continue

                    # DOCX内容校验
                    if not docx_info["is_supported"]:
                        rejected.append({"学号": student_id, "姓名": name, "原因": "非学生资助对象，不符合申请条件"})
                    elif docx_info["reason_length"] < 95:
                        rejected.append({"学号": student_id, "姓名": name, "原因": f"申请理由字数不足（{docx_info['reason_length']}字，需≥95字）"})
                    else:
                        candidates.append((student_id, name, recv_date))

                except Exception as e:
                    # 单封邮件处理失败，计数+1，继续下一封
                    error_count += 1
                    logger.error(f"处理邮件{uid}失败（已跳过）: {str(e)}", exc_info=True)
                    rejected.append({"学号": "未知", "姓名": "未知", "原因": f"邮件处理异常: {str(e)[:50]}"})
                    continue

        # 5. 处理候补名单（即使候选人为空也不报错）
        logger.info(f"邮件处理完成：总计{email_count}封，错误{error_count}封，有效候选{len(candidates)}人")
        remaining_quota = ADMISSION_QUOTA - len(admitted)
        logger.info(f"新鸿基录取{len(admitted)}人，剩余名额{remaining_quota}")
        
        if remaining_quota > 0 and candidates:
            candidates.sort(key=lambda x: x[2])
            admit_candidates = candidates[:remaining_quota]
            reject_candidates = candidates[remaining_quota:]
            
            for sid, name, _ in admit_candidates:
                admitted.append({"学号": sid, "姓名": name, "备注": "非新鸿基（候补）"})
            for sid, name, _ in reject_candidates:
                rejected.append({"学号": sid, "姓名": name, "原因": "符合条件但名额已满"})
        elif candidates:
            for sid, name, _ in candidates:
                rejected.append({"学号": sid, "姓名": name, "原因": "符合条件但名额已满"})

        # 6. 保存结果（最后一步，即使前面有部分错误也保存已有结果）
        try:
            save_results(admitted, rejected)
            logger.info(f"最终结果：录取{len(admitted)}人，拒绝{len(rejected)}人，结果已保存")
        except Exception as e:
            logger.error(f"保存结果失败: {e}", exc_info=True)
            raise RuntimeError(f"筛选完成但结果保存失败: {str(e)}") from e

    # 捕获主流程致命异常（无法继续执行的错误）
    except RuntimeError as e:
        logger.critical(f"主流程致命错误: {e}", exc_info=True)
        # 兜底：尝试保存已有结果
        if admitted or rejected:
            try:
                save_results(admitted, rejected)
                logger.warning(f"已保存部分结果（录取{len(admitted)}人，拒绝{len(rejected)}人）")
            except:
                pass
        raise  # 重新抛出，让前端捕获返回码
    except Exception as e:
        logger.critical(f"未预期的全局异常: {e}", exc_info=True)
        # 兜底保存
        if admitted or rejected:
            try:
                save_results(admitted, rejected)
            except:
                pass
        raise RuntimeError(f"程序执行异常: {str(e)}") from e

if __name__ == "__main__":
    main()
