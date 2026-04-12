import logging
import re
import os
from datetime import datetime, timezone
from email.header import decode_header
from email.message import Message
from email.utils import parsedate_to_datetime
from pathlib import Path
from typing import Generator, Tuple, Optional, List, Dict, Any

# ---------------------- 导入配置与工具函数 ----------------------
from config import (
    IMAP_HOST, IMAP_PORT, EMAIL_USER, EMAIL_PASSWORD,
    NEW_HONGJI_FILE, LAST_YEAR_FILE, BLACKLIST_FILE,
    ADMISSION_QUOTA, MIN_REASON_LENGTH, DATA_DIR
)
from excel_handler import read_student_list, save_results

# ---------------------- 日志配置 ----------------------
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(name)s - %(levelname)s - %(module)s:%(lineno)d - %(message)s",
    handlers=[
        logging.FileHandler(DATA_DIR / "processing.log", encoding="utf-8"),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

# ---------------------- 正则表达式（适配主题格式） ----------------------
# 支持的主题格式：姓名+学号（分隔符：空格/+/—/-）+书法班报名申请
SUBJECT_PATTERN = re.compile(
    r"^\s*[()（）\[\]【】\{\}｛｝]*"  # 可选括号
    r"([\u4e00-\u9fa5]{2,4})"        # 姓名（2-4个中文）
    r"\s*[\+＋\-—\s]*"              # 分隔符（空格/+/—/-）
    r"(\d{8,12}|-?\d{8,12})"        # 学号（8-12位数字，支持带短横线）
    r"\s*[\+＋\-—\s]*"              # 分隔符
    r"书法班报名申请"                # 固定后缀
    r"[()（）\[\]【】\{\}｛｝]*\s*$",  # 可选括号
    re.IGNORECASE
)

# ---------------------- IMAP邮箱客户端类 ----------------------
import imaplib
import ssl

class SecureIMAPClient:
    def __init__(self):
        self.host = IMAP_HOST
        self.port = IMAP_PORT
        self.user = EMAIL_USER
        self.password = EMAIL_PASSWORD
        self.mailbox = "INBOX"
        self.conn: Optional[imaplib.IMAP4_SSL] = None

    def __enter__(self) -> "SecureIMAPClient":
        # 校验配置完整性
        if not all([self.user, self.password, self.host]):
            raise ValueError(
                "IMAP配置不完整：\n"
                f"- 邮箱：{'已填' if self.user else '未填'}\n"
                f"- 密码：{'已填' if self.password else '未填'}\n"
                f"- 主机：{self.host}"
            )

        # 初始化SSL上下文
        context = ssl.create_default_context()
        context.check_hostname = False  # 兼容浙大邮箱证书
        context.verify_mode = ssl.CERT_NONE

        try:
            # 建立SSL连接
            self.conn = imaplib.IMAP4_SSL(
                self.host, self.port, ssl_context=context, timeout=15
            )
            logger.debug(f"IMAP SSL连接建立：{self.host}:{self.port}")

            # 登录邮箱（支持中文账号）
            try:
                self.conn.login(self.user.encode("utf-8"), self.password.encode("utf-8"))
            except imaplib.IMAP4.error as e:
                error_msg = str(e).strip()
                if "authentication failed" in error_msg.lower():
                    raise RuntimeError(
                        "邮箱登录失败！请检查：\n"
                        "1. 客户端专用密码是否正确\n"
                        "2. IMAP服务是否已开启\n"
                        "3. 账号是否有IMAP访问权限"
                    ) from e
                raise RuntimeError(f"IMAP登录错误：{error_msg}") from e

            # 选择收件箱
            status, _ = self.conn.select(self.mailbox, readonly=True)
            if status != "OK":
                raise RuntimeError(f"选择收件箱失败：{self.mailbox}")

            logger.info(f"邮箱登录成功：{self.user}")
            return self

        except TimeoutError:
            raise RuntimeError(f"连接超时：{self.host}:{self.port}（检查网络）")
        except ConnectionRefusedError:
            raise RuntimeError(f"连接被拒绝：{self.host}:{self.port}（检查主机/端口）")
        except Exception as e:
            raise RuntimeError(f"IMAP初始化失败：{str(e)}") from e

    def __exit__(self, exc_type, exc_value, traceback):
        if self.conn:
            try:
                if self.conn.state == "SELECTED":
                    self.conn.close()
                self.conn.logout()
                logger.info("IMAP连接已安全关闭")
            except Exception as e:
                logger.warning(f"关闭IMAP连接警告：{str(e)}")

    def fetch_emails(self) -> Generator[Tuple[str, Message], None, None]:
        """获取指定日期范围内的书法班报名邮件"""
        if not self.conn:
            logger.error("IMAP连接未初始化，无法获取邮件")
            return

        # 解析日期范围（从环境变量读取）
        try:
            start_date_str = os.environ.get("START_DATE", "01-Mar-2025")
            end_date_str = os.environ.get("END_DATE", datetime.now().strftime("%d-%b-%Y"))
            # 验证日期格式（IMAP要求：dd-Mon-yyyy）
            datetime.strptime(start_date_str, "%d-%b-%Y")
            datetime.strptime(end_date_str, "%d-%b-%Y")
        except ValueError as e:
            logger.error(f"日期格式错误：{str(e)}（正确格式：01-Mar-2025）")
            return

        # 搜索邮件（按日期范围+主题关键词）
        try:
            search_criteria = f'(SINCE "{start_date_str}" BEFORE "{end_date_str}")'
            logger.info(f"邮件搜索条件：{search_criteria}")
            status, data = self.conn.uid("SEARCH", None, search_criteria)

            if status != "OK":
                error_msg = data[0].decode("utf-8", errors="replace") if data else "未知错误"
                logger.error(f"邮件搜索失败：{error_msg}")
                return

            # 提取邮件UID
            uids = data[0].split() if data and data[0] else []
            if not uids:
                logger.info(f"未找到{start_date_str}至{end_date_str}的邮件")
                return
            logger.info(f"找到{len(uids)}封待处理邮件")

            # 遍历邮件并过滤“书法班”主题
            for uid_bytes in uids:
                uid = uid_bytes.decode("utf-8").strip()
                if not uid:
                    continue

                try:
                    # 获取邮件原始内容
                    status, msg_data = self.conn.uid("FETCH", uid, "(RFC822)")
                    if status != "OK":
                        logger.error(f"获取邮件{uid}失败：{msg_data}")
                        continue

                    # 解析邮件数据
                    raw_email = None
                    for part in msg_data:
                        if isinstance(part, tuple) and len(part) >= 2:
                            raw_email = part[1]
                            break

                    if not isinstance(raw_email, bytes):
                        logger.error(f"邮件{uid}数据格式异常：{type(raw_email)}")
                        continue

                    # 解析为Message对象
                    from email import message_from_bytes  # 避免循环导入
                    msg = message_from_bytes(raw_email)

                    # 过滤含“书法班”的主题
                    subject = self._decode_subject(msg)
                    if "书法班" in subject:
                        yield uid, msg
                    else:
                        logger.debug(f"邮件{uid}非书法班报名（主题：{subject[:50]}...）")

                except Exception as e:
                    logger.error(f"处理邮件{uid}失败（跳过）：{str(e)}", exc_info=True)
                    continue

        except Exception as e:
            logger.error(f"邮件获取流程异常：{str(e)}", exc_info=True)
            return

    @staticmethod
    def _decode_subject(msg: Message) -> str:
        """解码邮件主题（支持中文）"""
        subject = msg.get("Subject", "")
        decoded_parts = []
        for part, charset in decode_header(subject):
            try:
                if isinstance(part, bytes):
                    # 优先尝试常见编码
                    for encoding in [charset, "utf-8", "gb18030", "big5"]:
                        if not encoding:
                            continue
                        try:
                            decoded = part.decode(encoding)
                            break
                        except:
                            continue
                    else:
                        decoded = part.decode("utf-8", errors="replace")
                else:
                    decoded = str(part)
                decoded_parts.append(decoded)
            except Exception as e:
                logger.warning(f"主题解码失败：{str(e)}")
                decoded_parts.append("[解码失败]")
        return "".join(decoded_parts)

# ---------------------- 邮件附件处理器 ----------------------
class EmailProcessor:
    def __init__(self):
        self.attach_dir = DATA_DIR / "attachments"
        self.attach_dir.mkdir(exist_ok=True, parents=True)

    def save_attachments(self, msg: Message, student_id: str, name: str) -> List[Path]:
        """保存邮件中的附件，返回DOCX附件路径列表"""
        attachments = []
        try:
            for part in msg.walk():
                # 跳过非附件部分
                if part.get_content_maintype() == "multipart":
                    continue
                if part.get("Content-Disposition") is None:
                    continue

                # 获取附件文件名
                filename = part.get_filename()
                if not filename:
                    continue

                # 解码文件名（支持中文）
                decoded_filename = self._decode_filename(filename)
                # 生成安全文件名（避免重复/非法字符）
                file_ext = Path(decoded_filename).suffix.lower()
                safe_filename = f"{student_id}_{name}_{datetime.now().strftime('%Y%m%d%H%M%S')}{file_ext}"
                file_path = self.attach_dir / safe_filename

                # 保存附件（仅保留DOCX格式）
                if file_ext == ".docx":
                    try:
                        payload = part.get_payload(decode=True)
                        if isinstance(payload, bytes) and len(payload) > 0:
                            with open(file_path, "wb") as f:
                                f.write(payload)
                            attachments.append(file_path)
                            logger.debug(f"保存DOCX附件：{file_path.name}")
                        else:
                            logger.warning(f"附件{decoded_filename}内容为空或非字节类型")
                    except Exception as e:
                        logger.error(f"保存附件{decoded_filename}失败：{str(e)}")
            return attachments
        except Exception as e:
            logger.error(f"附件处理异常：{str(e)}", exc_info=True)
            return attachments

    @staticmethod
    def _decode_filename(filename: str) -> str:
        """解码附件文件名（支持中文）"""
        try:
            decoded_parts = decode_header(filename)
            return "".join([
                part.decode(charset or "utf-8", errors="replace") 
                if isinstance(part, bytes) else str(part)
                for part, charset in decoded_parts
            ])
        except Exception as e:
            logger.warning(f"文件名解码失败：{str(e)}")
            return f"unknown_{datetime.now().strftime('%Y%m%d%H%M%S')}"

# ---------------------- DOCX申请材料解析器 ----------------------
from docx import Document

def parse_docx(filepath: str | Path) -> Dict[str, Any]:
    """解析DOCX申请材料，提取资助对象状态和申请理由长度"""
    result = {
        "is_supported": False,  # 是否为资助对象
        "reason_length": 0      # 申请理由中文字数
    }
    filepath = Path(filepath)

    # 前置校验
    if not filepath.exists():
        logger.warning(f"DOCX文件不存在：{filepath.name}")
        return result
    if filepath.stat().st_size == 0:
        logger.warning(f"DOCX文件为空：{filepath.name}")
        return result

    try:
        # 打开DOCX文件
        try:
            doc = Document(filepath)
        except Exception as e:
            logger.error(f"打开DOCX文件失败：{filepath.name} - {str(e)}", exc_info=True)
            return result

        # 遍历表格提取关键信息（适配常见申请表格式）
        support_flag = False
        reason_text = ""
        for table in doc.tables:
            for row in table.rows:
                for cell_idx, cell in enumerate(row.cells):
                    cell_text = cell.text.strip()
                    if not cell_text:
                        continue

                    # 1. 判断是否为资助对象（关键词匹配）
                    if any(kw in cell_text for kw in ["学生资助对象", "贫困生", "资助资格"]):
                        # 查找下一个单元格的答案
                        if cell_idx + 1 < len(row.cells):
                            answer = row.cells[cell_idx + 1].text.strip()
                            support_flag = any(yes in answer for yes in ["是", "√", "确认", "有"])
                            support_flag = support_flag and not any(no in answer for no in ["否", "×", "无"])

                    # 2. 提取申请理由（关键词匹配）
                    if any(kw in cell_text for kw in ["申请理由", "申请原因", "报名说明"]):
                        # 提取当前行及下一行的内容（适配多行理由）
                        reason_parts = [cell_text]
                        # 当前行后续单元格
                        for idx in range(cell_idx + 1, len(row.cells)):
                            reason_parts.append(row.cells[idx].text.strip())
                        # 下一行所有单元格
                        try:
                            next_row = table.rows[table.rows.index(row) + 1]
                            for next_cell in next_row.cells:
                                reason_parts.append(next_cell.text.strip())
                        except IndexError:
                            pass  # 无下一行则跳过
                        reason_text = "".join(reason_parts)

        # 计算申请理由中文字数（仅保留中文）
        chinese_chars = re.findall(r"[\u4e00-\u9fa5]", reason_text)
        result["is_supported"] = support_flag
        result["reason_length"] = len(chinese_chars)

        logger.info(
            f"DOCX解析结果：{filepath.name}\n"
            f"是否资助对象：{'是' if support_flag else '否'}\n"
            f"申请理由字数：{len(chinese_chars)}字（要求≥{MIN_REASON_LENGTH}字）"
        )
        return result

    except Exception as e:
        logger.error(f"DOCX解析异常：{filepath.name} - {str(e)}", exc_info=True)
        return result

# ---------------------- 主题解析函数 ----------------------
def parse_subject(subject: str) -> Tuple[Optional[str], Optional[str]]:
    """解析邮件主题，提取姓名和学号"""
    if not subject:
        return None, None

    try:
        # 清理主题（移除多余空格）
        clean_subject = re.sub(r"\s+", "", subject)
        match = SUBJECT_PATTERN.match(clean_subject)
        if match:
            name = match.group(1).strip()
            student_id = match.group(2).strip()
            # 清理学号（移除可能的短横线/空格）
            student_id = student_id.replace("-", "").replace(" ", "")
            return name, student_id
        logger.debug(f"主题格式不匹配：{subject[:50]}...（正确格式：姓名+学号+书法班报名申请）")
        return None, None
    except Exception as e:
        logger.error(f"主题解析异常：{subject[:50]}... - {str(e)}", exc_info=True)
        return None, None

# ---------------------- 主筛选流程 ----------------------
def main():
    logger.info("="*60 + " 书法班报名筛选流程启动 " + "="*60)
    # 初始化结果容器
    admitted: List[Dict[str, str]] = []  # 录取名单
    rejected: List[Dict[str, str]] = []  # 拒绝名单
    candidates: List[Tuple[str, str, datetime]] = []  # 候补名单（学号、姓名、邮件时间）

    try:
        # 1. 读取基础名单（新鸿基/去年/黑名单）
        logger.info("="*30 + " 读取基础名单 " + "="*30)
        new_hongji_ids = read_student_list(NEW_HONGJI_FILE)
        last_year_ids = read_student_list(LAST_YEAR_FILE)
        blacklist_ids = read_student_list(BLACKLIST_FILE) if os.path.exists(BLACKLIST_FILE) else set()

        # 校验必选名单有效性
        if not new_hongji_ids:
            raise RuntimeError("新鸿基名单读取失败或无有效学号，无法继续筛选")
        if not last_year_ids:
            raise RuntimeError("去年报名名单读取失败或无有效学号，无法继续筛选")

        # 2. 解析筛选日期范围
        logger.info("="*30 + " 解析筛选日期 " + "="*30)
        try:
            start_date_str = os.environ.get("START_DATE", "01-Mar-2025")
            end_date_str = os.environ.get("END_DATE", datetime.now().strftime("%d-%b-%Y"))
            start_date = datetime.strptime(start_date_str, "%d-%b-%Y").replace(tzinfo=timezone.utc)
            end_date = datetime.strptime(end_date_str, "%d-%b-%Y").replace(tzinfo=timezone.utc)
            logger.info(f"筛选日期范围：{start_date_str} 至 {end_date_str}")
        except ValueError as e:
            raise RuntimeError(f"日期解析失败：{str(e)}（正确格式：01-Mar-2025）") from e

        # 3. 初始化工具类
        email_processor = EmailProcessor()

        # 4. 连接邮箱并处理邮件
        logger.info("="*30 + " 处理报名邮件 " + "="*30)
        email_count = 0  # 总处理邮件数
        error_count = 0  # 处理错误数

        with SecureIMAPClient() as imap_client:
            for uid, msg in imap_client.fetch_emails():
                email_count += 1
                try:
                    # 提取邮件接收时间
                    recv_date = None
                    date_str = msg.get("Date")
                    if date_str:
                        recv_date = parsedate_to_datetime(date_str)
                        # 统一时区为UTC
                        if recv_date.tzinfo is None:
                            recv_date = recv_date.replace(tzinfo=timezone.utc)
                        else:
                            recv_date = recv_date.astimezone(timezone.utc)

                    # 日期过滤（确保在筛选范围内）
                    if not recv_date or not (start_date <= recv_date <= end_date):
                        logger.debug(f"邮件{uid}时间不在范围（{recv_date}），跳过")
                        continue

                    # 解析主题提取姓名和学号
                    subject = imap_client._decode_subject(msg)
                    name, student_id = parse_subject(subject)

                    # 主题格式校验
                    if not name or not student_id:
                        rejected.append({
                            "学号": "未知",
                            "姓名": "未知",
                            "原主题": subject[:100],  # 截断过长主题
                            "原因": "主题格式错误（正确示例：张三12345678书法班报名申请）"
                        })
                        continue

                    # 打印当前处理的学生信息（便于调试）
                    logger.info(f"处理邮件{uid}：姓名={name}，学号={student_id}，时间={recv_date}")

                    # 5. 黑名单过滤
                    if student_id in blacklist_ids:
                        rejected.append({
                            "学号": student_id,
                            "姓名": name,
                            "原因": "黑名单用户，不参与筛选"
                        })
                        continue

                    # 6. 新鸿基学生直接录取
                    if student_id in new_hongji_ids:
                        admitted.append({
                            "学号": student_id,
                            "姓名": name,
                            "备注": "新鸿基推荐（直接录取）"
                        })
                        continue

                    # 7. 去年已录取学生拒绝
                    if student_id in last_year_ids:
                        rejected.append({
                            "学号": student_id,
                            "姓名": name,
                            "原因": "去年已录取，本次不重复录取"
                        })
                        continue

                    # 8. 处理附件（仅保留DOCX）
                    attachments = email_processor.save_attachments(msg, student_id, name)
                    docx_files = [a for a in attachments if a.suffix.lower() == ".docx"]

                    if not docx_files:
                        rejected.append({
                            "学号": student_id,
                            "姓名": name,
                            "原因": "缺少DOCX格式申请材料"
                        })
                        continue

                    # 9. 解析DOCX申请材料
                    docx_info = parse_docx(docx_files[0])

                    # 10. 申请材料校验
                    # 校验1：是否为资助对象
                    if not docx_info["is_supported"]:
                        rejected.append({
                            "学号": student_id,
                            "姓名": name,
                            "原因": "非学生资助对象，不符合报名条件"
                        })
                        continue

                    # 校验2：申请理由字数
                    if docx_info["reason_length"] < MIN_REASON_LENGTH:
                        rejected.append({
                            "学号": student_id,
                            "姓名": name,
                            "原因": f"申请理由字数不足（{docx_info['reason_length']}字，要求≥{MIN_REASON_LENGTH}字）"
                        })
                        continue

                    # 11. 加入候补名单（按邮件时间排序）
                    candidates.append((student_id, name, recv_date))
                    logger.info(f"加入候补名单：{name}（{student_id}），邮件时间：{recv_date}")

                except Exception as e:
                    # 单邮件处理失败，记录错误后继续
                    error_count += 1
                    logger.error(f"处理邮件{uid}失败（跳过）：{str(e)}", exc_info=True)
                    rejected.append({
                        "学号": "未知",
                        "姓名": "未知",
                        "原主题": imap_client._decode_subject(msg)[:100] if msg else "未知",
                        "原因": f"邮件处理异常：{str(e)[:50]}"
                    })
                    continue

        # 12. 处理候补名单（按邮件时间先到先得）
        logger.info("="*30 + " 处理候补名单 " + "="*30)
        logger.info(f"候补名单人数：{len(candidates)}人，已录取新鸿基学生：{len(admitted)}人")
        remaining_quota = ADMISSION_QUOTA - len(admitted)
        logger.info(f"总录取名额：{ADMISSION_QUOTA}人，剩余候补名额：{remaining_quota}人")

        if remaining_quota > 0 and candidates:
            # 按邮件接收时间升序排序（先到先得）
            candidates.sort(key=lambda x: x[2])
            # 录取前N名候补
            admit_candidates = candidates[:remaining_quota]
            reject_candidates = candidates[remaining_quota:]

            # 新增录取名单
            for sid, name, recv_date in admit_candidates:
                admitted.append({
                    "学号": sid,
                    "姓名": name,
                    "备注": f"非新鸿基（候补录取，邮件时间：{recv_date.strftime('%Y-%m-%d %H:%M')}"
                })
                logger.info(f"候补录取：{name}（{sid}），邮件时间：{recv_date}")

            # 名额已满拒绝剩余候补
            for sid, name, _ in reject_candidates:
                rejected.append({
                    "学号": sid,
                    "姓名": name,
                    "原因": f"符合条件但名额已满（总名额{ADMISSION_QUOTA}人，已录满）"
                })
                logger.info(f"候补拒绝：{name}（{sid}），原因：名额已满")

        elif remaining_quota <= 0 and candidates:
            # 无剩余名额，所有候补拒绝
            for sid, name, _ in candidates:
                rejected.append({
                    "学号": sid,
                    "姓名": name,
                    "原因": f"符合条件但名额已满（总名额{ADMISSION_QUOTA}人，新鸿基已录满）"
                })
            logger.warning(f"无剩余名额，{len(candidates)}名候补全部拒绝")

        # 13. 保存筛选结果
        logger.info("="*30 + " 保存筛选结果 " + "="*30)
        save_results(admitted, rejected)

        # 14. 打印最终统计
        logger.info("="*60 + " 筛选流程完成 " + "="*60)
        logger.info(f"最终统计：")
        logger.info(f"- 总处理邮件数：{email_count}封")
        logger.info(f"- 处理错误数：{error_count}封")
        logger.info(f"- 录取人数：{len(admitted)}人（新鸿基：{len([x for x in admitted if '新鸿基' in x['备注']])}人，候补：{len([x for x in admitted if '候补' in x['备注']])}人）")
        logger.info(f"- 拒绝人数：{len(rejected)}人")
        logger.info(f"- 录取名单已保存：{ADMITTED_FILE.name}")
        logger.info(f"- 拒绝名单已保存：{REJECTED_FILE.name}")

    except RuntimeError as e:
        # 致命错误，保存已有结果后退出
        logger.critical(f"筛选流程致命错误：{str(e)}", exc_info=True)
        if admitted or rejected:
            try:
                save_results(admitted, rejected)
                logger.warning(f"已保存部分结果（录取{len(admitted)}人，拒绝{len(rejected)}人）")
            except Exception as save_e:
                logger.error(f"保存部分结果失败：{str(save_e)}")
        raise  # 抛出错误，让前端捕获返回码

    except Exception as e:
        # 未预期异常，兜底处理
        logger.critical(f"筛选流程未预期异常：{str(e)}", exc_info=True)
        if admitted or rejected:
            try:
                save_results(admitted, rejected)
            except:
                pass
        raise RuntimeError(f"未预期异常：{str(e)}") from e

if __name__ == "__main__":
    main)vv
