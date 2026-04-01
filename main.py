import日志记录
来自电子邮件客户端import电子邮件客户端
来自Excel处理程序importExcel处理程序
来自Docx解析器importDocx解析器
来自 email_processor import EmailProcessor
import config

logging.basicConfig(level=logging.INFO, format="%(message)s")

def 主():
    logging.info("✅ 开始筛选")

    # 读取上传的3个名单
    excel = Excel处理程序()
xhj_ids = excel。读取学生ID("新鸿基名单.xlsx")
    black_ids = excel.读取学生ID("黑名单.xlsx")
    last_ids = excel.读取学生ID("去年名单.xlsx")

    # 收邮件
    client = EmailClient(配置。EMAIL_HOST, 配置。EMAIL_PORT, 配置。EMAIL_USER, 配置。EMAIL_PASS)
邮件 = 客户端.fetch_mails()
info(f" 共收取邮件：{len(邮件)}封")

接受 = []
拒绝 = []

    for mail in mails:
        sid = mail.get("student_id")
        name = mail.get("name")
        if not sid:
            继续

        # 规则
        if sid in black_ids:
            reject.append([sid, name, "黑名单"])
            继续
        if sid in last_ids:
            reject.append([sid, name, "去年已参加"])
            继续
        if sid not in xhj_ids:
            reject.append([sid, name, "非新鸿基"])
            继续

        # 附件检查
        doc_path = mail.get("attachment_path")
        如果 未提供文档路径：
            拒绝.追加([, name, "无申请表"])
            continue

        # 解析Word
        解析器 = DocxParser(doc_path)
是否为补贴 =解析器.是否为补贴()
        word_count = parser.count_reason()
        logging.info(f"{sid} 字数：{word_count}")

        if not is_subsidy:
            拒绝.追加([sid, , "非资助对象"])
            continue
        如果词数 <config.REASON_MIN_WORDS:
            拒绝.追加([,name“字数不足{config.REASON_MIN_WORDS}”])
            继续

        接受.追加([sid, name, “通过”])

    # 录取25人
    接受 = 接受[:.MAX_ACCEPT]
    excel.写入接受(接受)
    excel.写入拒绝(拒绝)

    logging.info(f"\n🎯 录取：{len(accept)}人")
    logging.info("✅ 录取名单.xlsx 已生成")
    logging.info("✅ 拒绝名单.xlsx 已生成")

如果__name__ =="__main__":
    主程序()
