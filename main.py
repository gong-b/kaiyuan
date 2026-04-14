# -*- coding: utf-8 -*-
# 主业务逻辑：串联所有模块，执行审核流程
import tempfile
import os
from config import *
from email_client import EmailClient
from email_processor import EmailProcessor
from docx_parser import DocxParser
from excel_handler import ExcelHandler

class KaiYuanAuditSystem:
    def __init__(self):
        # 初始化所有工具
        self.email_client = None
        self.email_processor = EmailProcessor(MAIL_KEYWORDS)
        self.docx_parser = DocxParser(MIN_REASON_LENGTH)
        self.excel_handler = ExcelHandler()

        # 数据存储
        self.raw_mails = []
        self.filtered_mails = []
        self.admit_list = []
        self.reject_list = []
        self.hongji_ids = set()
        self.last_ids = set()
        self.black_ids = set()

    # 加载名单
    def load_lists(self, hongji_file, last_file, black_file):
        self.hongji_ids = self.excel_handler.load_id_list(hongji_file)
        self.last_ids = self.excel_handler.load_id_list(last_file)
        self.black_ids = self.excel_handler.load_id_list(black_file)
        return True

    # 连接邮箱+抓取邮件
    def fetch_mails(self, user, pwd, start_date, end_date, progress_bar=None, status_text=None):
        # 初始化邮箱客户端
        self.email_client = EmailClient(IMAP_SERVER, IMAP_PORT, user, pwd)
        # 连接
        status, msg = self.email_client.connect()
        if not status:
            return False, msg
        
        # 打开文件夹
        status, msg = self.email_client.select_folder(MAIL_FOLDER, FALLBACK_FOLDER)
        if not status:
            return False, msg
        
        # 抓取邮件
        self.raw_mails, msg = self.email_client.fetch_emails(start_date, end_date, progress_bar, status_text)
        # 过滤邮件
        self.filtered_mails = self.email_processor.filter_mails(self.raw_mails)
        return True, f"抓取完成，共 {len(self.filtered_mails)} 封有效报名邮件"

    # 执行自动审核
    def audit_mails(self):
        if not self.filtered_mails:
            return False, "请先抓取邮件"
        
        self.admit_list = []
        self.reject_list = []
        tmp_dir = tempfile.mkdtemp()

        for mail in self.filtered_mails:
            sid = mail["sid"]
            name = mail["name"]

            # 基础校验
            if not sid or not name:
                self.reject_list.append({
                    "学号": "未知", "姓名": "未知", "原因": "邮件主题格式错误，无法识别姓名学号"
                })
                continue

            # 黑名单/去年参加/新鸿基优先
            if sid in self.black_ids:
                self.reject_list.append({"学号": sid, "姓名": name, "原因": "黑名单"})
                continue
            if sid in self.last_ids:
                self.reject_list.append({"学号": sid, "姓名": name, "原因": "去年已参加"})
                continue
            if sid in self.hongji_ids:
                self.admit_list.append({"学号": sid, "姓名": name, "审核结果": "新鸿基直接录取"})
                continue

            # 提取附件
            att_dir = os.path.join(tmp_dir, sid)
            attachments = self.email_processor.extract_attachments(mail, att_dir)
            if not attachments:
                self.reject_list.append({"学号": sid, "姓名": name, "原因": "未找到docx附件"})
                continue

            # 解析附件
            doc_info = self.docx_parser.parse_application(attachments[0])
            if not doc_info["is_supported"]:
                self.reject_list.append({"学号": sid, "姓名": name, "原因": "非学生资助对象"})
            elif doc_info["reason_length"] < MIN_REASON_LENGTH:
                self.reject_list.append({
                    "学号": sid, "姓名": name, 
                    "原因": f"申请理由字数不足（{doc_info['reason_length']}字）"
                })
            else:
                self.admit_list.append({"学号": sid, "姓名": name, "审核结果": "审核通过"})

        return True, f"审核完成：录取 {len(self.admit_list)} 人，拒绝 {len(self.reject_list)} 人"

    # 导出结果
    def export_results(self):
        admit_data, _ = self.excel_handler.export_to_excel(self.admit_list, "录取名单.xlsx")
        reject_data, _ = self.excel_handler.export_to_excel(self.reject_list, "拒绝名单.xlsx")
        return admit_data, reject_data
