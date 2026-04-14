# -*- coding: utf-8 -*-
# 邮箱客户端：负责连接、打开文件夹、抓取邮件
import imaplib
import ssl
import base64
from datetime import datetime, timedelta
from email import message_from_bytes
from email.header import decode_header

class EmailClient:
    def __init__(self, server, port, user, pwd):
        self.server = server
        self.port = port
        self.user = user
        self.pwd = pwd
        self.conn = None

    # 🔥 浙大邮箱中文文件夹专用UTF-7编码
    def _imap_utf7_encode(self, s):
        res = []
        for c in s:
            if ord(c) < 128 and c != '&':
                res.append(c)
            elif c == '&':
                res.append('&-')
            else:
                res.append('&' + base64.b64encode(c.encode('utf-16be')).decode().strip() + '-')
        return ''.join(res)

    # 连接邮箱
    def connect(self):
        try:
            ctx = ssl.create_default_context()
            self.conn = imaplib.IMAP4_SSL(self.server, self.port, ssl_context=ctx, timeout=30)
            self.conn.login(self.user, self.pwd)
            return True, "连接成功"
        except Exception as e:
            return False, f"连接失败：{str(e)}"

    # 打开指定文件夹
    def select_folder(self, folder_name, fallback_folder):
        if not self.conn:
            return False, "未连接邮箱"
        
        # 尝试打开目标文件夹
        encoded_folder = self._imap_utf7_encode(folder_name)
        status, data = self.conn.select(encoded_folder, readonly=True)
        if status == "OK":
            return True, f"成功打开文件夹：{folder_name}"
        
        # 失败则打开备用文件夹（收件箱）
        status, data = self.conn.select(fallback_folder, readonly=True)
        if status == "OK":
            return True, f"目标文件夹打开失败，已打开备用文件夹：{fallback_folder}"
        
        return False, f"所有文件夹打开失败：{data}"

    # 抓取指定时间范围的邮件
    def fetch_emails(self, start_date, end_date, progress_bar=None, status_text=None):
        if not self.conn:
            return [], "未连接邮箱"
        
        # 转换日期格式
        since = start_date.strftime("%d-%b-%Y")
        before = (end_date + timedelta(days=1)).strftime("%d-%b-%Y")
        
        # 搜索邮件
        status, data = self.conn.uid('SEARCH', None, 'SINCE', since, 'BEFORE', before)
        if status != "OK" or not data[0]:
            return [], "未找到符合条件的邮件"
        
        uids = data[0].split()
        total = len(uids)
        mails = []

        for i, uid in enumerate(uids):
            if progress_bar:
                progress_bar.progress((i+1)/total, text=f"已解析 {i+1}/{total} 封")
            if status_text:
                status_text.text(f"正在解析第 {i+1}/{total} 封邮件")
            
            try:
                _, dat = self.conn.uid('FETCH', uid, '(RFC822)')
                msg = message_from_bytes(dat[0][1])
                # 解析主题
                subject = "".join(
                    part.decode(charset or "utf-8", "replace") if isinstance(part, bytes) else str(part)
                    for part, charset in decode_header(msg.get("Subject", ""))
                )
                mails.append({
                    "uid": uid.decode(),
                    "subject": subject,
                    "msg_obj": msg,
                    "date": msg.get("Date")
                })
            except Exception as e:
                continue
        
        return mails, f"成功抓取 {len(mails)} 封邮件"

    # 关闭连接
    def close(self):
        if self.conn:
            self.conn.logout()
