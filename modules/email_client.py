import imaplib
import ssl
import base64
import logging
from email import message_from_bytes
from email.utils import parseaddr

logger = logging.getLogger(__name__)
# 补充日志配置（可选，方便调试）
logging.basicConfig(level=logging.INFO, format="%(asctime)s - %(name)s - %(levelname)s - %(message)s")

def imap_utf7_encode(text):
    try:
        def _modified_base64(s):
            return base64.b64encode(s.encode('utf-16-be')).decode('ascii').rstrip('=').replace('/', ',')
        res = []
        i = 0
        while i < len(text):
            c = text[i]
            if 0x20 <= ord(c) <= 0x7e:
                res.append('&-') if c == '&' else res.append(c)
                i += 1
            else:
                j = i
                while j < len(text) and not (0x20 <= ord(text[j]) <= 0x7e):
                    j += 1
                res.append('&' + _modified_base64(text[i:j]) + '-')
                i = j
        encoded_text = "".join(res)
        logger.debug(f"文件夹名编码：{text} -> {encoded_text}")
        return encoded_text
    except Exception as e:
        logger.error(f"IMAP UTF7 编码失败：{text}，错误：{str(e)}")
        return text

logger = logging.getLogger(__name__)
class SecureIMAPClient:
    def __init__(self, user, password, folder="INBOX"):
        self.user = user
        self.password = password
        self.folder = folder
        self.conn = None

    def __enter__(self):
        ctx = ssl.create_default_context()
        self.conn = imaplib.IMAP4_SSL("imap.zju.edu.cn", 993, ssl_context=ctx)
        self.conn.login(self.user, self.password)
        
        # 优化点：对文件夹名进行 UTF7 编码
        encoded_folder = imap_utf7_encode(self.folder)
        status, _ = self.conn.select(encoded_folder)
        if status != 'OK':
            logger.warning(f"无法进入文件夹 {self.folder}，尝试进入 INBOX")
            self.conn.select("INBOX")
        return self

    def __exit__(self, *args):
        try:
            if self.conn:
                self.conn.close()
                self.conn.logout()
        except:
            pass

    def fetch_emails(self, since_date):
        """
        since_date 格式应为 "DD-Mon-YYYY" (例如 "01-Apr-2024")
        """
        if not self.conn:
            return

        # 优化点：使用 ALL 搜索，不要使用 UNANSWERED，否则你回复过的邮件会搜不到
        status, data = self.conn.uid('SEARCH', 'ALL', 'SINCE', since_date)
        
        if status != 'OK' or not data[0]:
            logger.info(f"日期 {since_date} 之后未找到邮件")
            return

        uids = data[0].split()
        for uid in uids:
            try:
                # 获取整封邮件
                status, fetch_data = self.conn.uid('FETCH', uid, '(RFC822)')
                if status != 'OK' or not fetch_data[0]:
                    continue
                
                raw_email = fetch_data[0][1]
                msg = message_from_bytes(raw_email)
                
                # 【关键修复】：过滤掉发件人是您自己的邮件（解决“全是拒绝”的问题）
                from_header = msg.get("From", "")
                _, from_addr = parseaddr(from_header)
                if from_addr.lower() == self.user.lower():
                    logger.info(f"跳过发件人为自己的邮件: UID {uid.decode()}")
                    continue
                
                yield uid.decode(), msg
                
            except Exception as e:
                logger.error(f"解析邮件 UID {uid} 出错: {str(e)}")
                continue
