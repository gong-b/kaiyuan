import imaplib
import ssl
import base64
import logging
from email import message_from_bytes

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

class SecureIMAPClient:
    def __init__(self, user, password, folder="INBOX", imap_server="imap.zju.edu.cn", imap_port=993):
        self.user = user
        self.password = password
        self.folder = folder
        self.imap_server = imap_server  # 改为可配置
        self.imap_port = imap_port      # 改为可配置
        self.conn = None  # 初始化连接对象

    def __enter__(self):
        try:
            ctx = ssl.create_default_context()
            self.conn = imaplib.IMAP4_SSL(self.imap_server, self.imap_port, ssl_context=ctx)
            logger.info(f"成功连接 IMAP 服务器：{self.imap_server}:{self.imap_port}")
            
            # 登录
            self.conn.login(self.user, self.password)
            logger.info(f"用户 {self.user} 登录成功")
            
            # 选中文件夹（仅执行一次，使用编码后的名称）
            encoded_folder = imap_utf7_encode(self.folder)
            select_status, select_data = self.conn.select(encoded_folder)
            if select_status != 'OK':
                raise Exception(f"选中文件夹失败：{self.folder}（编码后：{encoded_folder}），返回：{select_data}")
            logger.info(f"成功选中文件夹：{self.folder}（邮件总数：{select_data[0].decode()}）")
            return self
        except Exception as e:
            logger.error(f"IMAP 连接/登录/选文件夹失败：{str(e)}")
            # 确保连接关闭
            if self.conn:
                try:
                    self.conn.logout()
                except:
                    pass
            raise  # 抛出异常，让调用方感知

    def __exit__(self, exc_type, exc_val, exc_tb):
        if self.conn:
            try:
                self.conn.logout()
                logger.info("IMAP 连接已正常退出")
            except Exception as e:
                logger.error(f"IMAP 退出时出错：{str(e)}")
        # 若有异常，返回 False 让异常继续抛出（符合上下文管理器规范）
        return False

    def fetch_emails(self, since_date):
        """
        拉取指定日期之后的邮件
        :param since_date: IMAP 格式的日期（如 "01-Jan-2024"）
        :yield: (uid, email_message)
        """
        if not self.conn:
            raise RuntimeError("IMAP 连接未初始化，请使用 with 语句创建实例")
        
        try:
            # SEARCH 语法修正：ALL 放在条件前
            search_status, search_data = self.conn.uid('SEARCH', 'ALL', 'SINCE', since_date)
            if search_status != 'OK':
                raise Exception(f"邮件搜索失败，返回状态：{search_status}，数据：{search_data}")
            
            # 处理无邮件的情况
            if not search_data[0]:
                logger.info(f"未找到 {since_date} 之后的邮件")
                return
            
            uids = search_data[0].split()
            logger.info(f"找到 {len(uids)} 封符合条件的邮件，开始拉取...")
            
            # 迭代拉取每封邮件
            for uid in uids:
                fetch_status, fetch_data = self.conn.uid('FETCH', uid, '(RFC822)')
                if fetch_status != 'OK':
                    logger.warning(f"拉取 UID {uid.decode()} 的邮件失败，跳过")
                    continue
                # 解析邮件内容
                email_msg = message_from_bytes(fetch_data[0][1])
                yield uid.decode(), email_msg
                
        except Exception as e:
            logger.error(f"拉取邮件时出错：{str(e)}")
            raise  # 抛出异常，让调用方处理
