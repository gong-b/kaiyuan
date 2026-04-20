from docx import Document
import re

class DocxParser:
    @staticmethod
    def parse(path):
        # 初始化结果，sid设为None便于后续逻辑判断
        res = {"is_supported": False, "reason_length": 0, "name": "未知", "sid": None}
        try:
            doc = Document(path)
            if not doc.tables:
                return res
            
            table = doc.tables[0]
            for row in table.rows:
                cells = row.cells
                for i, c in enumerate(cells):
                    t = c.text.strip()
                    
                    # 1. 提取姓名：匹配到“姓名”后，取下一个单元格
                    if t == "姓名" and i + 1 < len(cells):
                        res["name"] = cells[i+1].text.strip()
                    
                    # 2. 提取学号：匹配到“学号”后，取下一个单元格
                    if t == "学号" and i + 1 < len(cells):
                        raw_sid = cells[i+1].text.strip()
                        # 仅保留数字，防止学生填入“学号：3230...”等杂质
                        res["sid"] = "".join(filter(str.isdigit, raw_sid))

                    # 3. 资助对象判定
                    if "是否为学生资助对象" in t and i + 1 < len(cells):
                        v = cells[i + 1].text.strip()
                        # 兼容“是”、“是的”、“是 ”等情况
                        res["is_supported"] = ("是" in v) and ("不是" not in v)
                    
                    # ========== 优化：申请理由解析（兼容多种提示语格式） ==========
                    if "申请理由" in t:
                        # 正则匹配并移除所有申请理由提示前缀（兼容“申请理由：”“申请理由（不少于100字）：”等）
                        pattern = re.compile(r"申请理由\s*[:：]\s*|申请理由\s*（.*?）\s*[:：]\s*")
                        content = pattern.sub("", c.text).strip()
                        clean_content = re.sub(r"\s+", "", content)  # 移除所有空格
                        res["reason_length"] = len(clean_content)
        except Exception as e:
            print(f"解析附件出错: {e}")
        return res
