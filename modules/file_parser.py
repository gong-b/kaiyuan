import re
from docx import Document

class FileParser:
    @staticmethod
    def parse(path):
        return FileParser._parse_docx(path)

    @staticmethod
    def _parse_docx(path):
        res = {
            "is_supported": False,
            "reason_length": 0,
            "name": "未知",
            "sid": "",
            "apply_class": "",
            "phone": ""
        }

        try:
            doc = Document(path)

            # 提取班级
            for para in doc.paragraphs:
                text = para.text.strip()
                match = re.search(r"(.+班)报名申请表", text)
                if match:
                    res["apply_class"] = match.group(1)
                    break

            if not doc.tables:
                return res

            table = doc.tables[0]

            for row in table.rows:
                cells = [c.text.strip() for c in row.cells]
                for i, text in enumerate(cells):

                    # 姓名
                    if "姓名" in text and i + 1 < len(cells):
                        res["name"] = cells[i+1]

                    # 学号
                    if "学号" in text and i + 1 < len(cells):
                        res["sid"] = "".join(filter(str.isdigit, cells[i+1]))

                    # 联系方式（只取后一格）
                    if "联系方式" in text and i + 1 < len(cells):
                        res["phone"] = "".join(filter(str.isdigit, cells[i+1]))

                    # 是否资助对象（只看第一个 是/不是）
                    if "是否为学生资助对象" in text:
                        for j in range(i + 1, len(cells)):
                            val = cells[j].strip()
                            if val == "是":
                                res["is_supported"] = True
                                break
                            elif val == "不是":
                                res["is_supported"] = False
                                break

                    # 申请理由
                    if "申请理由" in text:
                        content = cells[i] if len(cells[i]) > 20 else (cells[i+1] if i+1 < len(cells) else "")
                        content = re.sub(r"申请理由.*?[:：]", "", content).strip()
                        res["reason_length"] = len(re.sub(r"\s+", "", content))

        except Exception:
            pass

        return res
