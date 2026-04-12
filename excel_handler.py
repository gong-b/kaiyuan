import pandas as pd
from typing import Set, Dict, Any
from pathlib import Path
import logging

logger = logging.getLogger(__name__)

def read_student_list(file_path: str | Path) -> Set[str]:
    """
    读取Excel中的学生学号（适配你的文件：支持多列名、多格式、含短横线学号）
    :param file_path: Excel文件路径
    :return: 去重后的有效学号集合
    """
    student_ids = set()
    file_path = Path(file_path)
    file_name = file_path.name  # 简化文件名引用

    # 1. 前置校验：文件存在性+非空
    if not file_path.exists():
        logger.warning(f"【Excel读取】文件不存在：{file_name}")
        return student_ids
    if file_path.stat().st_size == 0:
        logger.warning(f"【Excel读取】文件为空（0字节）：{file_name}")
        return student_ids

    try:
        # 2. 适配Excel格式（同时支持.xlsx和.xls）
        try:
            # 优先用openpyxl读取.xlsx（主流格式）
            xl_file = pd.ExcelFile(file_path, engine="openpyxl")
            engine = "openpyxl"
        except ImportError:
            logger.error("【Excel读取】缺少openpyxl依赖！请执行：pip install openpyxl")
            return student_ids
        except Exception:
            #  fallback用xlrd读取.xls（旧格式）
            try:
                xl_file = pd.ExcelFile(file_path, engine="xlrd")
                engine = "xlrd"
            except ImportError:
                logger.error("【Excel读取】缺少xlrd依赖！请执行：pip install xlrd==2.0.1")
                return student_ids
            except pd.errors.EmptyDataError:
                logger.warning(f"【Excel读取】文件无数据：{file_name}")
                return student_ids
            except Exception as e:
                logger.error(f"【Excel读取】打开.xls文件失败：{file_name} - {str(e)}", exc_info=True)
                return student_ids

        # 3. 关键：适配你的学号列名（可根据实际列名补充）
        ID_COL_KEYWORDS = [
            "学号", "学生学号", "id", "student id", 
            "编号", "学员编号", "工号"  # 常见非标准列名，可根据你的文件调整
        ]
        target_sheet = None  # 目标Sheet名称
        target_col = None    # 目标学号列名

        # 遍历所有Sheet，找到含学号列的表
        for sheet_name in xl_file.sheet_names:
            try:
                # 仅读取前10行检测列名，提升效率
                df_temp = pd.read_excel(xl_file, sheet_name=sheet_name, nrows=10, engine=engine)
                # 遍历列名匹配关键词（不区分大小写、空格）
                for col in df_temp.columns:
                    col_clean = str(col).lower().replace(" ", "").replace("-", "")
                    if any(kw.lower() in col_clean for kw in ID_COL_KEYWORDS):
                        target_sheet = sheet_name
                        target_col = col
                        break  # 找到列则退出列循环
                if target_sheet:
                    break  # 找到Sheet则退出Sheet循环
            except Exception as e:
                logger.debug(f"【Excel读取】检查Sheet[{sheet_name}]失败（跳过）：{str(e)}")
                continue

        # 4. 校验是否找到目标列
        if not target_sheet or not target_col:
            logger.error(
                f"【Excel读取】文件{file_name}未找到学号列！\n"
                f"支持的列名关键词：{', '.join(ID_COL_KEYWORDS)}"
            )
            return student_ids

        # 5. 读取完整数据（仅读取目标列，减少内存占用）
        try:
            df = pd.read_excel(
                xl_file,
                sheet_name=target_sheet,
                usecols=[target_col],  # 仅读取学号列
                dtype=str,  # 强制字符串类型，避免学号前导0丢失
                engine=engine
            )
        except Exception as e:
            logger.error(f"【Excel读取】读取Sheet[{target_sheet}]列[{target_col}]失败：{str(e)}", exc_info=True)
            return student_ids

        # 6. 清洗学号（去空、去重、支持含短横线格式）
        id_series = df[target_col].astype(str).dropna().str.strip()
        # 过滤规则：非空 + 非"nan" + 纯数字/含短横线数字（如2023-12345）
        student_ids = set()
        for sid in id_series:
            if not sid or sid.lower() == "nan":
                continue
            # 移除短横线后判断是否为纯数字
            sid_clean = sid.replace("-", "").replace("_", "")
            if sid_clean.isdigit():
                student_ids.add(sid)  # 保留原始格式（如带短横线）
            else:
                logger.debug(f"【Excel读取】无效学号（非数字）：{sid}（文件：{file_name}）")

        logger.info(
            f"【Excel读取】成功！\n"
            f"文件：{file_name} | Sheet：{target_sheet} | 列：{target_col} | 有效学号数：{len(student_ids)}"
        )
        return student_ids

    except Exception as e:
        logger.error(f"【Excel读取】文件{file_name}处理异常：{str(e)}", exc_info=True)
        return student_ids

def save_results(admitted: list[Dict[str, Any]], rejected: list[Dict[str, Any]]):
    """
    保存录取/拒绝结果（确保Excel格式兼容，避免乱码）
    :param admitted: 录取名单（含学号、姓名、备注）
    :param rejected: 拒绝名单（含学号、姓名、原主题、原因）
    """
    from config import ADMITTED_FILE, REJECTED_FILE  # 延迟导入，避免循环依赖

    try:
        # 确保输出目录存在
        for file in [ADMITTED_FILE, REJECTED_FILE]:
            file.parent.mkdir(exist_ok=True, parents=True)

        # ---------------------- 保存录取名单 ----------------------
        if admitted:
            df_admitted = pd.DataFrame(admitted)
            # 补全缺失列（避免KeyError）
            required_cols = ["学号", "姓名"]
            for col in required_cols:
                if col not in df_admitted.columns:
                    df_admitted[col] = ""
            # 补充备注列（可选）
            if "备注" not in df_admitted.columns:
                df_admitted["备注"] = ""
            # 固定列顺序
            df_admitted = df_admitted[["学号", "姓名", "备注"]]

            # 去重（按学号，保留第一条）
            df_admitted = df_admitted.drop_duplicates(subset=["学号"], keep="first")

            # 保存Excel（UTF-8编码，避免中文乱码）
            df_admitted.to_excel(
                ADMITTED_FILE,
                sheet_name="录取名单",
                index=False,
                engine="openpyxl",
                encoding="utf-8"
            )
            logger.info(f"【结果保存】录取名单已保存：{ADMITTED_FILE.name}（{len(df_admitted)}人）")

        # ---------------------- 保存拒绝名单 ----------------------
        if rejected:
            df_rejected = pd.DataFrame(rejected)
            # 补全缺失列
            required_cols = ["学号", "姓名", "原因"]
            for col in required_cols:
                if col not in df_rejected.columns:
                    df_rejected[col] = ""
            # 补充原主题列（可选）
            if "原主题" not in df_rejected.columns:
                df_rejected["原主题"] = ""
            # 固定列顺序
            df_rejected = df_rejected[["学号", "姓名", "原主题", "原因"]]

            # 去重（按学号+原因，避免重复记录）
            df_rejected = df_rejected.drop_duplicates(subset=["学号", "原因"], keep="first")

            # 保存Excel
            df_rejected.to_excel(
                REJECTED_FILE,
                sheet_name="拒绝名单",
                index=False,
                engine="openpyxl",
                encoding="utf-8"
            )
            logger.info(f"【结果保存】拒绝名单已保存：{REJECTED_FILE.name}（{len(df_rejected)}人）")

        # 无数据提示
        if not admitted and not rejected:
            logger.warning("【结果保存】录取/拒绝名单均为空，未生成文件")

    except PermissionError:
        logger.error(f"【结果保存】无写入权限！请检查目录：{ADMITTED_FILE.parent}")
        raise RuntimeError("保存结果失败：无文件写入权限")
    except Exception as e:
        logger.error(f"【结果保存】处理异常：{str(e)}", exc_info=True)
        raise RuntimeError(f"保存结果失败：{str(e)}")
