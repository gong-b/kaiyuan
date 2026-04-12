import pandas as pd
from typing import Any, Set
from pathlib import Path
import logging

# 初始化日志
logger = logging.getLogger(__name__)

def read_student_list(file_path: str | Path) -> Set[str]:
    """
    读取学生名单，返回学号集合（增强容错，精准捕获Excel相关异常）
    :param file_path: Excel文件路径
    :return: 有效学号集合
    """
    student_ids = set()
    try:
        # 路径标准化
        file_path = Path(file_path)
        
        # 前置校验：文件存在性+非空
        if not file_path.exists():
            logger.warning(f"Excel文件不存在: {file_path}")
            return student_ids
        if file_path.stat().st_size == 0:
            logger.warning(f"Excel文件为空（大小0字节）: {file_path}")
            return student_ids
        
        # 尝试打开Excel文件（捕获特定异常）
        try:
            xl_file = pd.ExcelFile(file_path, engine="openpyxl")
        except pd.errors.EmptyDataError:
            logger.warning(f"Excel文件无任何数据: {file_path}")
            return student_ids
        except ImportError:
            logger.error(f"缺少openpyxl依赖！请执行: pip install openpyxl")
            return student_ids
        except Exception as e:
            logger.error(f"打开Excel文件失败（可能文件损坏/格式错误）: {file_path} - {str(e)}", exc_info=True)
            return student_ids
        
        # 自动识别包含"学号"的Sheet（遍历所有Sheet）
        target_sheet = None
        for sheet_name in xl_file.sheet_names:
            try:
                # 仅读取前10行检测列名，提升效率
                df_temp = pd.read_excel(xl_file, sheet_name=sheet_name, nrows=10)
                # 匹配包含"学号"的列（不区分大小写/全半角）
                if any("学号" in str(col).lower().replace(" ", "") for col in df_temp.columns):
                    target_sheet = sheet_name
                    logger.debug(f"找到包含学号列的Sheet: {sheet_name}")
                    break
            except Exception as e:
                logger.debug(f"检查Sheet [{sheet_name}] 失败（跳过）: {str(e)}")
                continue
        
        if not target_sheet:
            logger.warning(f"文件{file_path}中未找到包含'学号'列的Sheet")
            return student_ids
        
        # 读取目标Sheet数据（捕获解析异常）
        try:
            df = pd.read_excel(
                xl_file,
                sheet_name=target_sheet,
                engine="openpyxl",
                dtype=str  # 强制所有列按字符串读取，避免学号被解析为数字
            )
        except Exception as e:
            logger.error(f"读取Sheet [{target_sheet}] 失败: {file_path} - {str(e)}", exc_info=True)
            return student_ids
        
        # 查找学号列（兼容不同列名：学号、学生学号、ID等）
        id_cols = [col for col in df.columns if "学号" in str(col).lower().replace(" ", "")]
        if not id_cols:
            logger.warning(f"Sheet [{target_sheet}] 中未找到'学号'相关列")
            return student_ids
        
        # 提取并清洗学号（去重、去空、去非数字）
        id_series = df[id_cols[0]].astype(str).dropna().str.strip()
        # 过滤规则：非空 + 不是"nan" + 纯数字
        student_ids = set(
            sid for sid in id_series 
            if sid and sid != "nan" and sid.replace(" ", "").isdigit()
        )
        
        logger.info(f"读取学生名单成功: {file_path} - Sheet[{target_sheet}] - 有效学号数: {len(student_ids)}")
        return student_ids
        
    except Exception as e:
        logger.error(f"读取学生名单未预期异常: {file_path} - {str(e)}", exc_info=True)
        return student_ids

def save_results(admitted: list[dict[str, Any]], rejected: list[dict[str, Any]]):
    """
    保存录取/拒绝结果（增强容错，确保文件可写、列完整）
    :param admitted: 录取名单列表
    :param rejected: 拒绝名单列表
    """
    try:
        from config import ADMITTED_FILE, REJECTED_FILE
        
        # 确保输出目录存在
        for file_path in [ADMITTED_FILE, REJECTED_FILE]:
            file_path.parent.mkdir(exist_ok=True, parents=True)
        
        # ---------------------- 保存录取名单 ----------------------
        if admitted:
            # 转换为DataFrame，确保核心列存在
            df_admitted = pd.DataFrame(admitted)
            # 补全缺失列（避免KeyError）
            for col in ["学号", "姓名"]:
                if col not in df_admitted.columns:
                    df_admitted[col] = ""
            # 补充备注列（可选）
            if "备注" not in df_admitted.columns:
                df_admitted["备注"] = ""
            # 列顺序标准化
            df_admitted = df_admitted[["学号", "姓名", "备注"]]
            
            # 去重（按学号，保留第一条）
            df_admitted = df_admitted.drop_duplicates(subset=["学号"], keep="first")
            
            # 保存文件（覆盖已有文件，编码UTF-8）
            try:
                df_admitted.to_excel(
                    ADMITTED_FILE,
                    sheet_name="录取名单",
                    index=False,
                    engine="openpyxl",
                    encoding="utf-8"
                )
                logger.info(f"录取名单已保存: {ADMITTED_FILE} - 共{len(df_admitted)}条记录")
            except PermissionError:
                logger.error(f"保存录取名单失败：无写入权限 - {ADMITTED_FILE}")
                raise RuntimeError(f"无法写入文件（权限不足）: {ADMITTED_FILE}")
            except Exception as e:
                logger.error(f"保存录取名单失败: {ADMITTED_FILE} - {str(e)}", exc_info=True)
                raise RuntimeError(f"保存录取名单失败: {str(e)}")
        
        # ---------------------- 保存拒绝名单 ----------------------
        if rejected:
            # 转换为DataFrame，确保核心列存在
            df_rejected = pd.DataFrame(rejected)
            # 补全缺失列
            for col in ["学号", "姓名", "原因"]:
                if col not in df_rejected.columns:
                    df_rejected[col] = ""
            # 补充原主题列（可选）
            if "原主题" not in df_rejected.columns:
                df_rejected["原主题"] = ""
            # 列顺序标准化
            df_rejected = df_rejected[["学号", "姓名", "原主题", "原因"]]
            
            # 去重（按学号+原因，保留第一条）
            df_rejected = df_rejected.drop_duplicates(subset=["学号", "原因"], keep="first")
            
            # 保存文件
            try:
                df_rejected.to_excel(
                    REJECTED_FILE,
                    sheet_name="拒绝名单",
                    index=False,
                    engine="openpyxl",
                    encoding="utf-8"
                )
                logger.info(f"拒绝名单已保存: {REJECTED_FILE} - 共{len(df_rejected)}条记录")
            except PermissionError:
                logger.error(f"保存拒绝名单失败：无写入权限 - {REJECTED_FILE}")
                raise RuntimeError(f"无法写入文件（权限不足）: {REJECTED_FILE}")
            except Exception as e:
                logger.error(f"保存拒绝名单失败: {REJECTED_FILE} - {str(e)}", exc_info=True)
                raise RuntimeError(f"保存拒绝名单失败: {str(e)}")
        
        # 无数据时的提示
        if not admitted and not rejected:
            logger.warning("录取/拒绝名单均为空，未生成任何Excel文件")
        
    except Exception as e:
        logger.error(f"保存结果流程未预期异常: {str(e)}", exc_info=True)
        raise
