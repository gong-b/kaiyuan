import pandas as pd
from typing import Any, Set
from pathlib import Path
import logging

logger = logging.getLogger(__name__)

def read_student_list(file_path: str | Path) -> Set[str]:
    """读取学生名单，返回学号集合（增强容错）"""
    student_ids = set()
    try:
        file_path = Path(file_path)
        if not file_path.exists():
            logger.warning(f"文件不存在: {file_path}")
            return student_ids
        
        # 读取所有sheet，自动识别包含学号的sheet
        xl_file = pd.ExcelFile(file_path, engine="openpyxl")
        target_sheet = None
        
        for sheet_name in xl_file.sheet_names:
            try:
                df_temp = pd.read_excel(xl_file, sheet_name=sheet_name, nrows=10)  # 只读取前10行检测列
                if any("学号" in str(col).lower() for col in df_temp.columns):
                    target_sheet = sheet_name
                    break
            except:
                continue
        
        if not target_sheet:
            logger.warning(f"文件{file_path}无学号列")
            return student_ids
        
        # 读取目标sheet
        df = pd.read_excel(
            file_path,
            sheet_name=target_sheet,
            engine="openpyxl"
        )
        
        # 查找学号列（不区分大小写）
        id_cols = [col for col in df.columns if "学号" in str(col).lower()]
        if not id_cols:
            logger.warning(f"文件{file_path}的sheet {target_sheet} 无学号列")
            return student_ids
        
        # 提取并清洗学号
        id_series = df[id_cols[0]].astype(str).dropna().str.strip()
        # 过滤无效值
        student_ids = set(
            sid for sid in id_series 
            if sid and sid != "nan" and sid.isdigit()
        )
        
        logger.info(f"读取学生名单成功: {file_path} - 共{len(student_ids)}个有效学号")
        return student_ids
        
    except Exception as e:
        logger.error(f"读取学生名单失败: {file_path} - {str(e)}")
        return student_ids

def save_results(admitted: list[dict[str, Any]], rejected: list[dict[str, Any]]):
    """保存录取/拒绝结果（修复Excel写入，增强兼容性）"""
    try:
        from config import ADMITTED_FILE, REJECTED_FILE
        
        # 确保目录存在
        ADMITTED_FILE.parent.mkdir(exist_ok=True, parents=True)
        
        # 保存录取名单
        if admitted:
            df_admitted = pd.DataFrame(admitted)
            # 确保必要列存在
            for col in ["学号", "姓名"]:
                if col not in df_admitted.columns:
                    df_admitted[col] = ""
            # 列顺序
            cols = ["学号", "姓名", "备注"] if "备注" in df_admitted.columns else ["学号", "姓名"]
            df_admitted = df_admitted[cols]
            
            # 去重（按学号）
            df_admitted = df_admitted.drop_duplicates(subset=["学号"], keep="first")
            
            # 保存（覆盖）
            df_admitted.to_excel(
                ADMITTED_FILE, 
                sheet_name="录取名单", 
                index=False, 
                engine="openpyxl",
                encoding="utf-8"
            )
            logger.info(f"录取名单已保存: {ADMITTED_FILE}")
        
        # 保存拒绝名单
        if rejected:
            df_rejected = pd.DataFrame(rejected)
            # 确保必要列存在
            for col in ["学号", "姓名", "原因"]:
                if col not in df_rejected.columns:
                    df_rejected[col] = ""
            # 列顺序
            cols = ["学号", "姓名", "原主题", "原因"] if "原主题" in df_rejected.columns else ["学号", "姓名", "原因"]
            df_rejected = df_rejected[cols]
            
            # 去重
            df_rejected = df_rejected.drop_duplicates(subset=["学号", "原因"], keep="first")
            
            # 保存（覆盖）
            df_rejected.to_excel(
                REJECTED_FILE, 
                sheet_name="拒绝名单", 
                index=False, 
                engine="openpyxl",
                encoding="utf-8"
            )
            logger.info(f"拒绝名单已保存: {REJECTED_FILE}")
        
        logger.info(f"结果保存成功 - 录取:{len(admitted)}人, 拒绝:{len(rejected)}人")
        
    except Exception as e:
        logger.error(f"保存结果失败: {str(e)}", exc_info=True)
        raise
