import streamlit as st
import os
import re
from docx import Document
import openpyxl
import pandas as pd
from openpyxl.styles import Alignment, Border, Side, Font, PatternFill
import tempfile
import locale

# ========== 核心修改：移除tkinter，改用Streamlit原生文件上传 ==========
# 设置中文排序（Windows/Linux兼容）
try:
    locale.setlocale(locale.LC_COLLATE, 'zh_CN.UTF-8')
except:
    try:
        locale.setlocale(locale.LC_COLLATE, 'Chinese_PRC.936')
    except:
        st.warning("⚠️ 无法设置中文排序，将按默认字符排序")

# -------------------------- 核心函数1：读取名单（支持3类名单） --------------------------
def read_name_list(uploaded_file, list_type):
    """
    读取名单文件（适配网页版上传）
    :param uploaded_file: Streamlit上传的文件对象
    :param list_type: 名单类型（blacklist/newhongji/participate）
    :return: 姓名集合
    """
    name_set = set()
    if uploaded_file is None:
        st.error(f"❌ 未上传{list_type}文件！")
        return name_set
    
    try:
        if uploaded_file.name.endswith((".xlsx", ".xls")):
            # 读取Excel文件（网页版适配）
            df = pd.read_excel(uploaded_file)
            # 优先匹配“姓名”列，兼容“学生姓名”等列名
            name_col = None
            for col in df.columns:
                if "姓名" in col:
                    name_col = col
                    break
            if not name_col:
                st.error(f"❌ {list_type}文件必须包含【姓名】列！")
                return name_set
            name_set = set(df[name_col].dropna().astype(str).str.strip())
        elif uploaded_file.name.endswith(".txt"):
            # 读取TXT文件（网页版适配）
            lines = uploaded_file.read().decode("utf-8").splitlines()
            name_set = set([line.strip() for line in lines if line.strip()])
        st.success(f"✅ {list_type}读取成功，共{len(name_set)}个姓名")
    except Exception as e:
        st.error(f"❌ {list_type}读取失败：{str(e)}")
    return name_set

# -------------------------- 核心函数2：从doc对象提取信息（移除文件路径依赖） --------------------------
def extract_form_info_from_doc(doc, file_name):
    """
    从docx对象提取信息（适配网页版上传的文件）
    :param doc: Document对象
    :param file_name: 上传的文件名
    :return: 提取的信息字典
    """
    info = {
        "姓名": "未提取到",
        "学号": "未提取到",
        "年级": "未提取到",
        "联系方式": "未提取到",
        "是否为学生资助对象": "未提取到",
        "申请理由字数": 0,
        "申请理由是否达标(≥100字)": "否",
        "报名班级": "未提取到",
        "是否为黑名单人员": "否",
        "是否为新鸿基对象": "否",
        "本学年是否参加过": "否",
        "是否报名成功": "否"
    }

    try:
        if not doc.tables:
            st.warning(f"⚠️ 文件 {file_name} 中未找到表格！")
            return info
        
        table = doc.tables[0]
        
        # 提取姓名（表格行0，单元格1）
        if len(table.rows) > 0 and len(table.rows[0].cells) > 1:
            name_cell = table.rows[0].cells[1]
            name_text = name_cell.text.strip()
            if name_text:
                info["姓名"] = name_text
        
        # 提取学号（表格行0，单元格3）
        if len(table.rows) > 0 and len(table.rows[0].cells) > 3:
            id_cell = table.rows[0].cells[3]
            id_text = id_cell.text.strip()
            if id_text:
                info["学号"] = id_text
        
        # 提取年级（表格行1，单元格3）
        if len(table.rows) > 1 and len(table.rows[1].cells) > 3:
            grade_cell = table.rows[1].cells[3]
            grade_text = grade_cell.text.strip()
            if grade_text:
                info["年级"] = grade_text
        
        # 提取联系方式（表格行3，单元格3）
        if len(table.rows) > 3 and len(table.rows[3].cells) > 3:
            phone_cell = table.rows[3].cells[3]
            phone_text = phone_cell.text.strip()
            if phone_text:
                info["联系方式"] = phone_text
        
        # 提取是否为学生资助对象（表格行5，单元格1）
        if len(table.rows) > 5 and len(table.rows[5].cells) > 1:
            subsidy_cell = table.rows[5].cells[1]
            subsidy_text = subsidy_cell.text.strip()
            if subsidy_text:
                info["是否为学生资助对象"] = subsidy_text
        
        # 提取申请理由并统计字数（表格行7，单元格0）
        if len(table.rows) > 7 and len(table.rows[7].cells) > 0:
            reason_cell = table.rows[7].cells[0]
            reason_text = reason_cell.text.strip()
            reason_text = re.sub(r"申请理由（不少于100字）：\s*", "", reason_text)
            if reason_text:
                real_chars = re.findall(r"[一-龥a-zA-Z]+", reason_text)
                real_count = len("".join(real_chars))
                info["申请理由字数"] = real_count
                info["申请理由是否达标(≥100字)"] = "是" if real_count >= 100 else "否"
        
        # 提取报名班级（从文档段落/文件名提取）
        # 先从段落提取
        for para in doc.paragraphs:
            para_text = para.text.strip()
            if "班" in para_text and "报名" in para_text:
                class_match = re.search(r"([^（）【】\s]+班)", para_text)
                if class_match:
                    info["报名班级"] = class_match.group(1)
                    break
        # 如果段落中没找到，从文件名提取
        if info["报名班级"] == "未提取到":
            class_match = re.search(r"([^（）【】\s]+班)", file_name)
            if class_match:
                info["报名班级"] = class_match.group(1)

    except Exception as e:
        info["姓名"] = f"解析失败：{file_name}"
        st.warning(f"⚠️ 解析文件 {file_name} 失败：{str(e)}")

    return info

# -------------------------- 核心函数3：报名成功判定逻辑 --------------------------
def judge_enroll_success(info):
    """
    判定是否报名成功：
    条件：本学年未参加过 + 不在黑名单中
    - 新鸿基对象 → 一定成功（无视字数/资助对象）
    - 非新鸿基对象：字数达标 + 资助对象 → 成功
    - 其余 → 失败
    """
    basic_condition = (info["本学年是否参加过"] == "否") and (info["是否为黑名单人员"] == "否")
    
    if not basic_condition:
        return "否"
    if info["是否为新鸿基对象"] == "是":
        return "是"
    if (info["申请理由是否达标(≥100字)"] == "是") and (info["是否为学生资助对象"] == "是"):
        return "是"
    return "否"

# -------------------------- 核心函数4：按班级分组排序 --------------------------
def sort_by_class(all_info):
    """按报名班级分组排序"""
    has_class = [info for info in all_info if info["报名班级"] != "未提取到"]
    no_class = [info for info in all_info if info["报名班级"] == "未提取到"]
    
    # 按班级+姓名排序
    try:
        has_class_sorted = sorted(has_class, key=lambda x: (locale.strxfrm(x["报名班级"]), locale.strxfrm(x["姓名"])))
    except:
        has_class_sorted = sorted(has_class, key=lambda x: (x["报名班级"], x["姓名"]))
    
    no_class_sorted = sorted(no_class, key=lambda x: x["姓名"])
    return has_class_sorted + no_class_sorted

# -------------------------- 核心函数5：生成Excel（适配网页版下载） --------------------------
def generate_excel(all_info):
    """生成Excel文件，返回字节流（适配Streamlit下载）"""
    # 创建临时文件
    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp_file:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "报名表信息汇总"

        # 表头
        headers = [
            "姓名", "学号", "年级", "联系方式", "是否为学生资助对象",
            "申请理由字数", "申请理由是否达标(≥100字)", "报名班级", 
            "是否为黑名单人员", "是否为新鸿基对象", "本学年是否参加过",
            "是否报名成功"
        ]
        ws.append(headers)

        # 样式配置
        thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), 
                            top=Side(style='thin'), bottom=Side(style='thin'))
        header_font = Font(bold=True, size=11)
        header_fill = PatternFill(start_color='E6E6FA', end_color='E6E6FA', fill_type='solid')
        header_align = Alignment(horizontal='center', vertical='center', wrap_text=True)
        data_align = Alignment(horizontal='center', vertical='center')
        number_align = Alignment(horizontal='right', vertical='center')
        class_title_fill = PatternFill(start_color='D3D3D3', end_color='D3D3D3', fill_type='solid')
        
        # 应用表头样式
        for cell in ws[1]:
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = header_align
            cell.border = thin_border

        # 写入数据
        row_idx = 2
        current_class = None
        for info in all_info:
            # 班级标题行
            if info["报名班级"] != current_class:
                current_class = info["报名班级"]
                title_cell = ws.cell(row=row_idx, column=1, value=f"【{current_class if current_class != '未提取到' else '未识别班级'}】")
                title_cell.font = Font(bold=True, size=12)
                title_cell.fill = class_title_fill
                title_cell.alignment = Alignment(horizontal='left', vertical='center')
                ws.merge_cells(start_row=row_idx, start_column=1, end_row=row_idx, end_column=12)
                row_idx += 1
            
            # 学生数据行
            row_data = [
                info["姓名"], info["学号"], info["年级"],
                info["联系方式"], info["是否为学生资助对象"],
                info["申请理由字数"], info["申请理由是否达标(≥100字)"],
                info["报名班级"], info["是否为黑名单人员"],
                info["是否为新鸿基对象"], info["本学年是否参加过"],
                info["是否报名成功"]
            ]
            for col_idx, value in enumerate(row_data, start=1):
                cell = ws.cell(row=row_idx, column=col_idx, value=value)
                cell.border = thin_border
                if col_idx == 6:
                    cell.alignment = number_align
                else:
                    cell.alignment = data_align
                # 报名成功标色
                if col_idx == 12:
                    if value == "是":
                        cell.fill = PatternFill(start_color='90EE90', end_color='90EE90', fill_type='solid')
                    else:
                        cell.fill = PatternFill(start_color='FFB6C1', end_color='FFB6C1', fill_type='solid')
            row_idx += 1

        # 调整列宽
        column_widths = [12, 18, 8, 15, 20, 12, 20, 15, 15, 15, 18, 15]
        for col_idx, width in enumerate(column_widths, start=1):
            col_letter = openpyxl.utils.get_column_letter(col_idx)
            ws.column_dimensions[col_letter].width = width

        # 冻结表头
        ws.freeze_panes = "A2"
        
        # 保存到临时文件
        wb.save(tmp_file.name)
        
        # 读取文件字节流
        with open(tmp_file.name, "rb") as f:
            excel_data = f.read()
        
        # 删除临时文件
        os.unlink(tmp_file.name)
        
        return excel_data

# -------------------------- 核心函数6：批量解析（适配网页版上传） --------------------------
def batch_extract(uploaded_docx_files, blacklist, newhongji_list, participate_list):
    """批量解析上传的docx文件"""
    all_info = []
    docx_count = len(uploaded_docx_files)

    if docx_count == 0:
        st.error("❌ 未上传任何.docx格式的Word报名表！")
        return None

    # 解析每个上传的文件
    for file in uploaded_docx_files:
        doc = Document(file)
        info = extract_form_info_from_doc(doc, file.name)
        # 匹配名单
        info["是否为黑名单人员"] = "是" if info["姓名"].strip() in blacklist else "否"
        info["是否为新鸿基对象"] = "是" if info["姓名"].strip() in newhongji_list else "否"
        info["本学年是否参加过"] = "是" if info["姓名"].strip() in participate_list else "否"
        # 判定报名成功
        info["是否报名成功"] = judge_enroll_success(info)
        all_info.append(info)

    # 按班级排序
    sorted_info = sort_by_class(all_info)
    
    st.success(f"✅ 批量解析完成，共处理 {docx_count} 个Word报名表（已按班级分组）")
    
    # 统计班级和人数
    class_stats = {}
    for info in sorted_info:
        cls = info["报名班级"]
        if cls not in class_stats:
            class_stats[cls] = {"total": 0, "success": 0}
        class_stats[cls]["total"] += 1
        if info["是否报名成功"] == "是":
            class_stats[cls]["success"] += 1
    
    # 显示班级统计
    st.info("📊 各班级报名统计：")
    for cls, stats in class_stats.items():
        st.write(f"• {cls}：总计{stats['total']}人，成功{stats['success']}人，失败{stats['total']-stats['success']}人")
    
    return sorted_info

# -------------------------- Streamlit网页界面 --------------------------
def main():
    # 页面配置
    st.set_page_config(
        page_title="浙江大学开源课堂报名表解析工具",
        page_icon="📋",
        layout="wide",
        initial_sidebar_state="collapsed"
    )
    
    # 页面标题
    st.title("📋 浙江大学开源课堂报名表解析工具")
    st.divider()
    
    # 分栏布局
    col_operate, col_result = st.columns([1, 2.5])
    
    with col_operate:
        st.header("⚙️ 操作区")
        
        # 1. 批量上传报名表docx文件（核心修改：替换tkinter文件夹选择）
        uploaded_docx_files = st.file_uploader(
            "📁 批量上传报名表docx文件（可多选）",
            type=["docx"],
            accept_multiple_files=True,
            help="按住Ctrl键可多选文件，支持批量上传"
        )
        
        # 2. 上传各类名单文件
        st.subheader("📄 上传对照名单")
        blacklist_file = st.file_uploader("🚫 黑名单文件（xlsx/txt）", type=["xlsx", "xls", "txt"])
        newhongji_file = st.file_uploader("🏢 新鸿基名单文件（xlsx/txt）", type=["xlsx", "xls", "txt"])
        participate_file = st.file_uploader("📝 本学年参加名单文件（xlsx/txt）", type=["xlsx", "xls", "txt"])
        
        # 3. 解析按钮
        if st.button("🚀 开始批量解析（按班级分组）", type="primary"):
            # 读取名单
            blacklist = read_name_list(blacklist_file, "黑名单")
            newhongji_list = read_name_list(newhongji_file, "新鸿基名单")
            participate_list = read_name_list(participate_file, "本学年参加名单")
            
            # 批量解析
            if blacklist and newhongji_list and participate_list:
                sorted_info = batch_extract(uploaded_docx_files, blacklist, newhongji_list, participate_list)
                
                if sorted_info:
                    # 生成Excel
                    excel_data = generate_excel(sorted_info)
                    
                    # 显示结果预览
                    with col_result:
                        st.header("📊 结果预览")
                        result_df = pd.DataFrame(sorted_info)
                        st.dataframe(
                            result_df,
                            use_container_width=True,
                            hide_index=True,
                            column_config={
                                "申请理由字数": st.column_config.NumberColumn(format="%d"),
                                "申请理由是否达标(≥100字)": st.column_config.SelectboxColumn(options=["是", "否"]),
                                "是否为黑名单人员": st.column_config.SelectboxColumn(options=["是", "否"]),
                                "是否为新鸿基对象": st.column_config.SelectboxColumn(options=["是", "否"]),
                                "本学年是否参加过": st.column_config.SelectboxColumn(options=["是", "否"]),
                                "是否报名成功": st.column_config.SelectboxColumn(options=["是", "否"])
                            }
                        )
                    
                    # 下载按钮
                    st.download_button(
                        label="📥 下载按班级分组的Excel",
                        data=excel_data,
                        file_name="报名表信息汇总（按班级分组）.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )

    # 页面说明
    st.divider()
    st.markdown("### 📌 使用说明")
    st.markdown("1. 批量上传报名表docx文件（可多选）；")
    st.markdown("2. 上传黑名单、新鸿基名单、本学年参加名单（Excel/TXT格式，需含姓名列）；")
    st.markdown("3. 点击解析按钮，等待完成后下载Excel文件；")
    st.markdown("4. Excel文件按班级分组，报名成功列标色（绿色=成功，红色=失败）。")

if __name__ == "__main__":
    main()
