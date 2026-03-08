import streamlit as st
import os
import re
from docx import Document
import openpyxl
import pandas as pd
from openpyxl.styles import Alignment, Border, Side, Font, PatternFill
import tempfile
import tkinter as tk
from tkinter import filedialog
import locale

# 设置中文排序（Windows）
try:
    locale.setlocale(locale.LC_COLLATE, 'zh_CN.UTF-8')
except:
    try:
        locale.setlocale(locale.LC_COLLATE, 'Chinese_PRC.936')
    except:
        st.warning("⚠️ 无法设置中文排序，将按默认字符排序")

# -------------------------- 工具函数：系统文件夹选择弹窗 --------------------------
def select_folder():
    root = tk.Tk()
    root.withdraw()
    root.wm_attributes('-topmost', 1)
    folder_path = filedialog.askdirectory(
        title="选择报名表所在文件夹",
        initialdir=os.path.expanduser("~")
    )
    root.destroy()
    return folder_path

# -------------------------- 核心函数1：读取名单（支持3类名单） --------------------------
def read_name_list(uploaded_file, list_type):
    """
    读取名单文件
    :param uploaded_file: 上传的文件
    :param list_type: 名单类型（blacklist/newhongji/participate）
    :return: 姓名集合
    """
    name_set = set()
    try:
        if uploaded_file.name.endswith((".xlsx", ".xls")):
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
            lines = uploaded_file.read().decode("utf-8").splitlines()
            name_set = set([line.strip() for line in lines if line.strip()])
        st.success(f"✅ {list_type}读取成功，共{len(name_set)}个姓名")
    except Exception as e:
        st.error(f"❌ {list_type}读取失败：{str(e)}")
    return name_set

# -------------------------- 核心函数2：精准提取（新增报名班级） --------------------------
def extract_form_info(docx_path):
    info = {
        "姓名": "未提取到",
        "学号": "未提取到",
        "年级": "未提取到",
        "联系方式": "未提取到",
        "是否为学生资助对象": "未提取到",
        "申请理由字数": 0,
        "申请理由是否达标(≥100字)": "否",
        "报名班级": "未提取到",  # 新增列
        "是否为黑名单人员": "否",
        "是否为新鸿基对象": "否",
        "本学年是否参加过": "否",
        "是否报名成功": "否"  # 新增判定列
    }

    try:
        doc = Document(docx_path)
        if not doc.tables:
            st.warning(f"⚠️ 文件 {os.path.basename(docx_path)} 中未找到表格！")
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
        
        # 提取报名班级（从文档标题/段落中提取，适配“多媒体软件班”等）
        # 先从段落提取
        for para in doc.paragraphs:
            para_text = para.text.strip()
            if "班" in para_text and "报名" in para_text:
                # 匹配“XXX班”格式
                class_match = re.search(r"([^（）【】\s]+班)", para_text)
                if class_match:
                    info["报名班级"] = class_match.group(1)
                    break
        # 如果段落中没找到，从文件名提取
        if info["报名班级"] == "未提取到":
            file_name = os.path.basename(docx_path)
            class_match = re.search(r"([^（）【】\s]+班)", file_name)
            if class_match:
                info["报名班级"] = class_match.group(1)

    except Exception as e:
        info["姓名"] = f"解析失败：{os.path.basename(docx_path)}"
        st.warning(f"⚠️ 解析文件 {os.path.basename(docx_path)} 失败：{str(e)}")

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
    # 基础条件：本学年未参加过 且 不在黑名单中
    basic_condition = (info["本学年是否参加过"] == "否") and (info["是否为黑名单人员"] == "否")
    
    if not basic_condition:
        return "否"
    
    # 新鸿基对象直接成功
    if info["是否为新鸿基对象"] == "是":
        return "是"
    
    # 非新鸿基对象：字数达标 + 资助对象 → 成功
    if (info["申请理由是否达标(≥100字)"] == "是") and (info["是否为学生资助对象"] == "是"):
        return "是"
    
    # 其余情况失败
    return "否"

# -------------------------- 核心函数4：按班级分组排序 --------------------------
def sort_by_class(all_info):
    """
    按报名班级分组排序：
    1. 先按班级名称中文排序
    2. 同班级内按姓名排序
    3. 未提取到班级的放在最后
    """
    # 分离有班级和无班级的信息
    has_class = [info for info in all_info if info["报名班级"] != "未提取到"]
    no_class = [info for info in all_info if info["报名班级"] == "未提取到"]
    
    # 按班级排序（中文），同班级按姓名排序
    try:
        # 按班级名称排序
        has_class_sorted = sorted(
            has_class,
            key=lambda x: (locale.strxfrm(x["报名班级"]), locale.strxfrm(x["姓名"]))
        )
    except:
        # 备用排序（按字符编码）
        has_class_sorted = sorted(
            has_class,
            key=lambda x: (x["报名班级"], x["姓名"])
        )
    
    # 无班级的按姓名排序
    no_class_sorted = sorted(no_class, key=lambda x: x["姓名"])
    
    # 合并结果（有班级的在前，无班级的在后）
    sorted_info = has_class_sorted + no_class_sorted
    
    return sorted_info

# -------------------------- 核心函数5：生成目标格式Excel（按班级分组） --------------------------
def generate_excel(all_info, save_path):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "报名表信息汇总"

    # 目标表头
    headers = [
        "姓名", "学号", "年级", "联系方式", "是否为学生资助对象",
        "申请理由字数", "申请理由是否达标(≥100字)", "报名班级", 
        "是否为黑名单人员", "是否为新鸿基对象", "本学年是否参加过",
        "是否报名成功"
    ]
    ws.append(headers)

    # 样式配置
    thin_border = Border(
        left=Side(style='thin', color='000000'),
        right=Side(style='thin', color='000000'),
        top=Side(style='thin', color='000000'),
        bottom=Side(style='thin', color='000000')
    )
    header_font = Font(bold=True, size=11)
    header_fill = PatternFill(start_color='E6E6FA', end_color='E6E6FA', fill_type='solid')
    header_align = Alignment(horizontal='center', vertical='center', wrap_text=True)
    data_align = Alignment(horizontal='center', vertical='center')
    number_align = Alignment(horizontal='right', vertical='center')
    # 班级标题行样式
    class_title_fill = PatternFill(start_color='D3D3D3', end_color='D3D3D3', fill_type='solid')
    
    # 应用表头样式
    for cell in ws[1]:
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = header_align
        cell.border = thin_border

    # 写入数据（按班级分组，添加班级标题行）
    row_idx = 2
    current_class = None
    
    for info in all_info:
        # 如果班级变化，添加班级标题行
        if info["报名班级"] != current_class:
            current_class = info["报名班级"]
            # 写入班级标题
            title_cell = ws.cell(row=row_idx, column=1, value=f"【{current_class if current_class != '未提取到' else '未识别班级'}】")
            title_cell.font = Font(bold=True, size=12)
            title_cell.fill = class_title_fill
            title_cell.alignment = Alignment(horizontal='left', vertical='center')
            # 合并标题行（跨12列）
            ws.merge_cells(start_row=row_idx, start_column=1, end_row=row_idx, end_column=12)
            row_idx += 1
        
        # 写入学生数据
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
            if col_idx == 6:  # 申请理由字数列右对齐
                cell.alignment = number_align
            else:
                cell.alignment = data_align
            # 给“是否报名成功”列标色
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

    wb.save(save_path)

# -------------------------- 核心函数6：批量解析（新增排序） --------------------------
def batch_extract(folder_path, blacklist, newhongji_list, participate_list):
    all_info = []
    docx_count = 0

    for root, _, files in os.walk(folder_path):
        for file in files:
            if file.lower().endswith(".docx"):
                docx_count += 1
                file_path = os.path.join(root, file)
                info = extract_form_info(file_path)
                # 匹配名单
                info["是否为黑名单人员"] = "是" if info["姓名"].strip() in blacklist else "否"
                info["是否为新鸿基对象"] = "是" if info["姓名"].strip() in newhongji_list else "否"
                info["本学年是否参加过"] = "是" if info["姓名"].strip() in participate_list else "否"
                # 判定报名成功
                info["是否报名成功"] = judge_enroll_success(info)
                all_info.append(info)

    if docx_count == 0:
        st.error("❌ 目标文件夹中未找到任何.docx格式的Word报名表！")
        return None

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
    
    # 验证提取结果
    test_info = next((i for i in sorted_info if i["姓名"] == "肖玉茂"), None)
    if test_info:
        st.success(f"""✅ 测试提取成功：
- 姓名={test_info['姓名']}，报名班级={test_info['报名班级']}
- 是否为新鸿基对象={test_info['是否为新鸿基对象']}，是否报名成功={test_info['是否报名成功']}""")

    return sorted_info

# -------------------------- Streamlit界面 --------------------------
def main():
    st.set_page_config(
        page_title="报名表解析工具",
        page_icon="📋",
        layout="wide",
        initial_sidebar_state="collapsed"
    )
    st.title("📋 浙江大学开源课堂报名表解析工具")
    st.divider()
    st.markdown("### 📌 功能更新")
    st.markdown("- ✅ 按**报名班级分组**，相同班级学生集中显示；")
    st.markdown("- ✅ 按班级名称中文排序，同班级内按姓名排序；")
    st.markdown("- ✅ Excel中添加班级标题行（灰色背景），便于查看；")
    st.markdown("- ✅ 新增各班级报名人数/成功人数统计。")
    st.divider()

    if "selected_folder" not in st.session_state:
        st.session_state.selected_folder = ""

    col_operate, col_result = st.columns([1, 2.5])
    with col_operate:
        st.header("⚙️ 操作区")
        # 文件夹选择
        if st.button("📁 点击选择报名表文件夹", use_container_width=True, type="secondary"):
            selected_path = select_folder()
            if selected_path:
                st.session_state.selected_folder = selected_path
                st.rerun()
        folder_path = st.text_input(
            "📂 报名表文件夹路径",
            value=st.session_state.selected_folder,
            placeholder="例如：C:\\Users\\91784\\Desktop\\报名表"
        )
        # 上传各类名单
        blacklist_file = st.file_uploader(
            "🚫 上传黑名单文件（xlsx/txt）",
            type=["xlsx", "xls", "txt"],
            help="Excel需含【姓名】列，TXT每行1个姓名"
        )
        newhongji_file = st.file_uploader(
            "🏢 上传新鸿基名单文件（xlsx/txt）",
            type=["xlsx", "xls", "txt"],
            help="Excel需含【姓名】列，TXT每行1个姓名"
        )
        participate_file = st.file_uploader(
            "📝 上传本学年参加名单文件（xlsx/txt）",
            type=["xlsx", "xls", "txt"],
            help="Excel需含【姓名】列，TXT每行1个姓名"
        )
        # 解析按钮
        extract_btn = st.button("🚀 开始批量解析（按班级分组）", use_container_width=True, type="primary")

    with col_result:
        st.header("📊 结果区")
        if extract_btn:
            # 输入校验
            if not folder_path or not os.path.exists(folder_path):
                st.error("❌ 请选择有效的文件夹路径！")
                st.stop()
            if not blacklist_file:
                st.error("❌ 请上传黑名单文件！")
                st.stop()
            if not newhongji_file:
                st.error("❌ 请上传新鸿基名单文件！")
                st.stop()
            if not participate_file:
                st.error("❌ 请上传本学年参加名单文件！")
                st.stop()

            with st.spinner("🔄 正在读取名单+解析报名表+按班级排序..."):
                # 读取三个名单
                blacklist = read_name_list(blacklist_file, "黑名单")
                newhongji_list = read_name_list(newhongji_file, "新鸿基名单")
                participate_list = read_name_list(participate_file, "本学年参加名单")
                # 批量解析+排序
                sorted_info = batch_extract(folder_path, blacklist, newhongji_list, participate_list)
                if not sorted_info:
                    st.stop()

                # 预览结果（按班级分组）
                st.subheader("📈 解析结果预览（按班级分组）")
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

                # 下载Excel
                with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp_excel:
                    generate_excel(sorted_info, tmp_excel.name)
                    st.download_button(
                        label="📥 下载按班级分组的Excel",
                        data=open(tmp_excel.name, "rb"),
                        file_name="报名表信息汇总（按班级分组）.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True,
                        type="primary"
                    )
                try:
                    os.unlink(tmp_excel.name)
                except:
                    pass

if __name__ == "__main__":
    main()