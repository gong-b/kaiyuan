import streamlit as st
import os
import re
from docx import Document
import openpyxl
import pandas as pd
from openpyxl.styles import Alignment, Border, Side, Font, PatternFill
import tempfile
import locale
import zipfile
import io
import time

# ========== 兼容处理：psutil 可选导入 ==========
尝试:
    导入psutil
    PSUTIL_AVAILABLE = True
除导入错误：
    PSUTIL_AVAILABLE = False
st.警告(
        "⚠️ psutil库未安装，文件占用检测功能受限\n"
        本地运行可执行: pip install psutil 
        "云部署无需处理，程序会自动降级运行"
    )

# ========== Streamlit 页面配置（必须放在最前面） ==========
st.set_page_config(
    page_title="浙江大学开源课堂报名表解析工具",
    page_icon="📋",
    layout="wide",
    initial_sidebar_state="collapsed",
    menu_items={
        “关于”: """
        ### 浙江大学开源课堂报名表解析工具
        版本：1.0.0
        功能：批量解析docx报名表，生成结构化Excel
部署：Streamlit社区云
        """
    }
)

# ========== 深色主题样式 ==========
def set_dark_theme():
    """设置深色主题样式"""
    st.markdown(
        """
        <style>
        /* 整体背景 */
        .stApp {
            background-color: #121212 !important;
            color: #e0e0e0 !important;
        }
        /* 文本颜色 */
        .stMarkdown, .stText, .stHeader, .stSubheader, 
        .stSuccess, .stInfo, .stWarning, .stError {
            color: #e0e0e0 !important;
        }
        /* 按钮样式 */
        .stButton>button {
            background-color: #2d2d2d;
            color: #ffffff;
            border: 1px solid #4CAF50;
            border-radius: 5px;
            transition: all 0.3s ease;
        }
        .stButton>button:hover {
            background-color: #4CAF50;
            color: white;
        }
        /* 文件上传框 */
        .stFileUploader {
            padding: 10px;
            border: 1px solid #3d3d3d;
            border-radius: 5px;
            background-color: #1e1e1e;
        }
        /* 数据表格 */
        .stDataFrame {
            background-color: #1e1e1e;
            color: #e0e0e0;
        }
        /* 进度条 */
        .stProgress > div > div {
            background-color: #4CAF50 !important;
        }
        /* 展开栏 */
        .stExpander {
            background-color: #1e1e1e;
            border: 1px solid #3d3d3d;
        }
        /* 下载按钮 */
        .stDownloadButton>button {
            background-color: #2d2d2d;
            color: #ffffff;
            border: 1px solid #2196F3;
            border-radius: 5px;
            transition: all 0.3s ease;
        }
        .stDownloadButton>button:hover {
            background-color: #2196F3;
        }
        /* 列布局 */
        [data-testid="column"] {
            background-color: #1e1e1e;
            padding: 10px;
            border-radius: 8px;
            margin: 5px;
        }
        /* 隐藏Streamlit默认页脚 */
        footer {
            visibility: hidden;
        }
        /* 自定义页脚 */
        .custom-footer {
            color: #888888;
            text-align: center;
            padding: 20px 0;
            font-size: 0.9em;
        }
        </style>
        """,
        unsafe_allow_html=True
    )

# 应用深色主题
set_dark_theme()

# ========== 工具函数 ==========
def safe_remove_file(file_path, max_retries=5, delay=0.5):
    """
    安全删除文件（兼容云部署，无psutil依赖）
    :param file_path: 文件路径
    :param max_retries: 最大重试次数
    :param delay: 重试间隔（秒）
    """
    retries = 0
    while retries < max_retries:
        try:
            if os.path.exists(file_path):
                # 仅本地环境使用psutil释放文件占用
                if PSUTIL_AVAILABLE and os.name == 'nt':
                    try:
                        for proc in psutil.process_iter(['pid', 'open_files']):
                            for open_file in proc.info['open_files'] or []:
                                if file_path.lower() == open_file.path.lower():
                                    proc.terminate()
                                    time.sleep(0.5)
                                    break
                    except Exception:
                        pass
                
                os.unlink(file_path)
                break
        except PermissionError:
            retries += 1
            time.sleep(delay)
    else:
        # 云环境临时文件会自动清理，仅提示
        if not st.secrets.get("IS_CLOUD", False):
            st.warning(f"⚠️ 临时文件 {os.path.basename(file_path)} 暂时无法删除，程序退出后会自动清理")

def is_valid_docx(file_bytes):
    """校验是否为有效的docx文件"""
    try:
        with zipfile.ZipFile(io.BytesIO(file_bytes)) as zf:
            if 'word/document.xml' not in zf.namelist():
                return False
            # 快速校验（避免大文件耗时）
            for zip_info in zf.infolist():
                if zip_info.file_size > 102400:
                    continue
                zf.read(zip_info.filename)
        return True
    except Exception:
        return False

def safe_read_docx(uploaded_file):
    """安全读取上传的docx文件"""
    try:
        file_bytes = uploaded_file.read()
        if not is_valid_docx(file_bytes):
            st.warning(f"⚠️ 文件 {uploaded_file.name} 不是有效的DOCX文件，已跳过")
            return None
        
        uploaded_file.seek(0)
        return Document(io.BytesIO(file_bytes))
    except Exception as e:
        st.warning(f"⚠️ 读取文件 {uploaded_file.name} 失败：{str(e)}，已跳过")
        return None

def init_session_state():
    """初始化会话状态（云部署兼容）"""
    if 'uploaded_files' not in st.session_state:
        st.session_state.uploaded_files = []
    if 'blacklist' not in st.session_state:
        st.session_state.blacklist = set()
    if 'newhongji_list' not in st.session_state:
        st.session_state.newhongji_list = set()
    if 'participate_list' not in st.session_state:
        st.session_state.participate_list = set()
    if 'parsed_result' not in st.session_state:
        st.session_state.parsed_result = None

# ========== 核心业务函数 ==========
def read_name_list(uploaded_file, list_type):
    """读取名单文件（支持Excel/TXT）"""
    name_set = set()
    if uploaded_file is None:
        st.error(f"❌ 未上传{list_type}文件！")
        return name_set
    
    try:
        if uploaded_file.name.endswith((".xlsx", ".xls")):
            df = pd.read_excel(uploaded_file)
            name_col = next((col for col in df.columns if "姓名" in col), None)
            if not name_col:
                st.error(f"❌ {list_type}文件必须包含【姓名】列！")
                return name_set
            name_set = set(df[name_col].dropna().astype(str).str.strip())
        elif uploaded_file.name.endswith(".txt"):
            lines = uploaded_file.read().decode("utf-8").splitlines()
            name_set = set([line.strip() for line in lines if line.strip()])
        
        st.success(f"✅ {list_type}读取成功，共{len(name_set)}个姓名")
        
        # 更新会话状态
        if list_type == "黑名单":
            st.session_state.blacklist = name_set
        elif list_type == "新鸿基名单":
            st.session_state.newhongji_list = name_set
        elif list_type == "本学年参加名单":
            st.session_state.participate_list = name_set
            
    except Exception as e:
        st.error(f"❌ {list_type}读取失败：{str(e)}")
    return name_set

def extract_form_info_from_doc(doc, file_name):
    """从docx文件提取报名表信息"""
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
        "本学年参加过": "否",
        "是否报名成功": "否"
    }

    try:
        if not doc.tables:
            st.warning(f"⚠️ 文件 {file_name} 中未找到表格！")
            return info
        
        table = doc.tables[0]
        
        # 提取核心信息
        extract_rules = [
            ("姓名", 0, 1),
            ("学号", 0, 3),
            ("年级", 1, 3),
            ("联系方式", 3, 3),
            ("是否为学生资助对象", 5, 1)
        ]
        
        for key, row_idx, col_idx in extract_rules:
            if len(table.rows) > row_idx and len(table.rows[row_idx].cells) > col_idx:
                cell_text = table.rows[row_idx].cells[col_idx].text.strip()
                if cell_text:
                    info[key] = cell_text
        
        # 提取申请理由
        if len(table.rows) > 7 and len(table.rows[7].cells) > 0:
            reason_text = table.rows[7].cells[0].text.strip()
            reason_text = re.sub(r"申请理由（不少于100字）：\s*", "", reason_text)
            if reason_text:
                real_chars = re.findall(r"[一-龥a-zA-Z]+", reason_text)
                info["申请理由字数"] = len("".join(real_chars))
                info["申请理由是否达标(≥100字)"] = "是" if info["申请理由字数"] >= 100 else "否"
        
        # 提取报名班级
        for para in doc.paragraphs:
            para_text = para.text.strip()
            class_match = re.search(r"([^（）【】\s]+班)", para_text)
            if class_match and "报名" in para_text:
                info["报名班级"] = class_match.group(1)
                break
        
        if info["报名班级"] == "未提取到":
            class_match = re.search(r"([^（）【】\s]+班)", file_name)
            if class_match:
                info["报名班级"] = class_match.group(1)

    except Exception as e:
        info["姓名"] = f"解析失败：{file_name}"
        st.warning(f"⚠️ 解析文件 {file_name} 失败：{str(e)}")

    return info

def judge_enroll_success(info):
    """判定报名成功状态"""
    basic_condition = (info["本学年参加过"] == "否") and (info["是否为黑名单人员"] == "否")
    
    if not basic_condition:
        return "否"
    if info["是否为新鸿基对象"] == "是":
        return "是"
    if (info["申请理由是否达标(≥100字)"] == "是") and (info["是否为学生资助对象"] == "是"):
        return "是"
    return "否"

def sort_by_class(all_info):
    """按班级和姓名排序"""
    has_class = [info for info in all_info if info["报名班级"] != "未提取到"]
    no_class = [info for info in all_info if info["报名班级"] == "未提取到"]
    
    # 中文排序兼容
    try:
        locale.setlocale(locale.LC_COLLATE, 'zh_CN.UTF-8')
        has_class_sorted = sorted(has_class, key=lambda x: (locale.strxfrm(x["报名班级"]), locale.strxfrm(x["姓名"])))
        no_class_sorted = sorted(no_class, key=lambda x: locale.strxfrm(x["姓名"]))
    except:
        has_class_sorted = sorted(has_class, key=lambda x: (x["报名班级"], x["姓名"]))
        no_class_sorted = sorted(no_class, key=lambda x: x["姓名"])
    
    return has_class_sorted + no_class_sorted

def generate_excel(all_info):
    """生成带样式的Excel文件，返回字节流"""
    try:
        # 使用临时文件（云部署兼容）
        with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp_file:
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "报名表信息汇总"

            # 表头配置
            headers = [
                "姓名", "学号", "年级", "联系方式", "是否为学生资助对象",
                "申请理由字数", "申请理由是否达标(≥100字)", "报名班级", 
                "是否为黑名单人员", "是否为新鸿基对象", "本学年参加过",
                "是否报名成功"
            ]
            ws.append(headers)

            # 样式定义
            thin_border = Border(left=Side(style='thin'), right=Side(style='thin'),
                                top=Side(style='thin'), bottom=Side(style='thin'))
            header_style = {
                "font": Font(bold=True, size=11),
                "fill": PatternFill(start_color='E6E6FA', end_color='E6E6FA', fill_type='solid'),
                "alignment": Alignment(horizontal='center', vertical='center', wrap_text=True),
                "border": thin_border
            }
            data_align = Alignment(horizontal='center', vertical='center')
            number_align = Alignment(horizontal='right', vertical='center')
            class_title_fill = PatternFill(start_color='D3D3D3', end_color='D3D3D3', fill_type='solid')
            
            # 应用表头样式
            for cell in ws[1]:
                for attr, value in header_style.items():
                    setattr(cell, attr, value)

            # 写入数据
            row_idx = 2
            current_class = None
            for info in all_info:
                # 班级标题行
                if info["报名班级"] != current_class:
                    current_class = info["报名班级"]
                    title_cell = ws.cell(row=row_idx, column=1, 
                                       value=f"【{current_class if current_class != '未提取到' else '未识别班级'}】")
                    title_cell.font = Font(bold=True, size=12)
                    title_cell.fill = class_title_fill
                    title_cell.alignment = Alignment(horizontal='left', vertical='center')
                    ws.merge_cells(f"A{row_idx}:L{row_idx}")
                    row_idx += 1
                
                # 学生数据行
                row_data = [
                    info["姓名"], info["学号"], info["年级"],
                    info["联系方式"], info["是否为学生资助对象"],
                    info["申请理由字数"], info["申请理由是否达标(≥100字)"],
                    info["报名班级"], info["是否为黑名单人员"],
                    info["是否为新鸿基对象"], info["本学年参加过"],
                    info["是否报名成功"]
                ]
                
                for col_idx, value in enumerate(row_data, start=1):
                    cell = ws.cell(row=row_idx, column=col_idx, value=value)
                    cell.border = thin_border
                    cell.alignment = number_align if col_idx == 6 else data_align
                    
                    # 报名成功标色
                    if col_idx == 12:
                        color = '90EE90' if value == "是" else 'FFB6C1'
                        cell.fill = PatternFill(start_color=color, end_color=color, fill_type='solid')
                
                row_idx += 1

            # 调整列宽
            column_widths = [12, 18, 8, 15, 20, 12, 20, 15, 15, 15, 18, 15]
            for col_idx, width in enumerate(column_widths, start=1):
                col_letter = openpyxl.utils.get_column_letter(col_idx)
                ws.column_dimensions[col_letter].width = width

            # 冻结表头
            ws.freeze_panes = "A2"
            
            # 保存并读取字节流
            wb.save(tmp_file.name)
            wb.close()
            
            with open(tmp_file.name, "rb") as f:
                excel_data = f.read()
        
        # 清理临时文件
        safe_remove_file(tmp_file.name)
        return excel_data
    
    except Exception as e:
        st.error(f"❌ 生成Excel失败：{str(e)}")
        return None

def batch_extract(uploaded_files, blacklist, newhongji_list, participate_list):
    """批量解析上传的文件"""
    if len(uploaded_files) == 0:
        st.error("❌ 未上传任何有效的.docx格式报名表！")
        return None

    all_info = []
    valid_count = 0
    total_files = len(uploaded_files)
    
    # 进度条
    progress_bar = st.progress(0)
    status_text = st.empty()
    
    for idx, file in enumerate(uploaded_files):
        status_text.text(f"正在解析：{file.name}（{idx+1}/{total_files}）")
        progress_bar.progress((idx+1)/total_files)
        
        doc = safe_read_docx(file)
        if doc is None:
            continue
        
        valid_count += 1
        info = extract_form_info_from_doc(doc, file.name)
        
        # 匹配名单
        name = info["姓名"].strip()
        info["是否为黑名单人员"] = "是" if name in blacklist else "否"
        info["是否为新鸿基对象"] = "是" if name in newhongji_list else "否"
        info["本学年参加过"] = "是" if name in participate_list else "否"
        info["是否报名成功"] = judge_enroll_success(info)
        
        all_info.append(info)

    # 清理进度条
    progress_bar.empty()
    status_text.empty()
    
    # 排序和统计
    sorted_info = sort_by_class(all_info)
    st.success(f"✅ 解析完成！处理 {valid_count}/{total_files} 个有效文件（按班级分组）")
    
    # 班级统计
    class_stats = {}
    for info in sorted_info:
        cls = info["报名班级"]
        class_stats.setdefault(cls, {"total": 0, "success": 0})
        class_stats[cls]["total"] += 1
        if info["是否报名成功"] == "是":
            class_stats[cls]["success"] += 1
    
    # 显示统计
    st.info("📊 各班级报名统计：")
    cols = st.columns(2)
    stats_items = list(class_stats.items())
    for i, (cls, stats) in enumerate(stats_items):
        with cols[i%2]:
            st.write(f"• {cls}：总计{stats['total']}人，成功{stats['success']}人，失败{stats['total']-stats['success']}人")
    
    return sorted_info

# ========== 主界面函数 ==========
def main():
    """主界面逻辑"""
    # 页面标题
    st.title("📋 浙江大学开源课堂报名表解析工具")
    st.divider()
    
    # 初始化会话状态
    init_session_state()
    
    # 布局
    col_operate, col_result = st.columns([1, 2.5])
    
    with col_operate:
        st.header("⚙️ 操作区")
        
        # 1. 上传报名表
        st.subheader("📁 上传报名表（可分批）")
        new_files = st.file_uploader(
            "选择docx文件（支持多选/分批上传）",
            type=["docx"],
            accept_multiple_files=True,
            help="可多次上传，文件会自动累加"
        )
        
        # 累加上传文件
        if new_files:
            for file in new_files:
                if file.name not in [f.name for f in st.session_state.uploaded_files]:
                    st.session_state.uploaded_files.append(file)
            st.success(f"✅ 累计上传：{len(st.session_state.uploaded_files)} 个文件")
        
        # 已上传文件列表
        if st.session_state.uploaded_files:
            with st.expander("📄 已上传文件列表", expanded=False):
                for idx, file in enumerate(st.session_state.uploaded_files):
                    st.write(f"{idx+1}. {file.name}")
            
            if st.button("🗑️ 清空上传文件", type="secondary"):
                st.session_state.uploaded_files = []
                st.rerun()
        
        # 2. 上传对照名单
        st.subheader("📄 上传对照名单")
        blacklist_file = st.file_uploader("🚫 黑名单（xlsx/txt）", type=["xlsx", "xls", "txt"])
        newhongji_file = st.file_uploader("🏢 新鸿基名单（xlsx/txt）", type=["xlsx", "xls", "txt"])
        participate_file = st.file_uploader("📝 本学年参加名单（xlsx/txt）", type=["xlsx", "xls", "txt"])
        
        # 3. 解析按钮
        if st.button("🚀 开始批量解析", type="primary", use_container_width=True):
            # 读取名单（优先使用会话状态）
            if blacklist_file and not st.session_state.blacklist:
                read_name_list(blacklist_file, "黑名单")
            if newhongji_file and not st.session_state.newhongji_list:
                read_name_list(newhongji_file, "新鸿基名单")
            if participate_file and not st.session_state.participate_list:
                read_name_list(participate_file, "本学年参加名单")
            
            # 校验名单
            if not all([st.session_state.blacklist, st.session_state.newhongji_list, st.session_state.participate_list]):
                st.error("❌ 请先上传并读取所有对照名单！")
            else:
                # 批量解析
                sorted_info = batch_extract(
                    st.session_state.uploaded_files,
                    st.session_state.blacklist,
                    st.session_state.newhongji_list,
                    st.session_state.participate_list
                )
                
                if sorted_info:
                    st.session_state.parsed_result = sorted_info
                    # 生成Excel
                    excel_data = generate_excel(sorted_info)
                    
                    if excel_data:
                        # 结果预览
                        with col_result:
                            st.header("📊 解析结果预览")
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
                                    "本学年参加过": st.column_config.SelectboxColumn(options=["是", "否"]),
                                    "是否报名成功": st.column_config.SelectboxColumn(options=["是", "否"])
                                }
                            )
                        
                        # 下载按钮
                        st.download_button(
                            label="📥 下载Excel汇总文件",
                            data=excel_data,
                            file_name="报名表信息汇总_按班级分组.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            use_container_width=True
                        )

    # 使用说明
    st.divider()
    st.markdown("📌 使用说明")
    st.markdown("""
    1. **分批上传**：多次上传docx报名表文件（避免单次上传过多）；
    2. **上传名单**：上传黑名单、新鸿基名单、本学年参加名单（需包含姓名列）；
    3. **开始解析**：点击解析按钮，等待完成后查看统计和预览；
    4. **下载文件**：下载带样式的Excel文件（绿色=报名成功，红色=失败）。
    """)
    
    # 自定义页脚
    st.markdown(
        '<div class="custom-footer">© 2025 浙江大学开源课堂报名表解析工具 | 部署于 Streamlit Community Cloud</div>',
        unsafe_allow_html=True
    )

# ========== 程序入口 ==========
if __name__ == "__main__":
    main()
