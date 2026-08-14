# -*- coding: utf-8 -*-
"""
Created on Tue Jul 28 16:21:28 2026

@author: ULS5SZH
"""
import streamlit as st
import pandas as pd
import docx
import datetime
import io
import os
import zipfile
import re
import xml.etree.ElementTree as ET
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.header import Header
from docx.shared import Inches
import openpyxl

# 页面基本配置
st.set_page_config(page_title="PCB Lesson Learn 模板生成器", layout="wide", page_icon="🌅")

st.markdown("""
# 🌅 PCB Lesson Learn 模板生成器 & 邮件草稿一键生成
本程序已完美匹配 Excel (sheet: **PUQ3 LL Overall list**，表头自第二行开始)。
您只需在下方表格中勾选所需的记录，即可一键自动填充并生成 Word 模板及配套 Outlook 邮件草稿。
""")
st.write("---")

# 1. 默认路径配置（G 盘固定路径）
DEFAULT_EXCEL_PATH = r"G:\02_7_M-PQA-RBAC1\08_PQA_AE\09_PQA2\11_PCB\04_Lessons learn\PCB Lesson Learn Master List.xlsx"
DEFAULT_TEMPLATE_PATH = r"G:\02_7_M-PQA-RBAC1\08_PQA_AE\09_PQA2\11_PCB\04_Lessons learn\LL Template complete version.docx"

# 侧边栏配置面板
st.sidebar.header("⚙️ 配置面板")
use_local_paths = st.sidebar.checkbox("使用本地固定路径 (G 盘)", value=True)

excel_file = None
template_file = None

if use_local_paths:
    excel_path = st.sidebar.text_input("Excel 文件路径:", DEFAULT_EXCEL_PATH)
    template_path = st.sidebar.text_input("Template 模板路径:", DEFAULT_TEMPLATE_PATH)
    
    if os.path.exists(excel_path):
        excel_file = excel_path
    else:
        st.sidebar.warning(f"⚠️ 未在指定路径找到 Excel 文件，请尝试手动上传。")
        
    if os.path.exists(template_path):
        template_file = template_path
    else:
        st.sidebar.warning(f"⚠️ 未在指定路径找到 Word 模板，请尝试手动上传。")
else:
    uploaded_excel = st.sidebar.file_uploader("手动上传 PCB Master List (Excel):", type=["xlsx"])
    uploaded_template = st.sidebar.file_uploader("手动上传 LL Template (Word):", type=["docx"])
    if uploaded_excel:
        excel_file = uploaded_excel
    if uploaded_template:
        template_file = uploaded_template

if excel_file is None:
    uploaded_excel = st.file_uploader("📁 请上传 PCB Lesson Learn Master List (Excel)", type=["xlsx"], key="excel_main")
    if uploaded_excel:
        excel_file = uploaded_excel
if template_file is None:
    uploaded_template = st.file_uploader("📝 请上传 LL Template (Word) 模板文件", type=["docx"], key="template_main")
    if uploaded_template:
        template_file = uploaded_template

# 提取Excel中的所有图片数据
@st.cache_data
def load_excel_images(file_source, sheet_name):
    img_dict = {}
    try:
        wb = openpyxl.load_workbook(file_source, data_only=True)
        ws = wb[sheet_name]
        if hasattr(ws, '_images'):
            for img in ws._images:
                try:
                    # 获取图片所在单元格坐标 (0索引)
                    row = img.anchor._from.row
                    col = img.anchor._from.col
                    img_dict[(row, col)] = img._data()
                except Exception:
                    pass
    except Exception as e:
        print(f"提取图片警告: {e}")
    return img_dict

# 加载Excel数据
def load_excel_robust(file_source):
    xl = pd.ExcelFile(file_source)
    sheet_names = xl.sheet_names
    
    target_sheet = next((s for s in sheet_names if "PUQ3 LL Overall list" in s), None)
    if not target_sheet:
        target_sheet = next((s for s in sheet_names if "Overall" in s or "LL" in s), sheet_names[0])
        
    df_temp = pd.read_excel(file_source, sheet_name=target_sheet, nrows=10, header=None)
    header_idx = 1 
    for idx, row in df_temp.iterrows():
        row_str = [str(val).strip().lower() for val in row.tolist()]
        if any('serials' in val or 'project/part' in val or 'failure mode' in val for val in row_str):
            header_idx = idx
            break
            
    df = pd.read_excel(file_source, sheet_name=target_sheet, header=header_idx)
    df.columns = df.columns.astype(str).str.strip()
    return df, target_sheet, header_idx

# 强力填充Word模板核心逻辑
def fill_word_template(template_source, row_data, img_data=None):
    doc = docx.Document(template_source)
    
    # 自动获取今天日期 (如：August 14, 2026)
    current_date_str = datetime.date.today().strftime('%B %d, %Y')
    
    # 提取需要的字段数据
    failure_mode_val = str(row_data.get('Failure Mode', '')).strip()
    should_or_not = str(row_data.get('Should or not to do', ''))
    
    # 解析 Should 和 Should not (应对多种换行和冒号格式)
    should_text, should_not_text = "", ""
    # 查找 Should: 到 Should not: 之间的内容
    should_match = re.search(r'Should[:：](.*?)(Should not[:：]|$)', should_or_not, re.DOTALL | re.IGNORECASE)
    # 查找 Should not: 之后的所有内容
    should_not_match = re.search(r'Should not[:：](.*)', should_or_not, re.DOTALL | re.IGNORECASE)
    
    if should_match:
        should_text = should_match.group(1).strip()
    if should_not_match:
        should_not_text = should_not_match.group(1).strip()

    content_map = {
        "Task/Scope": str(row_data.get('LL Supplier Scope', '')),
        "Failure Mode": failure_mode_val,
        "Project/Part name": str(row_data.get('Project/Part name', '')),
        "Process": str(row_data.get('Related Material Field / Process', '')),
        "Problem (Fundamental Problem)": str(row_data.get('LL Brief Description', '')),
        "Root Cause": str(row_data.get('Root Cause', '')),
        "Corrective Action": str(row_data.get('Corrective Action', '')),
        "1.1 What should we do in the future?": should_text,
        "1.2 What should we not do in the future?": should_not_text,
    }

    # 遍历Word中所有的段落 (含表格内的段落)
    all_paragraphs = []
    for p in doc.paragraphs:
        all_paragraphs.append(p)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    all_paragraphs.append(p)

    # 1. 替换日期 & 主标题追加
    date_pattern = re.compile(r'[A-Z][a-z]{2,8}\s\d{1,2}(st|nd|rd|th)?,\s20\d{2}') # 匹配常见日期如 May 24, 2022
    for p in all_paragraphs:
        p_text = p.text
        
        # 修复日期
        if date_pattern.search(p_text) or '2022' in p_text or '2023' in p_text:
            for run in p.runs:
                if run.text:
                    run.text = date_pattern.sub(current_date_str, run.text)
                    
        # 修复主标题 (Lesson Learn - Failure Mode)
        if 'Lesson Learn' in p_text and ('–' in p_text or '-' in p_text):
            if failure_mode_val and failure_mode_val not in p_text:
                for i, run in enumerate(p.runs):
                    if '–' in run.text or '-' in run.text:
                        run.text = run.text.rstrip(' -–') + f" – {failure_mode_val}"
                        break
    
    # 2. 定位标题并插入内容
    def set_next_para_text(current_para_index, text_to_insert):
        if not text_to_insert or pd.isna(text_to_insert) or text_to_insert == 'nan':
            return
        if current_para_index + 1 < len(all_paragraphs):
            next_p = all_paragraphs[current_para_index + 1]
            next_p.text = str(text_to_insert)
            # 设置下划线等基础样式
            if next_p.runs:
                next_p.runs[0].font.name = 'Arial'

    for i, p in enumerate(all_paragraphs):
        p_clean = p.text.strip()
        for key, val in content_map.items():
            if key in p_clean and len(p_clean) < len(key) + 10: # 确保是标题行
                set_next_para_text(i, val)

    # 3. 处理图片注入 (OK Picture & NG Picture)
    ok_img_bytes = img_data.get('OK') if img_data else None
    ng_img_bytes = img_data.get('NG') if img_data else None

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                cell_text = cell.text.strip()
                if cell_text == 'OK-Part' or cell_text == 'OK Picture':
                    cell.text = '' # 清空占位文本
                    p = cell.paragraphs[0]
                    p.alignment = docx.enum.text.WD_ALIGN_PARAGRAPH.CENTER
                    if ok_img_bytes:
                        try:
                            p.add_run().add_picture(io.BytesIO(ok_img_bytes), width=Inches(3.0))
                        except Exception:
                            p.add_run("图片加载失败")
                    else:
                        p.add_run("无 OK Picture 图片")
                elif cell_text == 'Not-OK-Part' or cell_text == 'NG Picture' or cell_text == 'NOK-Part':
                    cell.text = ''
                    p = cell.paragraphs[0]
                    p.alignment = docx.enum.text.WD_ALIGN_PARAGRAPH.CENTER
                    if ng_img_bytes:
                        try:
                            p.add_run().add_picture(io.BytesIO(ng_img_bytes), width=Inches(3.0))
                        except Exception:
                            p.add_run("图片加载失败")
                    else:
                        p.add_run("无 NG Picture 图片")

    return doc

# 生成邮件草稿
def generate_eml_file(row_data):
    serial_no = str(row_data.get('LL Serials No', 'LL-xxxx-xx')).strip()
    failure_mode = str(row_data.get('Failure Mode', '*****')).strip()
    subject = f"M/PQR-AP LL | {serial_no} | Title {failure_mode}"
    
    html_body = f"""
    <html>
    <head><style>body {{ font-family: 'Arial', sans-serif; font-size: 10.5pt; }} .red-bold {{ color: #FF0000; font-weight: bold; }}</style></head>
    <body>
        <p>Dear Supplier:</p>
        <p>Recently, we summarized a lesson learn about <strong>{failure_mode}</strong>. Please review the attached LL document and complete following tasks:</p>
        <ul>
            <li>Complete feedback form based on self-evaluation on your own processes and send to your responsible PQR and PUQ-PQA (ME) within <span class="red-bold">one week.</span></li>
            <li>After your self-evaluation, please close defined actions within <span class="red-bold">three weeks.</span></li>
            <li>Our PQR or PUQ-PQA colleague may conduct onsite verification according to the information in feedback form in <span class="red-bold">one month.</span></li>
        </ul>
        <p>If you have any question about this lesson learn, please contact with your responsible PQR and PUQ-PQA (ME).</p>
        <p>Best regards,</p><p><strong>Purchasing Quality Region Asia Pacific Team</strong></p>
    </body>
    </html>
    """
    msg = MIMEMultipart('alternative')
    msg['Subject'] = Header(subject, 'utf-8')
    msg['From'] = 'Sunny.LIU3@cn.bosch.com'
    msg['To'] = ''
    msg.attach(MIMEText(html_body, 'html', 'utf-8'))
    return msg.as_bytes()

# 页面渲染主逻辑
if excel_file is not None and template_file is not None:
    try:
        # 获取Excel所有底层图片
        excel_images = {}
        if not isinstance(excel_file, str):
            excel_file.seek(0)
            
        df, sheet_name, header_idx = load_excel_robust(excel_file)
        
        # 恢复指针再次读取图片
        if not isinstance(excel_file, str):
            excel_file.seek(0)
        excel_images_raw = load_excel_images(excel_file, sheet_name)
        
        st.success(f"🎉 成功加载工作表: **{sheet_name}**")
        
        # 寻找匹配列名
        supplier_scope_col = next((c for c in df.columns if 'Supplier Scope' in c or 'Scope' in c or 'Task' in c), 'LL Supplier Scope')
        ok_col_idx = df.columns.get_loc('OK Picture') if 'OK Picture' in df.columns else -1
        ng_col_idx = df.columns.get_loc('NG Picture') if 'NG Picture' in df.columns else -1
        
        st.markdown("### 👈 第一步：在下方表格中勾选需要生成的 Record")
        search_term = st.text_input("🔍 快速搜索：")
        
        filtered_df = df.copy()
        if search_term:
            filtered_df = df[df.astype(str).apply(lambda row: row.str.contains(search_term, case=False).any(), axis=1)]
        
        filtered_df.insert(0, '选择 (Select)', False)
        display_cols = ['选择 (Select)', 'LL Serials No', 'Failure Mode', 'Project/Part name', 'LL Brief Description']
        available_cols = [c for c in display_cols if c in filtered_df.columns]
        
        edited_df = st.data_editor(
            filtered_df[available_cols], hide_index=True,
            column_config={"选择 (Select)": st.column_config.CheckboxColumn("选择", default=False)},
            disabled=[col for col in available_cols if col != '选择 (Select)'], use_container_width=True
        )
        
        selected_indices = edited_df[edited_df['选择 (Select)'] == True].index
        selected_rows = filtered_df.loc[selected_indices]
        
        st.write("---")
        st.markdown("### 👉 第二步：一键生成与下载")
        
        if len(selected_rows) > 0:
            if st.button("🚀 开始生成 Word 与邮件", type="primary", use_container_width=True):
                with st.spinner("正在填充模板并提取图片..."):
                    
                    zip_buffer = io.BytesIO()
                    with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
                        for idx, (original_df_index, row) in enumerate(selected_rows.iterrows()):
                            
                            # 计算Excel中真实的物理行号 (0-indexed)。pandas索引 + 表头所在行 + 1 (因为数据从表头下一行开始)
                            excel_row_idx = original_df_index + header_idx + 1
                            
                            # 提取对应行的图片
                            row_img_data = {
                                'OK': excel_images_raw.get((excel_row_idx, ok_col_idx)) if ok_col_idx != -1 else None,
                                'NG': excel_images_raw.get((excel_row_idx, ng_col_idx)) if ng_col_idx != -1 else None
                            }
                            
                            row_data = {
                                'LL Serials No': row.get('LL Serials No', f"Record_{idx+1}"),
                                'Failure Mode': row.get('Failure Mode', ''),
                                'Project/Part name': row.get('Project/Part name', ''),
                                'Related Material Field / Process': row.get('Related Material Field / Process', ''),
                                'LL Brief Description': row.get('LL Brief Description', ''),
                                'Root Cause': row.get('Root Cause', ''),
                                'Corrective Action': row.get('Corrective Action', ''),
                                'Should or not to do': row.get('Should or not to do', ''),
                                'LL Supplier Scope': row.get(supplier_scope_col, '')
                            }
                            
                            doc = fill_word_template(template_file, row_data, img_data=row_img_data)
                            doc_io = io.BytesIO()
                            doc.save(doc_io)
                            serial_str = str(row_data['LL Serials No'])
                            zip_file.writestr(f"LL_Template_{serial_str}.docx", doc_io.getvalue())
                            zip_file.writestr(f"Email_Draft_{serial_str}.eml", generate_eml_file(row_data))
                            
                    zip_buffer.seek(0)
                    st.download_button(
                        label="📥 立即下载打包好的 ZIP 压缩包",
                        data=zip_buffer.read(),
                        file_name="LL_Templates_and_Emails_Batch.zip",
                        mime="application/zip",
                        use_container_width=True
                    )
                    st.success("✨ 批量生成成功！Word 模板 (含图片) 与配套邮件草稿已全部打包完毕。")
        else:
            st.warning("👉 请在上方勾选生成的记录。")
            
    except Exception as e:
        st.error(f"❌ 运行出错: {e}")
        st.info("排查提示：请确认 Excel 文件未被打开，且环境中已安装 `openpyxl` 库处理图片。")
else:
    st.info("ℹ️ 请在侧边栏上传 Excel 与 Word 模板。")















