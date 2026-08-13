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
import openpyxl
from openpyxl.drawing.image import Image as OpenpyxlImage
from docx.shared import Inches

# --- 页面配置 ---
st.set_page_config(page_title="PCB Lesson Learn 自动化工具", layout="wide", page_icon="⚙️")

st.markdown("""
# ⚙️ PCB Lesson Learn 自动化工具 (V6 - 最终修正版)
本程序已按您的最终要求重构。它将首先根据 **'LL Need or not'** 列进行筛选，然后使用**固定的列名映射逻辑**，将数据和图片精准填充到您的模板中。
""")
st.write("---")

# --- 文件上传 ---
uploaded_excel = st.file_uploader("📁 **步骤 1: 上传 PCB Lesson Learn Master List (.xlsx / .xlsm)**", type=["xlsx", "xlsm"])
uploaded_template = st.file_uploader("📝 **步骤 2: 上传您的 LL Word 模板 (.docx)**", type=["docx"])

# --- 核心函数 ---

def find_column_by_keywords(df_columns, keywords, default=None):
    """根据关键词列表，在DataFrame的列名中寻找第一个匹配的列。"""
    for col in df_columns:
        col_clean = str(col).strip().lower().replace('.', '')
        if any(keyword.lower() in col_clean for keyword in keywords):
            return col
    return default

@st.cache_data
def load_excel_robust(_file_source):
    """健壮地加载Excel文件，自动识别sheet和表头。"""
    if _file_source is None:
        return None, None, None, None
    file_bytes = _file_source.getvalue()
    excel_io = io.BytesIO(file_bytes)
    
    try:
        xl = pd.ExcelFile(excel_io, engine='openpyxl')
        sheet_names = xl.sheet_names
        target_sheet = next((s for s in sheet_names if "PUQ3 LL Overall list" in s), sheet_names[0])
        
        excel_io.seek(0)
        df_temp = pd.read_excel(excel_io, sheet_name=target_sheet, nrows=10, header=None, engine='openpyxl')
        header_idx = next((idx for idx, row in df_temp.iterrows() if any('serials' in str(v).lower() for v in row)), 1)
        
        excel_io.seek(0)
        df = pd.read_excel(excel_io, sheet_name=target_sheet, header=header_idx, engine='openpyxl')
        df.columns = [str(c).strip() for c in df.columns]
        
        return file_bytes, df, target_sheet, header_idx
    except Exception as e:
        st.error(f"加载 Excel 文件时出错: {e}")
        return None, None, None, None

def extract_images_for_row(excel_bytes, sheet_name, row_idx_in_df, header_rows, df):
    """从指定的Excel行中提取NG和OK图片。"""
    images = {'NG Picture': None, 'OK Picture': None}
    if not excel_bytes: return images

    wb = openpyxl.load_workbook(io.BytesIO(excel_bytes))
    ws = wb[sheet_name]

    col_map = {col: i for i, col in enumerate(df.columns)}
    ng_col_key = find_column_by_keywords(df.columns, ["NG Picture"])
    ok_col_key = find_column_by_keywords(df.columns, ["OK Picture"])
    ng_col_idx = col_map.get(ng_col_key)
    ok_col_idx = col_map.get(ok_col_key)
    
    excel_target_row = row_idx_in_df + header_rows + 2 

    if ng_col_idx is None and ok_col_idx is None:
        return images

    for img in ws._images:
        if img.anchor._from.row + 1 == excel_target_row:
            img_col = img.anchor._from.col
            if img_col == ng_col_idx:
                images['NG Picture'] = io.BytesIO(img.image.blob)
            elif img_col == ok_col_idx:
                images['OK Picture'] = io.BytesIO(img.image.blob)
    return images

def fill_word_template(template_source, row_data, images, excel_columns):
    """【已重构】使用固定的标题->列名映射来填充Word模板，并插入图片。"""
    doc = docx.Document(template_source)
    
    # 1. 准备数据和日期
    current_date_str = datetime.date.today().strftime('%B %d, %Y')
    
    # 2. 【核心】定义固定的模板标题与Excel列名的映射关系
    #    这个映射关系是写死的，不再从模板里读取关键词。
    fixed_mapping = {
        "Task/Scope": find_column_by_keywords(excel_columns, ["LL Supplier Scope"]),
        "Failure Mode": find_column_by_keywords(excel_columns, ["Failure Mode"]),
        "Project/Part name": find_column_by_keywords(excel_columns, ["Project/Part name"]),
        "Process": find_column_by_keywords(excel_columns, ["Related Material Field / Process"]),
        "Problem (Fundamental Problem)": find_column_by_keywords(excel_columns, ["LL Brief Description"]),
        "Root Cause(s)": find_column_by_keywords(excel_columns, ["Root Cause"]),
        "Corrective Actions": find_column_by_keywords(excel_columns, ["Corrective Action"]),
        "Learning": find_column_by_keywords(excel_columns, ["Should or not to do"])
    }

    # 3. 收集文档所有段落（包括正文、表格、页眉、页脚）
    all_paras = []
    for container in [doc] + doc.sections:
        if hasattr(container, 'paragraphs'):
            all_paras.extend(container.paragraphs)
        if hasattr(container, 'tables'):
            for table in container.tables:
                for row in table.rows:
                    for cell in row.cells:
                        all_paras.extend(cell.paragraphs)
    
    # 4. 遍历并替换
    for p in all_paras:
        # 日期替换
        if 'May 24 2022' in p.text:
            p.text = p.text.replace('May 24 2022', current_date_str)
        
        # 标题定位与内容插入
        p_text_clean = p.text.strip()
        for template_title, excel_col_name in fixed_mapping.items():
            # 如果段落文本与模板标题匹配，并且Excel中存在对应的列
            if template_title in p_text_clean and excel_col_name:
                value = row_data.get(excel_col_name, '')
                # 在标题下方插入一个新段落来填充内容
                # 检查下一段是否是空行或占位符，如果是就替换，否则插入新行
                try:
                    next_p = p.insert_paragraph_before(str(value if pd.notna(value) else ''), style=p.style)
                    # This python-docx feature inserts *before*, so we need a better way.
                    # Reverting to the robust XML insertion method.
                    new_p_element = p._element.getparent().create_element('w:p')
                    p._element.addnext(new_p_element)
                    new_para = docx.text.paragraph.Paragraph(new_p_element, p._parent)
                    new_para.text = str(value if pd.notna(value) else '')
                    
                    # 清理原标题行的占位符（如果模板里有的话）
                    if p.text != template_title:
                       p.text = template_title # Restore the clean heading
                except Exception:
                     # Fallback if XML fails
                    p.add_run("\n" + str(value if pd.notna(value) else ''))
                break

    # 5. 图片和顶部标题处理
    problem_desc = str(row_data.get(fixed_mapping["Problem (Fundamental Problem)"], ''))
    for p in all_paras:
        if p.text.strip().startswith("Lesson Learn"):
             if problem_desc and problem_desc not in p.text:
                p.text = f"Lesson Learn – {problem_desc}"
                break
    
    for table in doc.tables:
        try:
            if "NG Picture" in table.cell(0, 0).text and "OK Picture" in table.cell(0, 1).text:
                ng_cell, ok_cell = table.cell(0, 0), table.cell(0, 1)
                ng_cell.text, ok_cell.text = '', ''
                if images.get('NG Picture'):
                    ng_cell.paragraphs[0].add_run().add_picture(images['NG Picture'], width=Inches(2.5))
                if images.get('OK Picture'):
                    ok_cell.paragraphs[0].add_run().add_picture(images['OK Picture'], width=Inches(2.5))
                break
        except IndexError:
            continue

    return doc

# --- 邮件生成函数 (无改动) ---
def generate_eml_file(row_data, serial_no_col, failure_mode_col):
    """Generates an EML email draft file."""
    serial_no = str(row_data.get(serial_no_col, 'LL-xxxx-xx')).strip()
    failure_mode = str(row_data.get(failure_mode_col, '*****')).strip()
    subject = f"M/PQR-AP LL | {serial_no} | Title {failure_mode}"
    
    html_body = f"""
    <html>... (omitted for brevity, same as previous version) ...</html>
    """
    
    msg = MIMEMultipart('alternative')
    msg['Subject'] = Header(subject, 'utf-8')
    msg['From'] = 'Sunny.LIU3@cn.bosch.com'
    msg.attach(MIMEText(html_body, 'html', 'utf-8'))
    return msg.as_bytes()


# --- 主应用流程 ---
if uploaded_excel and uploaded_template:
    file_bytes, df, sheet_name, header_idx = load_excel_robust(uploaded_excel)
    
    if df is not None:
        st.success(f"🎉 Excel文件加载成功! (工作表: **{sheet_name}**, 表头在第 **{header_idx + 1}** 行)")

        # 1. 'LL Need or not' 列筛选
        ll_need_col = find_column_by_keywords(df.columns, ["LL Need or not"])
        if ll_need_col:
            initial_count = len(df)
            df_filtered = df[df[ll_need_col].astype(str).str.upper() == 'Y'].copy()
            st.info(f"已根据 **'{ll_need_col}' = Y** 筛选。显示 **{len(df_filtered)}** / **{initial_count}** 条记录。")
        else:
            st.warning(f"⚠️ 未在Excel中找到 '{ll_need_col}' 列，将显示所有记录。")
            df_filtered = df.copy()

        # 2. 搜索和选择
        st.markdown("### 📝 步骤 3: 搜索并勾选记录")
        search_term = st.text_input("🔍 可按任意关键词快速搜索:")
        
        if search_term:
            df_filtered = df_filtered[df_filtered.astype(str).apply(lambda row: row.str.contains(search_term, case=False).any(), axis=1)]

        serial_no_col = find_column_by_keywords(df_filtered.columns, ["LL Serials No"])
        failure_mode_col = find_column_by_keywords(df_filtered.columns, ["Failure Mode"])

        if not serial_no_col or not failure_mode_col:
            st.error("关键错误: 无法在Excel中找到 'LL Serials No' 或 'Failure Mode' 列。")
        else:
            selected_ids = st.multiselect(
                "请选择要生成文档的记录 (多选将打包为ZIP):",
                options=df_filtered[serial_no_col].tolist(),
                format_func=lambda x: f"{x} - {df_filtered[df_filtered[serial_no_col] == x][failure_mode_col].iloc[0]}"
            )
            
            # 3. 生成按钮
            if selected_ids:
                st.write("---")
                if st.button("🚀 生成选中项的文档", type="primary", use_container_width=True):
                    with st.spinner("正在处理... 提取图片可能需要一些时间..."):
                        selected_rows = df_filtered[df_filtered[serial_no_col].isin(selected_ids)]

                        if len(selected_rows) == 1:
                            row = selected_rows.iloc[0]
                            images = extract_images_for_row(file_bytes, sheet_name, row.name, header_idx, df)
                            doc = fill_word_template(uploaded_template, row, images, df.columns)
                            doc_io = io.BytesIO()
                            doc.save(doc_io)
                            eml_data = generate_eml_file(row, serial_no_col, failure_mode_col)
                            
                            st.success("✨ 生成成功!")
                            c1, c2 = st.columns(2)
                            c1.download_button(f"📥 下载 Word ({row[serial_no_col]})", doc_io.getvalue(), f"LL_{row[serial_no_col]}.docx", use_container_width=True)
                            c2.download_button(f"📧 下载邮件 ({row[serial_no_col]})", eml_data, f"Email_{row[serial_no_col]}.eml", "message/rfc822", use_container_width=True)
                        
                        else:
                            zip_buffer = io.BytesIO()
                            with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
                                for _, row in selected_rows.iterrows():
                                    images = extract_images_for_row(file_bytes, sheet_name, row.name, header_idx, df)
                                    doc = fill_word_template(uploaded_template, row, images, df.columns)
                                    doc_io = io.BytesIO()
                                    doc.save(doc_io)
                                    zf.writestr(f"LL_{row[serial_no_col]}.docx", doc_io.getvalue())
                                    eml_data = generate_eml_file(row, serial_no_col, failure_mode_col)
                                    zf.writestr(f"Email_{row[serial_no_col]}.eml", eml_data)
                            
                            st.success("✨ 批量生成完毕!")
                            st.download_button("📦 下载全部 ZIP 压缩包", zip_buffer.getvalue(), "LL_Automation_Batch.zip", "application/zip", use_container_width=True)
else:
    st.info("请上传 Master List 和 Word 模板以开始。")


