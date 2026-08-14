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
import xml.etree.ElementTree as ET
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.header import Header

# 页面基本配置
st.set_page_config(page_title="PCB Lesson Learn 模板生成器", layout="wide", page_icon="🌅")

st.markdown("""
# 🌅 PCB Lesson Learn 模板生成器 & 邮件草稿一键生成

本程序支持上传 **.xlsx** 及 **.xlsm（带宏）** 格式文件。
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
    uploaded_excel = st.sidebar.file_uploader("手动上传 PCB Master List (Excel):", type=["xlsx", "xlsm"])
    uploaded_template = st.sidebar.file_uploader("手动上传 LL Template (Word):", type=["docx"])
    if uploaded_excel:
        excel_file = uploaded_excel
    if uploaded_template:
        template_file = uploaded_template

# 如果没有找到默认文件，在主界面显示上传入口
if excel_file is None:
    uploaded_excel = st.file_uploader("📁 请上传 PCB Lesson Learn Master List (Excel/xlsm)", type=["xlsx", "xlsm"], key="excel_main")
    if uploaded_excel:
        excel_file = uploaded_excel
if template_file is None:
    uploaded_template = st.file_uploader("📝 请上传 LL Template (Word) 模板文件", type=["docx"], key="template_main")
    if uploaded_template:
        template_file = uploaded_template

# 新增：从 Excel 中智能提取 NG 和 OK 图片的函数
def get_images_for_row(file_source, sheet_name, header_idx, target_row_idx):
    import openpyxl
    try:
        if isinstance(file_source, str):
            wb = openpyxl.load_workbook(file_source, data_only=True)
        else:
            file_source.seek(0)
            wb = openpyxl.load_workbook(file_source, data_only=True)
            
        ws = wb[sheet_name]
        
        col_ng = -1
        col_ok = -1
        for col_idx in range(1, ws.max_column + 1):
            val = ws.cell(row=header_idx + 1, column=col_idx).value
            if val:
                val_str = str(val).strip()
                if 'NG Picture' in val_str: col_ng = col_idx - 1
                if 'OK Picture' in val_str: col_ok = col_idx - 1
                
        excel_target_row = header_idx + 1 + target_row_idx
        ok_img = None
        ng_img = None
        
        for img in getattr(ws, '_images', []):
            try:
                r = img.anchor._from.row
                c = img.anchor._from.col
                if r == excel_target_row:
                    if c == col_ng:
                        ng_img = img._data()
                    elif c == col_ok:
                        ok_img = img._data()
            except Exception:
                pass
                
        return ok_img, ng_img
    except Exception as e:
        return None, None

# 核心逻辑：加载 Excel，智能识别 Sheet 和表头行
def load_excel_robust(file_source):
    if not isinstance(file_source, str):
        file_source.seek(0)
    xl = pd.ExcelFile(file_source)
    sheet_names = xl.sheet_names
    
    target_sheet = None
    for sheet in sheet_names:
        if "PUQ3 LL Overall list" in sheet:
            target_sheet = sheet
            break
    if not target_sheet:
        for sheet in sheet_names:
            if "Overall" in sheet or "LL" in sheet:
                target_sheet = sheet
                break
    if not target_sheet:
        target_sheet = sheet_names[0]
        
    if not isinstance(file_source, str):
        file_source.seek(0)
    df_temp = pd.read_excel(file_source, sheet_name=target_sheet, nrows=10, header=None)
    header_idx = 1 
    
    for idx, row in df_temp.iterrows():
        row_str = [str(val).strip().lower() for val in row.tolist()]
        if any('serials' in val or 'project/part' in val or 'failure mode' in val for val in row_str):
            header_idx = idx
            break
            
    if not isinstance(file_source, str):
        file_source.seek(0)
    df = pd.read_excel(file_source, sheet_name=target_sheet, header=header_idx)
    df.columns = df.columns.astype(str).str.strip()
    return df, target_sheet, header_idx

# 核心逻辑：强力填充 Word 模板
def fill_word_template(template_source, row_data):
    doc = docx.Document(template_source)
    from docx.text.paragraph import Paragraph
    from docx.shared import Inches
    
    current_date_str = datetime.date.today().strftime('%b %d %Y')
    failure_mode_str = str(row_data.get('Failure Mode', '')).strip()
    ns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
    
    all_p_elements = []
    all_p_elements.extend(doc._element.findall('.//w:p', ns))
    
    for section in doc.sections:
        headers_footers = [
            section.header, section.footer,
            section.first_page_header, section.first_page_footer,
            section.even_page_header, section.even_page_footer
        ]
        for hf in headers_footers:
            if hf is None:
                continue
            all_p_elements.extend(hf._element.findall('.//w:p', ns))
            
    all_p_elements = list(set(all_p_elements))
    
    for p_elem in all_p_elements:
        t_elems = p_elem.findall('.//w:t', ns)
        if not t_elems:
            continue
            
        p_text = "".join([t.text for t in t_elems if t.text])
        p_text_lower = p_text.lower()
        
        if 'may 24 2022' in p_text_lower:
            new_text = p_text.replace('May 24 2022', current_date_str).replace('May 24, 2022', current_date_str)
            t_elems[0].text = new_text
            for t in t_elems[1:]:
                t.text = ""
                
        if 'lesson learn' in p_text_lower and len(p_text_lower) < 60:
            if failure_mode_str and failure_mode_str not in p_text:
                p_text_clean = p_text.strip().rstrip("–-—: ").strip()
                new_val = f"{p_text_clean} – {failure_mode_str}"
                t_elems[0].text = new_val
                for t in t_elems[1:]:
                    t.text = ""

    should_or_not = str(row_data.get('Should or not to do', ''))
    should_text = ""
    should_not_text = ""
    if 'Should not:' in should_or_not:
        parts = should_or_not.split('Should not:')
        should_text = parts[0].replace('Should:', '').strip()
        should_not_text = parts[1].strip()
    else:
        should_text = should_or_not.replace('Should:', '').strip()

    body_p_elements = doc._body._body.findall('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}p')
    
    headings_config = [
        {'keys': ['Task/Scope', 'Task'], 'value': row_data.get('LL Supplier Scope', '')},
        {'keys': ['Failure Mode'], 'value': row_data.get('Failure Mode', '')},
        {'keys': ['Project/Part name', 'Product / Process'], 'value': row_data.get('Project/Part name', '')},
        {'keys': ['Process'], 'value': row_data.get('Related Material Field / Process', '')},
        {'keys': ['Problem (Fundamental Problem)', 'Problem'], 'value': row_data.get('LL Brief Description', '')},
        {'keys': ['Root Cause(s)', 'Root Cause'], 'value': row_data.get('Root Cause', '')},
        {'keys': ['Corrective Actions', 'Corrective Action'], 'value': row_data.get('Corrective Action', '')},
        {'keys': ['What should we do in the future?'], 'value': should_text},
        {'keys': ['What should we not do in the future?'], 'value': should_not_text}
    ]
    
    p_indices = {}
    for idx, p_elem in enumerate(body_p_elements):
        try:
            p = Paragraph(p_elem, doc)
            p_text = p.text.strip()
            for i, config in enumerate(headings_config):
                for key in config['keys']:
                    if key in p_text:
                        if i not in p_indices:
                            p_indices[i] = idx
                        break
        except Exception:
            pass
            
    sorted_configs = sorted(p_indices.items(), key=lambda x: x[1], reverse=True)
    
    for i, p_idx in sorted_configs:
        val = headings_config[i]['value']
        if not val or pd.isna(val):
            val = ""
            
        p_elem = body_p_elements[p_idx]
        try:
            p = Paragraph(p_elem, doc)
            replaced = False
            
            if p_idx + 1 < len(body_p_elements):
                next_p_elem = body_p_elements[p_idx + 1]
                next_p = Paragraph(next_p_elem, doc)
                next_text = next_p.text.strip()
                
                if next_text in ['...', '…', ''] or len(next_text) < 5:
                    is_another_heading = False
                    for other_i, config in enumerate(headings_config):
                        for key in config['keys']:
                            if key in next_text:
                                is_another_heading = True
                    if not is_another_heading:
                        next_p.text = str(val)
                        replaced = True
                        
            if not replaced:
                new_p_element = p._element.getparent().create_element('w:p')
                p._element.addnext(new_p_element)
                new_para = Paragraph(new_p_element, p._parent)
                new_para.text = str(val)
        except Exception:
            pass
            
    # 【修复图片错位】彻底清除多余回车符，防止表格断层
    ok_img = row_data.get('OK Picture Bytes')
    ng_img = row_data.get('NG Picture Bytes')
    
    if ok_img or ng_img:
        pic_table = None
        for table in doc.tables:
            text = "".join(cell.text for row in table.rows for cell in row.cells)
            if "OK-Part" in text and "Not-OK-Part" in text:
                pic_table = table
                break
                
        if pic_table:
            ok_cell, ng_cell = None, None
            for row in pic_table.rows:
                for cell in row.cells:
                    if "Not-OK-Part" in cell.text:
                        ng_cell = cell
                    elif "OK-Part" in cell.text:
                        ok_cell = cell
            
            # 清理 OK-Part 并插入图片
            if ok_img and ok_cell:
                inserted = False
                for p in list(ok_cell.paragraphs):
                    p_text_lower = p.text.lower().replace(" ", "")
                    # 保留表头文字
                    if "ok-part" in p_text_lower:
                        continue
                    # 第一次遇到多余段落时，清空内容并塞入图片
                    if not inserted:
                        p.text = "" 
                        p.alignment = 1 # 居中对齐
                        r = p.add_run()
                        r.add_picture(io.BytesIO(ok_img), width=Inches(2.5))
                        inserted = True
                    else:
                        # 核心修复：物理删除多余的空段落（回车符），完美解决错位断层
                        p._element.getparent().remove(p._element)
                        
                if not inserted:
                    p = ok_cell.add_paragraph()
                    p.alignment = 1
                    r = p.add_run()
                    r.add_picture(io.BytesIO(ok_img), width=Inches(2.5))
                    
            # 清理 Not-OK-Part 并插入图片
            if ng_img and ng_cell:
                inserted = False
                for p in list(ng_cell.paragraphs):
                    p_text_lower = p.text.lower().replace(" ", "")
                    if "not-ok-part" in p_text_lower:
                        continue
                    if not inserted:
                        p.text = ""
                        p.alignment = 1
                        r = p.add_run()
                        r.add_picture(io.BytesIO(ng_img), width=Inches(2.5))
                        inserted = True
                    else:
                        # 物理删除多余的空段落
                        p._element.getparent().remove(p._element)
                        
                if not inserted:
                    p = ng_cell.add_paragraph()
                    p.alignment = 1
                    r = p.add_run()
                    r.add_picture(io.BytesIO(ng_img), width=Inches(2.5))
                                
    return doc

# 核心逻辑：自动生成符合 Outlook 规范的 EML 邮件文件
def generate_eml_file(row_data):
    serial_no = str(row_data.get('LL Serials No', 'LL-xxxx-xx')).strip()
    failure_mode = str(row_data.get('Failure Mode', '*****')).strip()
    
    subject = f"M/PQR-AP LL | {serial_no} | Title {failure_mode}"
    
    html_body = f"""
    <html>
    <head>
        <style>
            body {{
                font-family: 'Arial', sans-serif;
                font-size: 10.5pt;
                line-height: 1.6;
                color: #333333;
            }}
            .red-bold {{
                color: #FF0000;
                font-weight: bold;
            }}
            ul {{
                margin-top: 5px;
                margin-bottom: 15px;
                padding-left: 20px;
            }}
            li {{
                margin-bottom: 8px;
            }}
        </style>
    </head>
    <body>
        <p>Dear Supplier:</p>
        <p>Recently, we summarized a lesson learn about <strong>{failure_mode}</strong>. Please review the attached LL document and complete following tasks:</p>
        <ul>
            <li>Complete feedback form based on self-evaluation on your own processes and send to your responsible PQR and PUQ-PQA (ME) within <span class="red-bold">one week.</span></li>
            <li>After your self-evaluation, please close defined actions within <span class="red-bold">three weeks.</span></li>
            <li>Our PQR or PUQ-PQA colleague may conduct onsite verification according to the information in feedback form in <span class="red-bold">one month.</span></li>
        </ul>
        <p>If you have any question about this lesson learn, please contact with your responsible PQR and PUQ-PQA (ME).</p>
        <p>Best regards,</p>
        <p><strong>Purchasing Quality Region Asia Pacific Team</strong></p>
    </body>
    </html>
    """
    
    msg = MIMEMultipart('alternative')
    msg['Subject'] = Header(subject, 'utf-8')
    msg['From'] = 'Sunny.LIU3@cn.bosch.com'
    msg['To'] = ''  
    
    part = MIMEText(html_body, 'html', 'utf-8')
    msg.attach(part)
    
    return msg.as_bytes()

# 5. 页面渲染与交互
if excel_file is not None and template_file is not None:
    try:
        df, sheet_name, header_idx = load_excel_robust(excel_file)
        st.success(f"🎉 成功加载工作表: **{sheet_name}** (表头定位自第 {header_idx + 1} 行)")
        
        # 【新增功能】判断 LL Need or not 列并过滤
        ll_need_col = 'LL Need or not'
        if ll_need_col not in df.columns:
            for col in df.columns:
                if 'need or not' in str(col).lower():
                    ll_need_col = col
                    break
                    
        if ll_need_col in df.columns:
            original_len = len(df)
            df = df[df[ll_need_col].astype(str).str.strip().str.upper() == 'Y']
            filtered_len = len(df)
            st.info(f"💡 已自动过滤数据：仅保留 '{ll_need_col}' 状态为 'Y' 的有效记录 (共排除了 {original_len - filtered_len} 条非 Y 记录)。")
        
        serial_no_col = 'LL Serials No'
        for col in df.columns:
            col_clean = str(col).strip().replace('.', '').lower()
            if 'll serials no' in col_clean or 'serials no' in col_clean or 'serial no' in col_clean:
                serial_no_col = col
                break
                
        supplier_scope_col = 'LL Supplier Scope'
        if 'LL Supplier Scope' not in df.columns:
            potential = [c for c in df.columns if 'Supplier Scope' in c or 'Scope' in c or 'Task' in c]
            if potential:
                supplier_scope_col = potential[0]
                st.info(f"💡 'LL Supplier Scope' 自动匹配为 Excel 中的: '{supplier_scope_col}'")
        
        st.markdown("### 👈 第一步：在下方表格中搜索并勾选需要生成的 Record")
        search_term = st.text_input("🔍 快速搜索 (支持输入序列号、供应商、项目名称、失效模式进行实时筛选)：")
        
        filtered_df = df.copy()
        if search_term:
            filtered_df = df[df.astype(str).apply(lambda row: row.str.contains(search_term, case=False).any(), axis=1)]
        
        filtered_df.insert(0, '选择 (Select)', False)
        
        display_cols_list = ['选择 (Select)', serial_no_col, 'Failure Mode', 'Supplier Name', 'Project/Part name', 'LL Brief Description']
        available_display_cols = [c for c in display_cols_list if c in filtered_df.columns]
        
        edited_df = st.data_editor(
            filtered_df[available_display_cols],
            hide_index=True,
            column_config={
                "选择 (Select)": st.column_config.CheckboxColumn(
                    "选择",
                    help="勾选以选择此行生成 Template",
                    default=False,
                )
            },
            disabled=[col for col in available_display_cols if col != '选择 (Select)'],
            use_container_width=True
        )
        
        selected_indices = edited_df[edited_df['选择 (Select)'] == True].index
        selected_rows = filtered_df.loc[selected_indices]
        
        st.write("---")
        st.markdown("### 👉 第二步：一键生成与下载")
        
        if len(selected_rows) > 0:
            st.success(f"已勾选 **{len(selected_rows)}** 条记录。请点击下方按钮开始生成：")
            
            if st.button("🚀 开始批量生成 Word 与邮件草稿", type="primary", use_container_width=True):
                with st.spinner("正在填充模板与图片，请稍候... (提取图片视 Excel 大小可能会多花几秒)"):
                    
                    if len(selected_rows) == 1:
                        row = selected_rows.iloc[0]
                        target_idx = row.name
                        
                        ok_img, ng_img = get_images_for_row(excel_file, sheet_name, header_idx, target_idx)
                        
                        row_data = {
                            'LL Serials No': row.get(serial_no_col, 'LL-xxxx-xx'),
                            'Failure Mode': row.get('Failure Mode', '*****'),
                            'Project/Part name': row.get('Project/Part name', ''),
                            'LL Brief Description': row.get('LL Brief Description', ''),
                            'Root Cause': row.get('Root Cause', ''),
                            'Corrective Action': row.get('Corrective Action', ''),
                            'Should or not to do': row.get('Should or not to do', ''),
                            'Related Material Field / Process': row.get('Related Material Field / Process', ''),
                            'LL Supplier Scope': row.get(supplier_scope_col, ''),
                            'OK Picture Bytes': ok_img,
                            'NG Picture Bytes': ng_img
                        }
                        
                        doc = fill_word_template(template_file, row_data)
                        bio_doc = io.BytesIO()
                        doc.save(bio_doc)
                        bio_doc.seek(0)
                        
                        eml_data = generate_eml_file(row_data)
                        serial_str = str(row_data['LL Serials No'])
                        
                        col1, col2 = st.columns(2)
                        with col1:
                            st.download_button(
                                label=f"📥 下载 Word: LL_Template_{serial_str}.docx",
                                data=bio_doc.read(),
                                file_name=f"LL_Template_{serial_str}.docx",
                                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                                use_container_width=True
                            )
                        with col2:
                            st.download_button(
                                label=f"📧 下载 Outlook 邮件草稿: Email_Draft_{serial_str}.eml",
                                data=eml_data,
                                file_name=f"Email_Draft_{serial_str}.eml",
                                mime="message/rfc822",
                                use_container_width=True
                            )
                        st.success("✨ Word 模板与 Outlook 邮件草稿生成成功！请点击上方对应的按钮进行下载。")
                        
                    else:
                        zip_buffer = io.BytesIO()
                        with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
                            for idx, (real_row_idx, row) in enumerate(selected_rows.iterrows()):
                                
                                ok_img, ng_img = get_images_for_row(excel_file, sheet_name, header_idx, real_row_idx)
                                
                                row_data = {
                                    'LL Serials No': row.get(serial_no_col, f"Record_{idx+1}"),
                                    'Failure Mode': row.get('Failure Mode', '*****'),
                                    'Project/Part name': row.get('Project/Part name', ''),
                                    'LL Brief Description': row.get('LL Brief Description', ''),
                                    'Root Cause': row.get('Root Cause', ''),
                                    'Corrective Action': row.get('Corrective Action', ''),
                                    'Should or not to do': row.get('Should or not to do', ''),
                                    'Related Material Field / Process': row.get('Related Material Field / Process', ''),
                                    'LL Supplier Scope': row.get(supplier_scope_col, ''),
                                    'OK Picture Bytes': ok_img,
                                    'NG Picture Bytes': ng_img
                                }
                                
                                doc = fill_word_template(template_file, row_data)
                                doc_io = io.BytesIO()
                                doc.save(doc_io)
                                doc_io.seek(0)
                                
                                serial_str = str(row_data['LL Serials No'])
                                zip_file.writestr(f"LL_Template_{serial_str}.docx", doc_io.getvalue())
                                
                                eml_data = generate_eml_file(row_data)
                                zip_file.writestr(f"Email_Draft_{serial_str}.eml", eml_data)
                                
                        zip_buffer.seek(0)
                        st.download_button(
                            label="📥 立即下载打包好的 ZIP 压缩包 (包含所有已勾选的 Word 与邮件草稿)",
                            data=zip_buffer.read(),
                            file_name="LL_Templates_and_Emails_Batch.zip",
                            mime="application/zip",
                            use_container_width=True
                        )
                        st.success("✨ 批量生成成功！Word 模板与配套邮件草稿已全部打包成功。")
        else:
            st.warning("👉 请在上方表格的【选择】列中勾选您想要生成的记录。")
            
    except Exception as e:
        st.error(f"❌ 读取或处理文件时出错: {e}")
        st.info("排查提示：请确认 Excel 文件和 Word 模板未被本地 Excel/Word 软件独占打开。")
else:
    st.info("ℹ️ 请在侧边栏确认默认文件路径，或者直接在侧边栏中手动上传 Excel 数据源 and Word 模板。")
