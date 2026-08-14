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
from openpyxl import load_workbook
from docx.shared import Inches
from docx.text.paragraph import Paragraph

# 页面基本配置
st.set_page_config(page_title="PCB Lesson Learn 模板生成器", layout="wide", page_icon="🌅")

st.markdown("""
# 🌅 PCB Lesson Learn 模板生成器 & 邮件草稿一键生成
本程序已完美匹配 Excel (sheet: **PUQ3 LL Overall list**，表头通常自第二行开始)。
您只需在下方表格中勾选所需的记录，即可一键自动填充并生成 Word 模板及配套 Outlook 邮件草稿。
""")
st.write("---")

# 1. 默认路径配置
DEFAULT_EXCEL_PATH = r"G:\\02_7_M-PQA-RBAC1\\08_PQA_AE\\09_PQA2\\11_PCB\\04_Lessons learn\\PCB Lesson Learn Master List.xlsx"
DEFAULT_TEMPLATE_PATH = r"G:\\02_7_M-PQA-RBAC1\\08_PQA_AE\\09_PQA2\\11_PCB\\04_Lessons learn\\LL Template complete version.docx"

# 侧边栏配置面板
st.sidebar.header("⚙️ 配置面板")
use_local_paths = st.sidebar.checkbox("使用本地固定路径 (G 盘)", value=True)

excel_file = None
template_file = None

if use_local_paths:
    excel_path = st.sidebar.text_input("Excel 文件路径:", DEFAULT_EXCEL_PATH)
    template_path = st.sidebar.text_input("Template 模板路径:", DEFAULT_TEMPLATE_PATH)
    if os.path.exists(excel_path): excel_file = excel_path
    else: st.sidebar.warning("⚠️ 未在指定路径找到 Excel 文件，请尝试手动上传。")
    if os.path.exists(template_path): template_file = template_path
    else: st.sidebar.warning("⚠️ 未在指定路径找到 Word 模板，请尝试手动上传。")
else:
    excel_file = st.sidebar.file_uploader("手动上传 PCB Master List (Excel):", type=["xlsx"])
    template_file = st.sidebar.file_uploader("手动上传 LL Template (Word):", type=["docx"])

if not excel_file: excel_file = st.file_uploader("📁 请上传 PCB Lesson Learn Master List (Excel)", type=["xlsx"], key="excel_main")
if not template_file: template_file = st.file_uploader("📝 请上传 LL Template (Word) 模板文件", type=["docx"], key="template_main")

# 核心逻辑：加载 Excel
def load_excel_robust(file_source):
    xl = pd.ExcelFile(file_source)
    sheet_name = next((s for s in xl.sheet_names if "PUQ3 LL Overall list" in s), xl.sheet_names[0])
    df_temp = pd.read_excel(file_source, sheet_name=sheet_name, nrows=10, header=None)
    header_idx = 1
    for idx, row in df_temp.iterrows():
        row_str = [str(val).strip().lower() for val in row.tolist()]
        if any('serials' in val or 'project/part' in val for val in row_str):
            header_idx = idx
            break
    df = pd.read_excel(file_source, sheet_name=sheet_name, header=header_idx)
    df.columns = df.columns.astype(str).str.strip()
    return df, sheet_name, header_idx

# 核心逻辑：填充Word模板
def fill_word_template(template_source, row_data, df_index, excel_file_path, header_idx):
    doc = docx.Document(template_source)
    current_date_str = datetime.date.today().strftime('%B %d, %Y')

    # 替换页眉页脚和标题
    for section in doc.sections:
        for part in (section.header, section.footer, section.first_page_header, section.first_page_footer, section.even_page_header, section.even_page_footer):
            if part:
                for p in part.paragraphs:
                    p.text = p.text.replace('May 24 2022', current_date_str)

    # 替换主标题
    for p in doc.paragraphs:
        if 'Lesson Learn –' in p.text:
            p.text = f"Lesson Learn – {str(row_data.get('LL Brief Description', '')).strip()}"
            break

    # 查找并填充文本内容
    def find_and_fill(target_text, content_to_fill):
        content_to_fill = str(content_to_fill if pd.notna(content_to_fill) else "")
        for p in doc.paragraphs:
            if target_text.lower() in p.text.lower():
                # 尝试在下一段落填充
                next_elem = p._element.getnext()
                if next_elem is not None and next_elem.tag.endswith('p'):
                    next_p = Paragraph(next_elem, p_parent=p._parent)
                    # 清空并填充
                    for run in next_p.runs: run.clear()
                    next_p.add_run(content_to_fill)
                    return
                # 否则，在标题后插入新段落
                new_p = p.insert_paragraph_before(content_to_fill, style=p.style)
                return

    # 映射并填充
    find_and_fill('Task/Scope', row_data.get('LL Supplier Scope', ''))
    find_and_fill('Failure Mode', row_data.get('Failure Mode', ''))
    find_and_fill('Project/Part name', row_data.get('Project/Part name', ''))
    find_and_fill('Process', row_data.get('Related Material Field / Process', ''))
    find_and_fill('Problem (Fundamental Problem)', row_data.get('LL Brief Description', ''))
    find_and_fill('Root Cause(s)', row_data.get('Root Cause', ''))
    find_and_fill('Corrective Actions', row_data.get('Corrective Action', ''))

    # 拆分填充Should/Should not
    should_text = str(row_data.get('Should or not to do', ''))
    should_do = should_text.split('Should not:')[0].replace('Should:', '').strip()
    should_not_do = should_text.split('Should not:')[1].strip() if 'Should not:' in should_text else ''
    find_and_fill('What should we do in the future?', should_do)
    find_and_fill('What should we not do in the future?', should_not_do)
    
    # 提取并插入图片
    try:
        wb = load_workbook(excel_file_path)
        sheet_name = next((s for s in wb.sheetnames if "PUQ3 LL Overall list" in s), wb.sheetnames[0])
        ws = wb[sheet_name]
        
        col_map = {cell.value.strip(): cell.column for cell in ws[header_idx + 1]}
        ok_pic_col = col_map.get('OK Picture')
        ng_pic_col = col_map.get('NG Picture')
        
        # Excel行号 = dataframe的index + 表头所在行数 + 1 (因为pandas是0-indexed, openpyxl是1-indexed)
        excel_row = df_index + header_idx + 1

        pic_table = next((tbl for tbl in doc.tables if 'OK-Part' in tbl.cell(0,0).text), None)
        if pic_table:
            ok_cell, ng_cell = pic_table.cell(1, 0), pic_table.cell(1, 1) # Data is in the second row
            ok_cell.paragraphs[0].clear()
            ng_cell.paragraphs[0].clear()
            
            for img in ws._images:
                row_anchor = img.anchor._from.row + 1
                col_anchor = img.anchor._from.col + 1
                if row_anchor == excel_row:
                    img_data = io.BytesIO(img.ref)
                    if col_anchor == ok_pic_col:
                        ok_cell.paragraphs[0].add_run().add_picture(img_data, width=Inches(2.5))
                    elif col_anchor == ng_pic_col:
                        ng_cell.paragraphs[0].add_run().add_picture(img_data, width=Inches(2.5))
    except Exception as e:
        st.warning(f"⚠️ 处理图片时发生错误: {e}. 请确保Excel文件未被加密且格式正确。")
        
    return doc

# 核心逻辑：生成EML文件
def generate_eml_file(row_data):
    serial_no = str(row_data.get('LL Serials No', 'LL-xxxx-xx')).strip()
    failure_mode = str(row_data.get('Failure Mode', '******')).strip()
    subject = f"M/PQR-AP LL | {serial_no} | Title {failure_mode}"
    html_body = f"""
    <html><head><style>body{{font-family:'Arial',sans-serif;font-size:10.5pt;line-height:1.6;color:#333;}} .red-bold{{color:#FF0000;font-weight:bold;}} ul{{margin:5px 0 15px 20px;}} li{{margin-bottom:8px;}}</style></head>
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
    </body></html>
    """
    msg = MIMEMultipart('alternative')
    msg['Subject'] = Header(subject, 'utf-8')
    msg['From'] = 'Sunny.LIU3@cn.bosch.com'
    msg.attach(MIMEText(html_body, 'html', 'utf-8'))
    return msg.as_bytes()

# 页面渲染与交互
if excel_file and template_file:
    try:
        excel_path = excel_file if isinstance(excel_file, str) else excel_file.name
        if not isinstance(excel_file, str):
            with open(excel_path, "wb") as f: f.write(excel_file.getbuffer())

        df, sheet_name, header_idx = load_excel_robust(excel_path)
        st.success(f"🎉 成功加载工作表: **{sheet_name}** (表头定位自第 {header_idx + 1} 行)")
        
        required_cols = ['LL Serials No', 'Failure Mode', 'Project/Part name', 'LL Brief Description', 'Root Cause', 'Corrective Action', 'Should or not to do', 'Related Material Field / Process', 'LL Supplier Scope', 'OK Picture', 'NG Picture']
        if missing_cols := [c for c in required_cols if c not in df.columns]:
            st.error(f"❌ Excel 中缺失关键列: {', '.join(missing_cols)}")
        else:
            st.markdown("### 👈 第一步：在下方表格中搜索并勾选需要生成的 Record")
            search_term = st.text_input("🔍 快速搜索:")
            filtered_df = df[df.astype(str).apply(lambda row: row.str.contains(search_term, case=False).any(), axis=1)] if search_term else df
            
            edited_df = st.data_editor(
                filtered_df[['LL Serials No', 'Failure Mode', 'Supplier Name', 'Project/Part name']].assign(**{'选择': False}),
                column_config={"选择": st.column_config.CheckboxColumn(required=True)},
                hide_index=True, use_container_width=True
            )
            
            selected_rows = edited_df[edited_df['选择']]
            if not selected_rows.empty:
                st.markdown("### 👉 第二步：一键生成与下载")
                if st.button("🚀 开始批量生成 Word 与邮件草稿", type="primary", use_container_width=True):
                    with st.spinner("处理中..."):
                        # Get original indices from filtered_df
                        original_indices = selected_rows.index
                        rows_to_process = filtered_df.loc[original_indices]

                        # Single file generation
                        if len(rows_to_process) == 1:
                            row = rows_to_process.iloc[0]
                            doc = fill_word_template(template_file, row, rows_to_process.index[0], excel_path, header_idx)
                            doc_io = io.BytesIO()
                            doc.save(doc_io)
                            
                            col1, col2 = st.columns(2)
                            serial = str(row.get('LL Serials No', 'file'))
                            col1.download_button("📥 下载 Word", doc_io.getvalue(), f"LL_Template_{serial}.docx", use_container_width=True)
                            col2.download_button("📧 下载邮件草稿", generate_eml_file(row), f"Email_Draft_{serial}.eml", use_container_width=True)
                            st.success("✨ 生成成功！")
                        
                        # Zip generation for multiple files
                        else:
                            zip_buffer = io.BytesIO()
                            with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zf:
                                for idx, row in rows_to_process.iterrows():
                                    doc = fill_word_template(template_file, row, idx, excel_path, header_idx)
                                    doc_io = io.BytesIO()
                                    doc.save(doc_io)
                                    serial = str(row.get('LL Serials No', f"file_{idx}"))
                                    zf.writestr(f"LL_Template_{serial}.docx", doc_io.getvalue())
                                    zf.writestr(f"Email_Draft_{serial}.eml", generate_eml_file(row))
                            st.download_button("📥 下载全部 (ZIP)", zip_buffer.getvalue(), "LL_Batch.zip", "application/zip", use_container_width=True)
                            st.success("✨ 批量生成成功！")
            else:
                st.warning("👉 请在上方表格的【选择】列中勾选您想要生成的记录。")
    except Exception as e:
        st.error(f"❌ 处理文件时出错: {e}")
        st.info("排查提示：请确认 Excel 和 Word 模板格式正确且未被加密。")
else:
    st.info("ℹ️ 请在侧边栏配置或上传 Excel 数据源和 Word 模板。")








