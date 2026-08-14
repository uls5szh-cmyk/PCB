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
本程序已完美匹配 Excel (sheet: **PUQ3 LL Overall list**，表头自第二行开始)。
您只需在下方表格中勾选所需的记录，即可一键自动填充并生成 Word 模板及配套 Outlook 邮件草稿。
""")
st.write("---")

# 1. 默认路径配置（G 盘固定路径）
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
    else: st.sidebar.warning(f"⚠️ 未在指定路径找到 Excel 文件，请尝试手动上传。")
    if os.path.exists(template_path): template_file = template_path
    else: st.sidebar.warning(f"⚠️ 未在指定路径找到 Word 模板，请尝试手动上传。")
else:
    # 关键修改：同时支持 .xlsx 和 .xlsm
    uploaded_excel = st.sidebar.file_uploader("手动上传 PCB Master List (Excel):", type=["xlsx", "xlsm"])
    uploaded_template = st.sidebar.file_uploader("手动上传 LL Template (Word):", type=["docx"])
    if uploaded_excel: excel_file = uploaded_excel
    if uploaded_template: template_file = uploaded_template

# 主界面备用上传入口
if excel_file is None:
    uploaded_excel = st.file_uploader("📁 请上传 PCB Lesson Learn Master List (Excel)", type=["xlsx", "xlsm"], key="excel_main")
    if uploaded_excel: excel_file = uploaded_excel
if template_file is None:
    uploaded_template = st.file_uploader("📝 请上传 LL Template (Word) 模板文件", type=["docx"], key="template_main")
    if uploaded_template: template_file = uploaded_template

# 核心逻辑：加载 Excel (完全保留原版逻辑)
def load_excel_robust(file_source):
    # 关键修改：使用 openpyxl 引擎以支持 .xlsm
    engine = 'openpyxl'
    xl = pd.ExcelFile(file_source, engine=engine)
    sheet_names = xl.sheet_names
    
    target_sheet = None
    if "PUQ3 LL Overall list" in sheet_names:
        target_sheet = "PUQ3 LL Overall list"
    else:
        for sheet in sheet_names:
            if "Overall" in sheet or "LL" in sheet:
                target_sheet = sheet
                break
    if not target_sheet: target_sheet = sheet_names[0]
        
    df_temp = pd.read_excel(file_source, sheet_name=target_sheet, nrows=10, header=None, engine=engine)
    header_idx = 1 
    
    # 完全保留您原版的、健壮的表头识别逻辑
    for idx, row in df_temp.iterrows():
        row_str = [str(val).strip().lower() for val in row.tolist()]
        if any('serials' in val or 'project/part' in val or 'failure mode' in val for val in row_str):
            header_idx = idx
            break
            
    df = pd.read_excel(file_source, sheet_name=target_sheet, header=header_idx, engine=engine)
    df.columns = df.columns.astype(str).str.strip()
    return df, target_sheet, header_idx

# 核心逻辑：强力填充 Word 模板 (在原版基础上增加功能)
def fill_word_template(template_source, row_data, df_index, excel_source, header_idx):
    doc = docx.Document(template_source)
    
    # 1. 自动获取当前日期并格式化 (原版逻辑)
    current_date_str = datetime.date.today().strftime('%B %d, %Y')
    problem_text = str(row_data.get('LL Brief Description', '')).strip()
    ns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}

    # 2. 全域 w:p 节点深度穿透扫描列表 (完全保留您原版的XML遍历逻辑)
    all_p_elements = []
    all_p_elements.extend(doc._element.findall('.//w:p', ns))
    for section in doc.sections:
        for hf in (section.header, section.footer, section.first_page_header, section.first_page_footer, section.even_page_header, section.even_page_footer):
            if hf: all_p_elements.extend(hf._element.findall('.//w:p', ns))
    all_p_elements = list(set(all_p_elements))
    
    # 3. 终极深度 XML 文本节点替换 (完全保留您原版的XML替换逻辑)
    for p_elem in all_p_elements:
        t_elems = p_elem.findall('.//w:t', ns)
        p_text = "".join([t.text for t in t_elems if t.text])
        p_text_normalized = " ".join(p_text.lower().split())
        
        if 'may 24 2022' in p_text_normalized:
            for t in t_elems:
                if t.text:
                    t.text = t.text.replace('May 24 2022', current_date_str).replace('May 24, 2022', current_date_str)
                        
        if 'lesson' in p_text_normalized and 'learn' in p_text_normalized and len(p_text_normalized) < 50:
            if problem_text and problem_text not in p_text:
                p_text_clean = p_text.strip().rstrip("–-—: ").strip()
                new_val = f"{p_text_clean} – {problem_text}"
                if t_elems:
                    t_elems[0].text = new_val
                    for t in t_elems[1:]: t.text = ""
    
    # 4. **新增功能**: 提取并插入图片
    try:
        wb = load_workbook(excel_source, read_only=True, keep_vba=True)
        sheet_name = next((s for s in wb.sheetnames if "PUQ3 LL Overall list" in s), wb.sheetnames[0])
        ws = wb[sheet_name]
        col_map = {str(cell.value).strip(): cell.column for cell in ws[header_idx + 1]}
        ok_pic_col = col_map.get('OK Picture')
        ng_pic_col = col_map.get('NG Picture')
        excel_row = df_index + header_idx + 2
        
        # 找到图片表格
        pic_table = next((tbl for tbl in doc.tables if 'OK-Part' in tbl.cell(0, 0).text and 'Not-OK-Part' in tbl.cell(0, 1).text), None)
        if pic_table and len(pic_table.rows) > 1:
            ok_cell, ng_cell = pic_table.cell(1, 0), pic_table.cell(1, 1)
            ok_cell.paragraphs[0].clear()
            ng_cell.paragraphs[0].clear()
            
            for img in ws._images:
                if img.anchor._from.row + 1 == excel_row:
                    img_data = io.BytesIO(img.ref)
                    if img.anchor._from.col + 1 == ok_pic_col: ok_cell.paragraphs[0].add_run().add_picture(img_data, width=Inches(2.5))
                    elif img.anchor._from.col + 1 == ng_pic_col: ng_cell.paragraphs[0].add_run().add_picture(img_data, width=Inches(2.5))
    except Exception as e:
        st.warning(f"⚠️ 处理图片时发生错误: {e}")

    # 5. 精准定位正文主体内容并填充 (在原版基础上增加和修改映射)
    body_p_elements = doc._body._body.findall('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}p')
    
    # **关键修改**: 更新 headings_config 以匹配新需求
    headings_config = [
        {'keys': ['Failure Mode'], 'value': row_data.get('Failure Mode', '')},
        {'keys': ['Project/Part name'], 'value': row_data.get('Project/Part name', '')},
        {'keys': ['Process', 'Related Material Field / Process'], 'value': row_data.get('Related Material Field / Process', '')},
        {'keys': ['Problem (Fundamental Problem)'], 'value': row_data.get('LL Brief Description', '')},
        {'keys': ['Root Cause(s)', 'Root Cause'], 'value': row_data.get('Root Cause', '')},
        {'keys': ['Task', 'Task/Scope'], 'value': row_data.get('LL Supplier Scope', '')}, # 匹配 Task/Scope
        {'keys': ['Corrective Actions'], 'value': row_data.get('Corrective Action', '')},
    ]

    # **新增功能**: 拆分并填充 Should/Should not
    should_text = str(row_data.get('Should or not to do', ''))
    should_do = should_text.split('Should not:')[0].replace('Should:', '').strip()
    should_not_do = should_text.split('Should not:')[1].strip() if 'Should not:' in should_text else ''
    headings_config.extend([
        {'keys': ['What should we do in the future?'], 'value': should_do},
        {'keys': ['What should we not do in the future?'], 'value': should_not_do}
    ])

    p_indices = {}
    for idx, p_elem in enumerate(body_p_elements):
        try:
            p = Paragraph(p_elem, doc)
            p_text = p.text.strip()
            for i, config in enumerate(headings_config):
                for key in config['keys']:
                    if key in p_text and i not in p_indices:
                        p_indices[i] = idx
                        break
        except: pass

    # 逆序填充 (完全保留您原版的填充逻辑)
    for i, p_idx in sorted(p_indices.items(), key=lambda x: x[1], reverse=True):
        val = str(headings_config[i]['value'] if pd.notna(headings_config[i]['value']) else "")
        p = Paragraph(body_p_elements[p_idx], doc)
        new_p = p.insert_paragraph_before(val)
        # 尝试保持样式
        if p.style: new_p.style = p.style
                                
    return doc

# 核心逻辑：生成EML邮件 (原版逻辑)
def generate_eml_file(row_data):
    serial_no = str(row_data.get('LL Serials No', 'LL-xxxx-xx')).strip()
    failure_mode = str(row_data.get('Failure Mode', '******')).strip()
    subject = f"M/PQR-AP LL | {serial_no} | Title {failure_mode}"
    html_body = f"""
    <html>
    <head><style>body{{font-family:'Arial',sans-serif;font-size:10.5pt;line-height:1.6;color:#333;}}.red-bold{{color:#FF0000;font-weight:bold;}}ul{{margin:5px 0 15px 20px;}}li{{margin-bottom:8px;}}</style></head>
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

# 5. 页面渲染与交互 (在原版基础上适配新列)
if excel_file is not None and template_file is not None:
    try:
        df, sheet_name, header_idx = load_excel_robust(excel_file)
        st.success(f"🎉 成功加载工作表: **{sheet_name}** (表头定位自第 {header_idx + 1} 行)")
        
        # **关键修改**: 检查所有需要的列
        required_cols = ['Project/Part name', 'LL Brief Description', 'Root Cause', 'LL point', 'LL Supplier Scope', 'Failure Mode', 'Corrective Action', 'Should or not to do', 'Related Material Field / Process', 'OK Picture', 'NG Picture']
        missing_cols = [c for c in required_cols if c not in df.columns]
        
        if missing_cols:
            st.error(f"❌ Excel 中缺失生成所需的关键列: {', '.join(missing_cols)}")
        else:
            st.markdown("### 👈 第一步：在下方表格中搜索并勾选需要生成的 Record")
            search_term = st.text_input("🔍 快速搜索 (支持输入序列号、供应商、项目名称、失效模式进行实时筛选)：")
            
            filtered_df = df.copy()
            if search_term:
                filtered_df = df[df.astype(str).apply(lambda row: row.str.contains(search_term, case=False).any(), axis=1)]
            
            filtered_df.insert(0, '选择 (Select)', False)
            display_cols = ['选择 (Select)', 'LL Serials No', 'Failure Mode', 'Supplier Name', 'Project/Part name', 'LL Brief Description']
            available_display_cols = [c for c in display_cols if c in filtered_df.columns]
            
            edited_df = st.data_editor(
                filtered_df[available_display_cols],
                hide_index=True,
                column_config={"选择 (Select)": st.column_config.CheckboxColumn("选择", default=False)},
                disabled=[col for col in available_display_cols if col != '选择 (Select)'],
                use_container_width=True
            )
            
            # 使用 original index 进行选择
            selected_indices = edited_df[edited_df['选择 (Select)'] == True].index
            selected_rows = filtered_df.loc[selected_indices]
            
            st.write("---")
            st.markdown("### 👉 第二步：一键生成与下载")
            
            if len(selected_rows) > 0:
                st.success(f"已勾选 **{len(selected_rows)}** 条记录。请点击下方按钮开始生成：")
                
                if st.button("🚀 开始批量生成 Word 与邮件草稿", type="primary", use_container_width=True):
                    with st.spinner("正在填充模板与邮件，请稍候..."):
                        
                        # 单文件生成
                        if len(selected_rows) == 1:
                            original_index = selected_rows.index[0]
                            row = selected_rows.iloc[0]
                            
                            doc = fill_word_template(template_file, row.to_dict(), original_index, excel_file, header_idx)
                            bio_doc = io.BytesIO()
                            doc.save(bio_doc)
                            
                            serial_str = str(row.get('LL Serials No', 'LL-xxxx-xx'))
                            col1, col2 = st.columns(2)
                            col1.download_button(f"📥 下载 Word: LL_Template_{serial_str}.docx", bio_doc.getvalue(), f"LL_Template_{serial_str}.docx", use_container_width=True)
                            col2.download_button(f"📧 下载邮件草稿: Email_Draft_{serial_str}.eml", generate_eml_file(row.to_dict()), f"Email_Draft_{serial_str}.eml", use_container_width=True)
                            st.success("✨ Word 模板与 Outlook 邮件草稿生成成功！")
                            
                        # 多文件打包
                        else:
                            zip_buffer = io.BytesIO()
                            with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
                                for original_index, row in selected_rows.iterrows():
                                    doc = fill_word_template(template_file, row.to_dict(), original_index, excel_file, header_idx)
                                    doc_io = io.BytesIO()
                                    doc.save(doc_io)
                                    serial_str = str(row.get('LL Serials No', f"Record_{original_index}"))
                                    zf.writestr(f"LL_Template_{serial_str}.docx", doc_io.getvalue())
                                    zf.writestr(f"Email_Draft_{serial_str}.eml", generate_eml_file(row.to_dict()))
                                    
                            st.download_button("📥 立即下载打包好的 ZIP 压缩包", zip_buffer.getvalue(), "LL_Templates_and_Emails_Batch.zip", "application/zip", use_container_width=True)
                            st.success("✨ 批量生成成功！")
            else:
                st.warning("👉 请在上方表格的【选择】列中勾选您想要生成的记录。")
                
    except Exception as e:
        st.error(f"❌ 读取或处理文件时出错: {e}")
        st.info("排查提示：请确认 Excel 文件和 Word 模板未被独占打开，且所有必需列均存在。")
else:
    st.info("ℹ️ 请在侧边栏确认默认文件路径，或者直接在上方手动上传 Excel 数据源和 Word 模板。")












