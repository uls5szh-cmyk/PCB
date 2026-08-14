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
from docx.shared import Inches
import re

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

# 如果没有找到默认文件，在主界面显示上传入口
if excel_file is None:
    uploaded_excel = st.file_uploader("📁 请上传 PCB Lesson Learn Master List (Excel)", type=["xlsx"], key="excel_main")
    if uploaded_excel:
        excel_file = uploaded_excel

if template_file is None:
    uploaded_template = st.file_uploader("📝 请上传 LL Template (Word) 模板文件", type=["docx"], key="template_main")
    if uploaded_template:
        template_file = uploaded_template

# 核心逻辑：加载 Excel，智能识别 Sheet 和表头行
def load_excel_robust(file_source):
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

# 核心逻辑：强力填充 Word 模板 (解决表格、文本框嵌套与日期、标题死角)
def fill_word_template(template_source, row_data):
    doc = docx.Document(template_source)
    from docx.text.paragraph import Paragraph

    # 辅助函数：清空段落并设置新文本
    def set_paragraph_text(p, text):
        for r in p.runs:
            r.clear()
        p.add_run(str(text))

    # 1. 全局替换日期和主标题
    current_date_str = datetime.date.today().strftime('%B %d, %Y')
    brief_description = str(row_data.get('LL Brief Description', ''))

    # 遍历所有段落（包括页眉页脚）进行替换
    all_paragraphs = doc.paragraphs
    for section in doc.sections:
        for hf in [section.header, section.footer, section.first_page_header, section.first_page_footer, section.even_page_header, section.even_page_footer]:
            if hf:
                all_paragraphs.extend(hf.paragraphs)

    for p in all_paragraphs:
        # 使用内联替换，避免破坏格式
        inline = p.runs
        for i in range(len(inline)):
            if 'May 24, 2022' in inline[i].text:
                text = inline[i].text.replace('May 24, 2022', current_date_str)
                inline[i].text = text
        # 填充主标题
        if p.text.strip().startswith("Lesson Learn –"):
             if not p.text.strip().endswith(brief_description):
                p.add_run(f" {brief_description}")

    # 2. 拆分 "Should or not to do" 列的内容
    should_text = ""
    should_not_text = ""
    learning_text = str(row_data.get('Should or not to do', ''))
    should_match = re.search(r'Should:(.*?)(Should not:|$)', learning_text, re.DOTALL | re.IGNORECASE)
    should_not_match = re.search(r'Should not:(.*)', learning_text, re.DOTALL | re.IGNORECASE)
    if should_match:
        should_text = should_match.group(1).strip()
    if should_not_match:
        should_not_text = should_not_match.group(1).strip()

    # 3. 定义内容映射：标题 -> Excel数据
    content_map = {
        "1. Failure Mode": str(row_data.get('LL Supplier Scope', '')),
        "Failure Mode": str(row_data.get('Failure Mode', '')),
        "Project/Part name": str(row_data.get('Project/Part name', '')),
        "Processe": str(row_data.get('Related Material Field / Process', '')),
        "Problem (Fundamental Problem)": str(row_data.get('LL Brief Description', '')),
        "Root Cause(s)": str(row_data.get('Root Cause', '')),
        "Corrective Actions": str(row_data.get('Corrective Action', '')),
        "8.1 What should we do in the future?": should_text,
        "8.2 What should we not do in the future?": should_not_text,
    }

    # 4. 遍历正文段落，根据标题填充内容
    body_paragraphs = doc.paragraphs
    for i, p in enumerate(body_paragraphs):
        p_text_stripped = p.text.strip()
        for heading, content in content_map.items():
            if p_text_stripped.startswith(heading):
                # 找到标题后，填充其紧随的下一个段落
                if i + 1 < len(body_paragraphs):
                    next_p = body_paragraphs[i+1]
                    set_paragraph_text(next_p, content)
                break 

    # 5. 在表格中填充图片
    ok_image_path = row_data.get('OK Picture')
    ng_image_path = row_data.get('NG Picture')

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                cell_text = cell.text.strip()
                if cell_text == 'OK-Part':
                    cell.text = ''
                    p = cell.paragraphs[0]
                    p.alignment = docx.enum.text.WD_ALIGN_PARAGRAPH.CENTER
                    r = p.add_run()
                    if ok_image_path and os.path.exists(str(ok_image_path)):
                        try:
                            r.add_picture(str(ok_image_path), width=Inches(2.5))
                        except Exception as e:
                            set_paragraph_text(p, f"图片加载失败: {e}")
                    else:
                        set_paragraph_text(p, "图片路径无效或文件不存在")
                elif cell_text == 'Not-OK-Part':
                    cell.text = ''
                    p = cell.paragraphs[0]
                    p.alignment = docx.enum.text.WD_ALIGN_PARAGRAPH.CENTER
                    r = p.add_run()
                    if ng_image_path and os.path.exists(str(ng_image_path)):
                        try:
                            r.add_picture(str(ng_image_path), width=Inches(2.5))
                        except Exception as e:
                            set_paragraph_text(p, f"图片加载失败: {e}")
                    else:
                        set_paragraph_text(p, "图片路径无效或文件不存在")
                                
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
            body {{ font-family: 'Arial', sans-serif; font-size: 10.5pt; line-height: 1.6; color: #333333; }}
            .red-bold {{ color: #FF0000; font-weight: bold; }}
            ul {{ margin-top: 5px; margin-bottom: 15px; padding-left: 20px; }}
            li {{ margin-bottom: 8px; }}
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
        
        required_cols = [
            'Project/Part name', 'LL Brief Description', 'Root Cause', 'Corrective Action', 
            'Should or not to do', 'Failure Mode', 'Related Material Field / Process'
        ]
        missing_cols = [c for c in required_cols if c not in df.columns]
        
        supplier_scope_col = 'LL Supplier Scope'
        if 'LL Supplier Scope' not in df.columns:
            potential = [c for c in df.columns if 'Supplier Scope' in c or 'Scope' in c or 'Task' in c]
            if potential:
                supplier_scope_col = potential[0]
                st.info(f"💡 'LL Supplier Scope' 自动匹配为 Excel 中的: '{supplier_scope_col}'")
            else:
                st.warning("⚠️ 未在 Excel 中检测到 'LL Supplier Scope'，生成时 Task 部分将保持空白。")
        
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
                    with st.spinner("正在填充模板与邮件，请稍候..."):
                        
                        if len(selected_rows) == 1:
                            row = selected_rows.iloc[0]
                            row_data = {
                                'LL Serials No': row.get('LL Serials No', 'LL-xxxx-xx'),
                                'Failure Mode': row.get('Failure Mode', ''),
                                'Project/Part name': row.get('Project/Part name', ''),
                                'Related Material Field / Process': row.get('Related Material Field / Process', ''),
                                'LL Brief Description': row.get('LL Brief Description', ''),
                                'Root Cause': row.get('Root Cause', ''),
                                'Corrective Action': row.get('Corrective Action', ''),
                                'Should or not to do': row.get('Should or not to do', ''),
                                'OK Picture': row.get('OK Picture'),
                                'NG Picture': row.get('NG Picture'),
                                'LL Supplier Scope': row.get(supplier_scope_col, '')
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
                                for idx, (_, row) in enumerate(selected_rows.iterrows()):
                                    row_data = {
                                        'LL Serials No': row.get('LL Serials No', f"Record_{idx+1}"),
                                        'Failure Mode': row.get('Failure Mode', ''),
                                        'Project/Part name': row.get('Project/Part name', ''),
                                        'Related Material Field / Process': row.get('Related Material Field / Process', ''),
                                        'LL Brief Description': row.get('LL Brief Description', ''),
                                        'Root Cause': row.get('Root Cause', ''),
                                        'Corrective Action': row.get('Corrective Action', ''),
                                        'Should or not to do': row.get('Should or not to do', ''),
                                        'OK Picture': row.get('OK Picture'),
                                        'NG Picture': row.get('NG Picture'),
                                        'LL Supplier Scope': row.get(supplier_scope_col, '')
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
        st.info("排查提示：请确认 Excel 文件和 Word 模板未被本地 Excel/Word 软件独占打开，并检查Excel中的列名是否正确。")
else:
    st.info("ℹ️ 请在侧边栏确认默认文件路径，或者直接在侧边栏中手动上传 Excel 数据源 and Word 模板。")














