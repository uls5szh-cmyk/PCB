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
    # 扩展支持上传 xlsm 格式文件
    uploaded_excel = st.sidebar.file_uploader("手动上传 PCB Master List (Excel):", type=["xlsx", "xlsm"])
    uploaded_template = st.sidebar.file_uploader("手动上传 LL Template (Word):", type=["docx"])
    if uploaded_excel:
        excel_file = uploaded_excel
    if uploaded_template:
        template_file = uploaded_template

# 如果没有找到默认文件，在主界面显示上传入口
if excel_file is None:
    # 扩展支持上传 xlsm 格式文件
    uploaded_excel = st.file_uploader("📁 请上传 PCB Lesson Learn Master List (Excel/xlsm)", type=["xlsx", "xlsm"], key="excel_main")
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
    
    # 查找目标工作表 "PUQ3 LL Overall list"
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
        
    # 读取前10行，智能判定表头所在行（表头从第二行开始，即 index=1）
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
    
    # 1. 自动获取当前日期并格式化（例如 "October 24 2024"）
    current_date_str = datetime.date.today().strftime('%B %d %Y')
    
    problem_text = row_data.get('LL Brief Description', '')
    problem_text_str = str(problem_text).strip() if (problem_text and not pd.isna(problem_text)) else ""
    
    ns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
    
    # 2. 全域 w:p 节点深度穿透扫描列表（打破表格与正文的隔离）
    all_p_elements = []
    all_p_elements.extend(doc._element.findall('.//w:p', ns))
    
    # 页眉与页脚中的所有段落节点
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
            
    # 去除重复节点
    all_p_elements = list(set(all_p_elements))
    
    # 3. 终极深度 XML 文本节点替换与追加
    for p_elem in all_p_elements:
        t_elems = p_elem.findall('.//w:t', ns)
        p_text = "".join([t.text for t in t_elems if t.text])
        p_text_normalized = " ".join(p_text.lower().split())
        
        # A. 全域无死角替换老日期
        if 'may 24 2022' in p_text_normalized:
            for t in t_elems:
                if t.text:
                    if 'May 24 2022' in t.text:
                        t.text = t.text.replace('May 24 2022', current_date_str)
                    if 'May 24, 2022' in t.text:
                        t.text = t.text.replace('May 24, 2022', current_date_str)
                        
        # B. 终极覆盖：追加 'Lesson Learn –' 标题内容（XML直接覆盖，完美规避格式碎块）
        if 'lesson' in p_text_normalized and 'learn' in p_text_normalized and len(p_text_normalized) < 50:
            if problem_text_str and problem_text_str not in p_text:
                p_text_clean = p_text.strip().rstrip("–-—: ").strip()
                new_val = f"{p_text_clean} – {problem_text_str}"
                
                if t_elems:
                    t_elems[0].text = new_val
                    # 将该段落内其他碎裂的子文字运行块清空，防止文字重复
                    for t in t_elems[1:]:
                        t.text = ""
                else:
                    r_elem = ET.SubElement(p_elem, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}r')
                    t_elem = ET.SubElement(r_elem, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t')
                    t_elem.text = new_val

    # 4. 精准定位正文主体内容 1 到 5 各部分标题配置并填入 (使用 Body 快速定位)
    body_p_elements = doc._body._body.findall('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}p')
    
    headings_config = [
        {
            'keys': ['Product / Process', '1 Product / Process', 'Product/Process'],
            'value': row_data.get('Project/Part name', '')
        },
        {
            'keys': ['Problem (Fundamental Problem)', '2 Problem (Fundamental Problem)', 'Problem'],
            'value': row_data.get('LL Brief Description', '')
        },
        {
            'keys': ['Root Cause(s)', '3 Root Cause(s)', 'Root Cause'],
            'value': row_data.get('Root Cause', '')
        },
        {
            'keys': ['Learning', '4 Learning', 'LL point'],
            'value': row_data.get('LL point', '')
        },
        {
            'keys': ['Task', '5 Task', 'LL Supplier Scope'],
            'value': row_data.get('LL Supplier Scope', '')
        }
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

    # 逆序填充
    sorted_configs = sorted(p_indices.items(), key=lambda x: x[1], reverse=True)
    
    for i, p_idx in sorted_configs:
        val = headings_config[i]['value']
        if not val or pd.isna(val):
            val = ""
            
        p_elem = body_p_elements[p_idx]
        try:
            p = Paragraph(p_elem, doc)
            replaced = False
            
            # 查找下一段看是否是占位行
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
                                
    return doc

# 核心逻辑：自动生成符合 Outlook 规范的 EML 邮件文件（带 HTML 红色加粗标记）
def generate_eml_file(row_data):
    serial_no = str(row_data.get('LL Serials No', 'LL-xxxx-xx')).strip()
    failure_mode = str(row_data.get('Failure Mode', '*****')).strip()
    
    # 拼接主题：M/PQR-AP LL | LL-xxxx-xx | Title *****
    subject = f"M/PQR-AP LL | {serial_no} | Title {failure_mode}"
    
    # 拼接富文本 HTML 正文排版 (还原截图，包含红色加粗)
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
    
    # 组装 EML 规范对象
    msg = MIMEMultipart('alternative')
    msg['Subject'] = Header(subject, 'utf-8')
    msg['From'] = 'Sunny.LIU3@cn.bosch.com'
    msg['To'] = ''  # 留空供用户手动填写
    
    # 附加 HTML 内容
    part = MIMEText(html_body, 'html', 'utf-8')
    msg.attach(part)
    
    return msg.as_bytes()

# 5. 页面渲染与交互
if excel_file is not None and template_file is not None:
    try:
        df, sheet_name, header_idx = load_excel_robust(excel_file)
        st.success(f"🎉 成功加载工作表: **{sheet_name}** (表头定位自第 {header_idx + 1} 行)")
        
        # 检查核心列
        required_cols = ['Project/Part name', 'LL Brief Description', 'Root Cause', 'LL point']
        missing_cols = [c for c in required_cols if c not in df.columns]
        
        # 智能检索并兼容 "LL Serials No" 和带点的 "LL Serials No."
        serial_no_col = 'LL Serials No'
        for col in df.columns:
            col_clean = str(col).strip().replace('.', '').lower()
            if 'll serials no' in col_clean or 'serials no' in col_clean or 'serial no' in col_clean:
                serial_no_col = col
                break
                
        # 匹配 Supplier Scope
        supplier_scope_col = 'LL Supplier Scope'
        if 'LL Supplier Scope' not in df.columns:
            potential = [c for c in df.columns if 'Supplier Scope' in c or 'Scope' in c or 'Task' in c]
            if potential:
                supplier_scope_col = potential[0]
                st.info(f"💡 'LL Supplier Scope' 自动匹配为 Excel 中的: '{supplier_scope_col}'")
            else:
                st.warning("⚠️ 未在 Excel 中检测到 'LL Supplier Scope'，生成时 Task 部分将保持空白。")
        
        if missing_cols:
            st.error(f"❌ Excel 中缺少生成所需的关键列: {', '.join(missing_cols)}")
        else:
            st.markdown("### 👈 第一步：在下方表格中搜索并勾选需要生成的 Record")
            search_term = st.text_input("🔍 快速搜索 (支持输入序列号、供应商、项目名称、失效模式进行实时筛选)：")
            
            filtered_df = df.copy()
            if search_term:
                filtered_df = df[df.astype(str).apply(lambda row: row.str.contains(search_term, case=False).any(), axis=1)]
            
            filtered_df.insert(0, '选择 (Select)', False)
            
            # 使用动态匹配的序列号列
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
                    with st.spinner("正在填充模板与邮件，请稍候..."):
                        
                        # 场景 A：仅选中一条
                        if len(selected_rows) == 1:
                            row = selected_rows.iloc[0]
                            # 【核心修复】：补全 row_data 中的所有属性，以便完美支持 Word 追加和 EML 生成
                            row_data = {
                                'LL Serials No': row.get(serial_no_col, 'LL-xxxx-xx'),
                                'Failure Mode': row.get('Failure Mode', '*****'),
                                'Project/Part name': row.get('Project/Part name', ''),
                                'LL Brief Description': row.get('LL Brief Description', ''),
                                'Root Cause': row.get('Root Cause', ''),
                                'LL point': row.get('LL point', ''),
                                'LL Supplier Scope': row.get(supplier_scope_col, '')
                            }
                            
                            # 生成 Word
                            doc = fill_word_template(template_file, row_data)
                            bio_doc = io.BytesIO()
                            doc.save(bio_doc)
                            bio_doc.seek(0)
                            
                            # 生成 Email
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
                            
                        # 场景 B：选中了多条（直接全部打包成一个 ZIP 包供用户一键下载）
                        else:
                            zip_buffer = io.BytesIO()
                            with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
                                for idx, (_, row) in enumerate(selected_rows.iterrows()):
                                    # 【核心修复】：补全批量 row_data 中的所有属性
                                    row_data = {
                                        'LL Serials No': row.get(serial_no_col, f"Record_{idx+1}"),
                                        'Failure Mode': row.get('Failure Mode', '*****'),
                                        'Project/Part name': row.get('Project/Part name', ''),
                                        'LL Brief Description': row.get('LL Brief Description', ''),
                                        'Root Cause': row.get('Root Cause', ''),
                                        'LL point': row.get('LL point', ''),
                                        'LL Supplier Scope': row.get(supplier_scope_col, '')
                                    }
                                    
                                    # 生成并打包 Word
                                    doc = fill_word_template(template_file, row_data)
                                    doc_io = io.BytesIO()
                                    doc.save(doc_io)
                                    doc_io.seek(0)
                                    
                                    serial_str = str(row_data['LL Serials No'])
                                    zip_file.writestr(f"LL_Template_{serial_str}.docx", doc_io.getvalue())
                                    
                                    # 生成并打包 Eml 邮件草稿
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

