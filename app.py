# -*- coding: utf-8 -*-
"""
Bosch PCB Lesson Learn & Quality Automation Studio
Created: 2026
Author: Quality Engineering Team
"""

import streamlit as st
import pandas as pd
import docx
from docx.text.paragraph import Paragraph
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
import datetime
import io
import os
import zipfile
import re
import openpyxl
import xml.etree.ElementTree as ET
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from email.header import Header

# -----------------------------------------------------------------------------
# 1. 页面基本配置与博世工业风 UI 注入
# -----------------------------------------------------------------------------
st.set_page_config(
    page_title="Bosch | PCB Lesson Learn Quality Studio",
    layout="wide",
    page_icon="🔴"
)

# 注入博世品牌定制 CSS 样式
BOSCH_CSS = """
<style>
    /* 全局主色调与字体 */
    :root {
        --bosch-red: #E20015;
        --bosch-blue: #005691;
        --bosch-light-blue: #007BC0;
        --bosch-dark-gray: #1C2B39;
        --bosch-gray: #525F6B;
        --bosch-bg: #F4F6F8;
    }
    
    .main {
        background-color: var(--bosch-bg);
    }
    
    /* 顶部博世特征色带 */
    .bosch-header-bar {
        height: 6px;
        background: linear-gradient(90deg, #E20015 0%, #E20015 25%, #005691 25%, #005691 65%, #007BC0 65%, #007BC0 100%);
        margin-bottom: 20px;
        border-radius: 3px;
    }
    
    .bosch-title {
        color: #005691;
        font-family: 'Arial', sans-serif;
        font-weight: 700;
        margin-bottom: 5px;
    }
    
    .bosch-subtitle {
        color: #525F6B;
        font-size: 0.95rem;
        margin-bottom: 20px;
    }
    
    /* 卡片式容器 */
    .bosch-card {
        background: #FFFFFF;
        padding: 20px;
        border-radius: 8px;
        box-shadow: 0 2px 6px rgba(0,0,0,0.05);
        border-left: 4px solid #005691;
        margin-bottom: 20px;
    }
    
    /* 按钮定制 */
    .stButton>button {
        border-radius: 4px;
        font-weight: 600;
    }
</style>
<div class="bosch-header-bar"></div>
"""
st.markdown(BOSCH_CSS, unsafe_allow_html=True)

# 顶部 Header 区域
st.markdown("""
<div style="display: flex; align-items: center; justify-content: space-between;">
    <div>
        <h1 class="bosch-title">🔴 BOSCH | PCB Lesson Learn Management Studio</h1>
        <p class="bosch-subtitle">品质与供应商协同自动化系统 · Standardized FEBER Report & Outlook Automation</p>
    </div>
</div>
""", unsafe_allow_html=True)

# -----------------------------------------------------------------------------
# 2. 侧边栏路径与数据源配置
# -----------------------------------------------------------------------------
DEFAULT_EXCEL_PATH = r"G:\02_7_M-PQA-RBAC1\08_PQA_AE\09_PQA2\11_PCB\04_Lessons learn\PCB Lesson Learn Master List.xlsx"
DEFAULT_TEMPLATE_PATH = r"G:\02_7_M-PQA-RBAC1\08_PQA_AE\09_PQA2\11_PCB\04_Lessons learn\LL Template complete version.docx"
TEAMS_BOT_URL = "https://teams.microsoft.com/l/app/ffcadcc0-464f-4110-a065-0e3b4733baa9?source=bot-header-share-entrypoint"

st.sidebar.markdown("### ⚙️ 系统数据源配置")
use_local_paths = st.sidebar.checkbox("使用 G 盘本地固定路径", value=True)

excel_file = None
template_file = None

if use_local_paths:
    excel_path = st.sidebar.text_input("Excel Master List 路径:", DEFAULT_EXCEL_PATH)
    template_path = st.sidebar.text_input("Word Template 模板路径:", DEFAULT_TEMPLATE_PATH)
    
    if os.path.exists(excel_path):
        excel_file = excel_path
    else:
        st.sidebar.warning("⚠️ 未找到本地 Excel，请切换为手动上传。")
        
    if os.path.exists(template_path):
        template_file = template_path
    else:
        st.sidebar.warning("⚠️ 未找到本地 Word 模板，请手动上传。")
else:
    uploaded_excel = st.sidebar.file_uploader("上传 Master List (.xlsx / .xlsm):", type=["xlsx", "xlsm"])
    uploaded_template = st.sidebar.file_uploader("上传 Word 模板 (.docx):", type=["docx"])
    if uploaded_excel:
        excel_file = uploaded_excel
    if uploaded_template:
        template_file = uploaded_template

# 快捷直达卡片
st.sidebar.markdown("---")
st.sidebar.markdown(f"""
<div style="background: #F0F4F8; padding: 12px; border-radius: 6px; border: 1px solid #D0DCE5;">
    <strong style="color: #005691;">🤖 M-PU ChatGPT Bot 直达</strong><br>
    <span style="font-size: 0.8rem; color: #525F6B;">点击跳转至 MS Teams 专属机器人进行文本润色</span><br><br>
    <a href="{TEAMS_BOT_URL}" target="_blank" style="display: block; text-align: center; background: #005691; color: #fff; text-decoration: none; padding: 6px 12px; border-radius: 4px; font-size: 0.85rem; font-weight: bold;">🚀 打开 Teams AI Bot</a>
</div>
""", unsafe_allow_html=True)

# -----------------------------------------------------------------------------
# 3. 核心功能函数集合
# -----------------------------------------------------------------------------

def load_supplier_emails(file_source):
    """从 Vendor code sheet 解析供应商名称与邮箱映射"""
    try:
        if not isinstance(file_source, str):
            file_source.seek(0)
        xl = pd.ExcelFile(file_source)
        if 'Vendor code' not in xl.sheet_names:
            return {}
        df_vendor = pd.read_excel(file_source, sheet_name='Vendor code')
        df_vendor.columns = df_vendor.columns.astype(str).str.strip()
        
        if 'Supplier' in df_vendor.columns and 'Name' in df_vendor.columns:
            df_vendor['Supplier'] = df_vendor['Supplier'].ffill()
            supplier_dict = {}
            for _, row in df_vendor.iterrows():
                supplier = str(row['Supplier']).strip()
                name_val = str(row['Name'])
                if pd.isna(row['Supplier']) or supplier in ['nan', 'None']:
                    continue
                emails = re.findall(r'[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+', name_val)
                if emails:
                    if supplier not in supplier_dict:
                        supplier_dict[supplier] = set()
                    for email in emails:
                        supplier_dict[supplier].add(email)
            for k in supplier_dict:
                supplier_dict[k] = list(supplier_dict[k])
            return supplier_dict
    except Exception as e:
        print(f"Vendor code 解析失败: {e}")
    return {}

def get_images_for_row(file_source, sheet_name, header_idx, target_row_idx):
    """提取指定行的 OK/NG 图片二进制流"""
    try:
        if isinstance(file_source, str):
            wb = openpyxl.load_workbook(file_source, data_only=True)
        else:
            file_source.seek(0)
            wb = openpyxl.load_workbook(file_source, data_only=True)
        ws = wb[sheet_name]
        
        col_ng, col_ok = -1, -1
        for col_idx in range(1, ws.max_column + 1):
            val = ws.cell(row=header_idx + 1, column=col_idx).value
            if val:
                val_str = str(val).strip()
                if 'NG Picture' in val_str: col_ng = col_idx - 1
                if 'OK Picture' in val_str: col_ok = col_idx - 1
                
        excel_target_row = header_idx + 1 + target_row_idx
        ok_img, ng_img = None, None
        
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
    except Exception:
        return None, None

def load_excel_robust(file_source):
    """智能识别主表并过滤 LL Need or not == Y"""
    if not isinstance(file_source, str):
        file_source.seek(0)
    xl = pd.ExcelFile(file_source)
    sheet_names = xl.sheet_names
    
    target_sheet = next((s for s in sheet_names if "PUQ3 LL Overall list" in s), None)
    if not target_sheet:
        target_sheet = next((s for s in sheet_names if "Overall" in s or "LL" in s), sheet_names[0])
        
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

def fill_word_template(template_source, row_data):
    """强力填充 Word 模板（彻底清除排版杂质）"""
    doc = docx.Document(template_source)
    current_date_str = datetime.date.today().strftime('%b %d %Y')
    failure_mode_str = str(row_data.get('Failure Mode', '')).strip()
    ns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
    
    # 1. 扫描全部段落并更新标题与日期
    all_p_elements = list(set(doc._element.findall('.//w:p', ns)))
    for section in doc.sections:
        for hf in [section.header, section.footer, section.first_page_header, section.first_page_footer]:
            if hf:
                all_p_elements.extend(hf._element.findall('.//w:p', ns))
    all_p_elements = list(set(all_p_elements))
    
    for p_elem in all_p_elements:
        t_elems = p_elem.findall('.//w:t', ns)
        if not t_elems:
            continue
        p_text = "".join([t.text for t in t_elems if t.text])
        p_text_lower = p_text.lower()
        
        # 替换老旧日期
        if 'may 24 2022' in p_text_lower:
            new_text = p_text.replace('May 24 2022', current_date_str).replace('May 24, 2022', current_date_str)
            t_elems[0].text = new_text
            for t in t_elems[1:]:
                t.text = ""
                
        # 追加主标题失效模式
        if 'lesson learn' in p_text_lower and len(p_text_lower) < 60:
            if failure_mode_str and failure_mode_str not in p_text:
                p_text_clean = p_text.strip().rstrip("–-—: ").strip()
                t_elems[0].text = f"{p_text_clean} – {failure_mode_str}"
                for t in t_elems[1:]:
                    t.text = ""

    # 2. 智能切分 Should / Should not
    should_or_not = str(row_data.get('Should or not to do', ''))
    should_text, should_not_text = "", ""
    if 'Should not:' in should_or_not:
        parts = should_or_not.split('Should not:')
        should_text = parts[0].replace('Should:', '').strip()
        should_not_text = parts[1].strip()
    else:
        should_text = should_or_not.replace('Should:', '').strip()

    # 3. 各章节标题及对应值映射
    headings_config = [
        {'keys': ['Task/Scope', 'Task'], 'value': row_data.get('LL Supplier Scope', '')},
        {'keys': ['Failure Mode'], 'value': row_data.get('Failure Mode', '')},
        {'keys': ['Project/Part name', 'Product / Process'], 'value': row_data.get('Project/Part name', '')},
        {'keys': ['Process'], 'value': row_data.get('Related Material Field / Process', '')},
        {'keys': ['Problem (Fundamental Problem)', 'Problem'], 'value': row_data.get('LL Brief Description', '')},
        {'keys': ['Root Cause(s)', 'Root Cause'], 'value': row_data.get('Root Cause', '')},
        {'keys': ['Corrective Actions', 'Corrective Action'], 'value': row_data.get('Corrective Action', '')},
        {'keys': ['What should we do in the future?'], 'value': should_text},
        {'keys': ['What should we not do in the future?'], 'value': should_not_text},
        {'keys': ['What else could be additionally affected?'], 'value': row_data.get('What else could be additionally affected?', '')},
        {'keys': ['Where can the problem additionally occur?'], 'value': row_data.get('Where can the problem additionally occur?', '')},
        {'keys': ['When can the problem additionally appear?'], 'value': row_data.get('When can the problem additionally appear?', '')},
        {'keys': ['Who else can be affected?'], 'value': row_data.get('Who else can be affected?', '')}
    ]
    
    body_p_elements = doc._body._body.findall('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}p')
    p_indices = {}
    for idx, p_elem in enumerate(body_p_elements):
        try:
            p_text = Paragraph(p_elem, doc).text.strip()
            for i, config in enumerate(headings_config):
                for key in config['keys']:
                    if key in p_text and i not in p_indices:
                        p_indices[i] = idx
                        break
        except Exception:
            pass
            
    # 逆序填充正文并强制重置段落格式
    sorted_configs = sorted(p_indices.items(), key=lambda x: x[1], reverse=True)
    for i, p_idx in sorted_configs:
        val = headings_config[i]['value']
        if not val or pd.isna(val):
            val = ""
            
        p_elem = body_p_elements[p_idx]
        try:
            p = Paragraph(p_elem, doc)
            replaced = False
            ans_p = None
            
            if p_idx + 1 < len(body_p_elements):
                next_p = Paragraph(body_p_elements[p_idx + 1], doc)
                next_text = next_p.text.strip()
                if next_text in ['...', '…', ''] or len(next_text) < 5:
                    is_another_heading = any(any(k in next_text for k in c['keys']) for c in headings_config)
                    if not is_another_heading:
                        ans_p = next_p
                        replaced = True
                        
            if not replaced:
                new_p_element = p._element.getparent().create_element('w:p')
                p._element.addnext(new_p_element)
                ans_p = Paragraph(new_p_element, p._parent)
                
            if ans_p is not None:
                # 彻底消除原有段落可能残留的缩进、居中和行距杂质
                if ans_p._element.pPr is not None:
                    ans_p._element.remove(ans_p._element.pPr)
                ans_p.text = ""
                run = ans_p.add_run(str(val))
                run.font.name = 'Arial'
                run.font.size = Pt(10.5)
                run.font.bold = False
        except Exception:
            pass
            
    # 4. 图片填充逻辑 (自适应宽度，清理占位符)
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
            
            # 处理 OK-Part
            if ok_img and ok_cell:
                inserted = False
                for p in list(ok_cell.paragraphs):
                    if "ok-part" in p.text.lower().replace(" ", ""):
                        continue
                    if not inserted:
                        if p._element.pPr is not None:
                            p._element.remove(p._element.pPr)
                        p.text = ""
                        p.alignment = 1
                        r = p.add_run()
                        r.add_picture(io.BytesIO(ok_img), width=Inches(2.5))
                        inserted = True
                    else:
                        p._element.getparent().remove(p._element)
                if not inserted:
                    p = ok_cell.add_paragraph()
                    p.alignment = 1
                    p.add_run().add_picture(io.BytesIO(ok_img), width=Inches(2.5))
                    
            # 处理 Not-OK-Part
            if ng_img and ng_cell:
                inserted = False
                for p in list(ng_cell.paragraphs):
                    if "not-ok-part" in p.text.lower().replace(" ", ""):
                        continue
                    if not inserted:
                        if p._element.pPr is not None:
                            p._element.remove(p._element.pPr)
                        p.text = ""
                        p.alignment = 1
                        r = p.add_run()
                        r.add_picture(io.BytesIO(ng_img), width=Inches(2.5))
                        inserted = True
                    else:
                        p._element.getparent().remove(p._element)
                if not inserted:
                    p = ng_cell.add_paragraph()
                    p.alignment = 1
                    p.add_run().add_picture(io.BytesIO(ng_img), width=Inches(2.5))
                    
    return doc

def generate_eml_file(row_data, to_emails="", doc_bytes=None, doc_filename="LL_Template.docx"):
    """生成内置编辑模式（X-Unsent: 1）且带 Word 附件的 Outlook 草稿"""
    serial_no = str(row_data.get('LL Serials No', 'LL-xxxx-xx')).strip()
    failure_mode = str(row_data.get('Failure Mode', '*****')).strip()
    subject = f"M/PQR-AP LL | {serial_no} | Title {failure_mode}"
    
    html_body = f"""
    <html>
    <head>
        <style>
            body {{ font-family: 'Arial', sans-serif; font-size: 10.5pt; line-height: 1.6; color: #333333; }}
            .red-bold {{ color: #E20015; font-weight: bold; }}
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
    msg = MIMEMultipart('mixed')
    msg['Subject'] = Header(subject, 'utf-8')
    msg['From'] = 'Sunny.LIU3@cn.bosch.com'
    msg['To'] = to_emails
    msg.add_header('X-Unsent', '1') # 设为草稿编辑模式
    
    # 附加正文
    alt_part = MIMEMultipart('alternative')
    alt_part.attach(MIMEText(html_body, 'html', 'utf-8'))
    msg.attach(alt_part)
    
    # 附加生成的 Word 文档
    if doc_bytes:
        part_doc = MIMEBase('application', 'vnd.openxmlformats-officedocument.wordprocessingml.document')
        part_doc.set_payload(doc_bytes)
        encoders.encode_base64(part_doc)
        part_doc.add_header('Content-Disposition', f'attachment; filename="{doc_filename}"')
        msg.attach(part_doc)
        
    return msg.as_bytes()

def parse_ai_generated_text(ai_text):
    """智能解析 M-PU ChatGPT Bot 生成的标准结构化长文本"""
    data = {
        'LL Serials No': 'LL-AI-Generated',
        'Failure Mode': '',
        'Project/Part name': '',
        'LL Brief Description': '',
        'Root Cause': '',
        'Corrective Action': '',
        'Should or not to do': '',
        'Related Material Field / Process': '',
        'LL Supplier Scope': '',
        'What else could be additionally affected?': '',
        'Where can the problem additionally occur?': '',
        'When can the problem additionally appear?': '',
        'Who else can be affected?': ''
    }
    
    # 正则提取 Failure Mode
    fm_match = re.search(r'(?:Failure Mode|Issue|Problem):\s*(.+)', ai_text, re.I)
    if fm_match:
        data['Failure Mode'] = fm_match.group(1).split('\n')[0].strip()
        
    # 正则提取 1. Product / Process
    p_match = re.search(r'1\.\s*Product\s*/\s*Process\s*([\s\S]*?)(?=2\.\s*Problem|$)', ai_text, re.I)
    if p_match:
        data['Project/Part name'] = p_match.group(1).strip()
        
    # 正则提取 2. Problem
    prob_match = re.search(r'2\.\s*Problem[^\n]*\n([\s\S]*?)(?=3\.\s*Lessons|$)', ai_text, re.I)
    if prob_match:
        data['LL Brief Description'] = prob_match.group(1).strip()
        
    # 正则提取 3. Lessons
    lessons_match = re.search(r'3\.\s*Lessons[^\n]*\n([\s\S]*?)(?=4\.\s*Potentially|$)', ai_text, re.I)
    if lessons_match:
        content = lessons_match.group(1).strip()
        data['Root Cause'] = content
        data['Corrective Action'] = content
        data['Should or not to do'] = content
        
    # 正则提取 4. Potentially affected
    pot_match = re.search(r'4\.\s*Potentially affected[^\n]*\n([\s\S]*?)(?=5\.\s*Appendix|$)', ai_text, re.I)
    if pot_match:
        pot_text = pot_match.group(1)
        w1 = re.search(r'What else[^\n]*\n([^\n]+)', pot_text, re.I)
        w2 = re.search(r'Where can[^\n]*\n([^\n]+)', pot_text, re.I)
        w3 = re.search(r'When can[^\n]*\n([^\n]+)', pot_text, re.I)
        w4 = re.search(r'Who else[^\n]*\n([^\n]+)', pot_text, re.I)
        if w1: data['What else could be additionally affected?'] = w1.group(1).strip()
        if w2: data['Where can the problem additionally occur?'] = w2.group(1).strip()
        if w3: data['When can the problem additionally appear?'] = w3.group(1).strip()
        if w4: data['Who else can be affected?'] = w4.group(1).strip()
        
    return data

# -----------------------------------------------------------------------------
# 4. 主界面：多标签页设计 (两大工作流)
# -----------------------------------------------------------------------------
tab1, tab2, tab3 = st.tabs([
    "📊 引擎 ①：Master List 批量生成与分发",
    "🤖 引擎 ②：M-PU AI 文本直接转报告",
    "📖 标准 FEBER 结构与 Prompt 指南"
])

# =============================================================================
# TAB 1: 传统 Master List 流程
# =============================================================================
with tab1:
    if excel_file is not None and template_file is not None:
        try:
            df, sheet_name, header_idx = load_excel_robust(excel_file)
            supplier_dict = load_supplier_emails(excel_file)
            
            # 过滤 LL Need or not
            ll_need_col = 'LL Need or not'
            if ll_need_col not in df.columns:
                for col in df.columns:
                    if 'need or not' in str(col).lower():
                        ll_need_col = col
                        break
            if ll_need_col in df.columns:
                orig_len = len(df)
                df = df[df[ll_need_col].astype(str).str.strip().str.upper() == 'Y']
                st.success(f"🎉 成功载入 **{sheet_name}**：已过滤保留 `{ll_need_col} = 'Y'` 的 **{len(df)}** 条有效记录（排除 {orig_len - len(df)} 条）。")
                
            # 序列号和 Supplier Scope 列定位
            serial_no_col = next((c for c in df.columns if 'serial' in str(c).lower()), 'LL Serials No')
            supplier_scope_col = next((c for c in df.columns if 'scope' in str(c).lower() or 'task' in str(c).lower()), 'LL Supplier Scope')
            
            st.markdown("#### 1️⃣ 勾选需要生成的 Lessons Learned 记录")
            search_term = st.text_input("🔍 关键字快速过滤 (支持序列号/供应商/失效模式/项目):", key="search_tab1")
            filtered_df = df.copy()
            if search_term:
                filtered_df = df[df.astype(str).apply(lambda row: row.str.contains(search_term, case=False).any(), axis=1)]
                
            filtered_df.insert(0, '选择 (Select)', False)
            disp_cols = ['选择 (Select)', serial_no_col, 'Failure Mode', 'Supplier Name', 'Project/Part name', 'LL Brief Description']
            valid_disp_cols = [c for c in disp_cols if c in filtered_df.columns]
            
            edited_df = st.data_editor(
                filtered_df[valid_disp_cols],
                hide_index=True,
                column_config={"选择 (Select)": st.column_config.CheckboxColumn("选择", default=False)},
                disabled=[c for c in valid_disp_cols if c != '选择 (Select)'],
                use_container_width=True
            )
            
            selected_rows = filtered_df.loc[edited_df[edited_df['选择 (Select)'] == True].index]
            
            # 供应商与生成控制台
            st.markdown("#### 2️⃣ 选择收件供应商并一键生成")
            col_sup, col_info = st.columns([2, 2])
            with col_sup:
                selected_suppliers = st.multiselect("👥 指定接收此 LL 的供应商 (自动提取邮箱):", options=list(supplier_dict.keys()))
            
            to_emails_list = []
            for s in selected_suppliers:
                to_emails_list.extend(supplier_dict[s])
            to_emails_str = "; ".join(list(set(to_emails_list)))
            
            with col_info:
                if to_emails_str:
                    st.info(f"📧 **收件人:** `{to_emails_str}`")
                else:
                    st.caption("ℹ️ 若未指定供应商，生成的邮件收件人栏将保持空白供手动填写。")
                    
            if len(selected_rows) > 0:
                st.write("")
                if st.button("🚀 开始批量生成 Word 报告与邮件草稿", type="primary", use_container_width=True):
                    with st.spinner("正在抽取图片、规范排版并打包附件..."):
                        if len(selected_rows) == 1:
                            row = selected_rows.iloc[0]
                            ok_img, ng_img = get_images_for_row(excel_file, sheet_name, header_idx, row.name)
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
                                'What else could be additionally affected?': row.get('What else could be additionally affected?', ''),
                                'Where can the problem additionally occur?': row.get('Where can the problem additionally occur?', ''),
                                'When can the problem additionally appear?': row.get('When can the problem additionally appear?', ''),
                                'Who else can be affected?': row.get('Who else can be affected?', ''),
                                'OK Picture Bytes': ok_img,
                                'NG Picture Bytes': ng_img
                            }
                            doc = fill_word_template(template_file, row_data)
                            bio_doc = io.BytesIO()
                            doc.save(bio_doc)
                            doc_bytes = bio_doc.getvalue()
                            serial_str = str(row_data['LL Serials No'])
                            doc_filename = f"LL_Template_{serial_str}.docx"
                            eml_data = generate_eml_file(row_data, to_emails_str, doc_bytes, doc_filename)
                            
                            c1, c2 = st.columns(2)
                            with c1:
                                st.download_button(f"📥 下载 Word: {doc_filename}", doc_bytes, doc_filename, mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", use_container_width=True)
                            with c2:
                                st.download_button(f"📧 下载 Outlook 草稿 (含Word附件): Email_Draft_{serial_str}.eml", eml_data, f"Email_Draft_{serial_str}.eml", mime="message/rfc822", use_container_width=True)
                            st.success("✨ 生成成功！Word 报告已自动嵌入为邮件附件。")
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
                                        'What else could be additionally affected?': row.get('What else could be additionally affected?', ''),
                                        'Where can the problem additionally occur?': row.get('Where can the problem additionally occur?', ''),
                                        'When can the problem additionally appear?': row.get('When can the problem additionally appear?', ''),
                                        'Who else can be affected?': row.get('Who else can be affected?', ''),
                                        'OK Picture Bytes': ok_img,
                                        'NG Picture Bytes': ng_img
                                    }
                                    doc = fill_word_template(template_file, row_data)
                                    doc_io = io.BytesIO()
                                    doc.save(doc_io)
                                    doc_bytes = doc_io.getvalue()
                                    serial_str = str(row_data['LL Serials No'])
                                    doc_filename = f"LL_Template_{serial_str}.docx"
                                    zip_file.writestr(doc_filename, doc_bytes)
                                    eml_data = generate_eml_file(row_data, to_emails_str, doc_bytes, doc_filename)
                                    zip_file.writestr(f"Email_Draft_{serial_str}.eml", eml_data)
                                    
                            zip_buffer.seek(0)
                            st.download_button("📥 立即下载打包 ZIP (包含所有 Word 与带附件的 Outlook 草稿)", zip_buffer.read(), "LL_Batch_Package.zip", mime="application/zip", use_container_width=True)
                            st.success("✨ 批量打包成功！")
            else:
                st.warning("👉 请在上方表格中勾选需要处理的行。")
        except Exception as e:
            st.error(f"❌ 处理异常: {e}")
    else:
        st.info("ℹ️ 请在侧边栏确认 Excel 与 Word 模板文件路径。")

# =============================================================================
# TAB 2: AI 文本直接转报告 (免回填 Excel 流程)
# =============================================================================
with tab2:
    st.markdown("""
    <div class="bosch-card">
        <h4 style="color: #005691; margin-top:0;">🤖 M-PU ChatGPT Bot 文本即时转文档引擎</h4>
        <p style="font-size: 0.9rem; color: #525F6B;">
            当您在 Teams 中使用 <strong>M-PU ChatGPT Bot</strong> 润色完成之后，无需再繁琐地把文字复制回 Excel。只需将 Bot 输出的 Markdown/英文长文本粘贴在下方，程序会自动解析各章节并直接生成 Word 和邮件！
        </p>
    </div>
    """, unsafe_allow_html=True)
    
    col_input, col_config = st.columns([3, 2])
    with col_input:
        ai_pasted_text = st.text_area(
            "📋 在此粘贴 M-PU ChatGPT Bot 输出的完整文本：",
            height=320,
            placeholder="粘贴包含 0. Abstract, 1. Product / Process, 2. Problem, 3. Lessons, 4. Potentially affected 的完整回复..."
        )
        
    with col_config:
        st.markdown("##### ⚙️ 报告附加属性")
        ai_serial_no = st.text_input("报告编号 (LL Serials No.):", value=f"LL-{datetime.date.today().strftime('%Y%m%d')}-01")
        ai_failure_mode = st.text_input("失效模式 (若不填则自动从文本提取):", placeholder="例如: Hole copper thickness out of spec")
        
        # 供应商选择
        supplier_dict_ai = load_supplier_emails(excel_file) if excel_file else {}
        selected_ai_suppliers = st.multiselect("👥 发送给哪些供应商:", options=list(supplier_dict_ai.keys()), key="ai_sup_select")
        ai_emails_list = []
        for s in selected_ai_suppliers:
            ai_emails_list.extend(supplier_dict_ai[s])
        ai_to_emails_str = "; ".join(list(set(ai_emails_list)))
        if ai_to_emails_str:
            st.caption(f"📧 自动收件人: `{ai_to_emails_str}`")
            
    if st.button("⚡ 立即将 AI 输出转为 Word 报告 & Outlook 草稿", type="primary", use_container_width=True):
        if not ai_pasted_text.strip():
            st.warning("⚠️ 请先在文本框中粘贴 AI 的输出内容！")
        elif template_file is None:
            st.error("❌ 未检测到 Word 模板，请在侧边栏配置模板路径。")
        else:
            with st.spinner("正在进行语义解析与版面组装..."):
                parsed_data = parse_ai_generated_text(ai_pasted_text)
                parsed_data['LL Serials No'] = ai_serial_no
                if ai_failure_mode.strip():
                    parsed_data['Failure Mode'] = ai_failure_mode.strip()
                    
                doc = fill_word_template(template_file, parsed_data)
                bio_doc = io.BytesIO()
                doc.save(bio_doc)
                doc_bytes = bio_doc.getvalue()
                
                doc_filename = f"LL_Template_{ai_serial_no}.docx"
                eml_data = generate_eml_file(parsed_data, ai_to_emails_str, doc_bytes, doc_filename)
                
                st.success("🎉 解析并生成成功！")
                ca1, ca2 = st.columns(2)
                with ca1:
                    st.download_button(f"📥 下载 Word 报告: {doc_filename}", doc_bytes, doc_filename, mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", use_container_width=True)
                with ca2:
                    st.download_button(f"📧 下载 Outlook 邮件草稿 (含附件): Email_Draft_{ai_serial_no}.eml", eml_data, f"Email_Draft_{ai_serial_no}.eml", mime="message/rfc822", use_container_width=True)

# =============================================================================
# TAB 3: 标准 Prompt 与 FEBER 结构指南
# =============================================================================
with tab3:
    st.markdown("""
    ### 📖 博世标准 Lessons Learned (FEBER) 规范与 Prompt 模板
    
    当您向 **M-PU ChatGPT Bot** 发送请求时，可以直接使用以下标准化指令（已为您配置好博世要求）：
    """)
    
    standard_prompt_text = """Please create me a short and precise lessons learned report out of the attached document in American English.
You are an honest engineer; you provide always links to the sources and name the original slide/page number.
Please stick to the facts. In case you have additional topics, supporting or additional useful information be creative, add them and highlight them in italic.

Please write the headings in bold. Use key words that are understood by others in Bosch. Describe the report "user-friendly", so others can read it easily. One page for chapter 1-4 is appropriate. Delete all hints in blue letters.

If you are asked to create a lesson learned report, or to search for a lessons learned report, structure the answer as follows:
0. Abstract - write a short summary of the report with the structure - issue; problem; learnings; tags

Abstract
Issue: Describe briefly. Do not use abbreviations that are not commonly known.
Problem: Briefly describe the main problem using key words.
Lessons: Concentrate on the actual lessons learned. What is new? How can we prevent this issue in the future?

1. Product / Process
Product / Process:
Component:
Sub-Component:

2. Problem (Fundamental Problem)
Briefly describe the fundamental problem.

3. Lessons
- Root Cause: Describe the main root cause only.
- Measures & Sustainable Solutions: Focus on Process Improvements and Standards Updates.

4. Potentially affected
- What else could be additionally affected?
- Where can the problem additionally occur?
- When can the problem additionally appear?
- Who else can be affected?

5. Appendix (Optional)"""
    
    st.code(standard_prompt_text, language="text")
    st.markdown(f"""
    👉 **操作建议：**
    1. 点击上方代码块右上角的 **Copy** 按钮复制 Prompt；
    2. 粘贴至 <a href="{TEAMS_BOT_URL}" target="_blank">M-PU ChatGPT Bot (Teams)</a> 并附带您的原始信息；
    3. 获取 Bot 输出后，直接切换至本工具的 **【🤖 引擎 ②】** 粘贴即可瞬间出单！
    """, unsafe_allow_html=True)
