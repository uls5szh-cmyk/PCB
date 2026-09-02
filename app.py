# -*- coding: utf-8 -*-
"""
=============================================================================
BOSCH | PCB Lesson Learn Quality Studio (Standard FEBER Table Edition)
- 1:1 Markdown Table Mirror Prompt Generator
- Table-to-Table Safe Direct Population (3-Column Lessons & 4-Row Potentially Affected)
- Auto-Fit Picture Insertion & Recipient-Ready Outlook EML Draft
=============================================================================
"""

import streamlit as st
import pandas as pd
import docx
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
import datetime
import io
import os
import zipfile
import re
import openpyxl
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from email.header import Header

# -----------------------------------------------------------------------------
# 1. 页面基本配置与博世工业视觉体系 (Bosch Corporate Identity)
# -----------------------------------------------------------------------------
st.set_page_config(
    page_title="Bosch | PCB Lesson Learn Quality Studio",
    layout="wide",
    page_icon="🔴"
)

BOSCH_UI_STYLE = """
<style>
    :root {
        --bosch-red: #E20015;
        --bosch-blue: #005691;
        --bosch-light-blue: #007BC0;
        --bosch-dark-gray: #1C2B39;
        --bosch-gray: #525F6B;
        --bosch-bg: #F4F6F8;
    }
    .stApp { background-color: var(--bosch-bg); }
    .bosch-top-bar {
        height: 6px;
        background: linear-gradient(90deg, #E20015 0%, #E20015 25%, #005691 25%, #005691 65%, #007BC0 65%, #007BC0 100%);
        border-radius: 3px;
        margin-bottom: 20px;
    }
    .stButton>button {
        background-color: var(--bosch-blue);
        color: white;
        border-radius: 4px;
        font-weight: 600;
        border: none;
    }
    .stButton>button:hover {
        background-color: var(--bosch-light-blue);
        color: white;
    }
</style>
<div class="bosch-top-bar"></div>
"""
st.markdown(BOSCH_UI_STYLE, unsafe_allow_html=True)

st.markdown("""
<div style="display: flex; align-items: center; justify-content: space-between; margin-bottom: 15px;">
    <div>
        <h2 style="color: #005691; margin-bottom: 2px;">🔴 BOSCH | PCB Lesson Learn 协同工作台</h2>
        <p style="color: #525F6B; font-size: 0.9rem; margin: 0;">FEBER 表格 1:1 镜像对齐 · 3列表格动态增行 · 4项评估表格精准回填 · 邮件草稿一键闭环</p>
    </div>
</div>
""", unsafe_allow_html=True)

# -----------------------------------------------------------------------------
# 2. 侧边栏路径配置
# -----------------------------------------------------------------------------
DEFAULT_EXCEL_PATH = r"G:\02_7_M-PQA-RBAC1\08_PQA_AE\09_PQA2\11_PCB\04_Lessons learn\PCB Lesson Learn Master List.xlsx"
DEFAULT_TEMPLATE_PATH = r"G:\02_7_M-PQA-RBAC1\08_PQA_AE\09_PQA2\11_PCB\04_Lessons learn\LL Template complete version.docx"
TEAMS_BOT_URL = "https://teams.microsoft.com/l/app/ffcadcc0-464f-4110-a065-0e3b4733baa9?source=bot-header-share-entrypoint"

st.sidebar.markdown("### ⚙️ 数据源路径配置")
use_local = st.sidebar.checkbox("使用本地固定路径 (G 盘)", value=True)

excel_file = None
template_file = None

if use_local:
    excel_path = st.sidebar.text_input("Master List 路径:", DEFAULT_EXCEL_PATH)
    template_path = st.sidebar.text_input("Word Template 路径:", DEFAULT_TEMPLATE_PATH)
    if os.path.exists(excel_path): excel_file = excel_path
    else: st.sidebar.warning("⚠️ 未在指定路径找到 Excel，请手动上传。")
    if os.path.exists(template_path): template_file = template_path
    else: st.sidebar.warning("⚠️ 未在指定路径找到 Word 模板，请手动上传。")
else:
    up_excel = st.sidebar.file_uploader("上传 Master List (Excel):", type=["xlsx", "xlsm"])
    up_template = st.sidebar.file_uploader("上传 Word 模板 (.docx):", type=["docx"])
    if up_excel: excel_file = up_excel
    if up_template: template_file = up_template

# -----------------------------------------------------------------------------
# 3. 辅助解析函数
# -----------------------------------------------------------------------------

def load_supplier_emails(file_source):
    """从 Vendor code 表中读取供应商与邮箱映射"""
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
    except Exception:
        pass
    return {}

def get_images_for_row(file_source, sheet_name, header_idx, target_row_idx):
    """从 Excel 指定行提取 OK 和 NG 图片二进制流"""
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
    """加载 Excel 并定位表头"""
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

# -----------------------------------------------------------------------------
# 4. 深度解析 Bot 回复（支持多行 Markdown 表格与 4 问结构）
# -----------------------------------------------------------------------------

def parse_bot_feber_response(bot_text):
    """解析 Bot 按照 FEBER 规范输出的结构化文本与表格"""
    parsed = {
        'Abstract_Issue': '',
        'Abstract_Problem': '',
        'Abstract_Lessons': '',
        'Product_Process': '',
        'Component': '',
        'Sub_Component': '',
        'Problem': '',
        'Lessons_Rows': [],  # list of tuples: (Lesson, Measures, RootCause)
        'What_Else': '',
        'Where': '',
        'When': '',
        'Who': ''
    }
    
    # 0. Abstract
    m_abs = re.search(r'(?:0\.\s*Abstract|Abstract)\s*([\s\S]*?)(?=1\.\s*Product|$)', bot_text, re.I)
    if m_abs:
        t = m_abs.group(1)
        i_m = re.search(r'Issue:\s*([^\n]+)', t, re.I)
        p_m = re.search(r'Problem:\s*([^\n]+)', t, re.I)
        l_m = re.search(r'Lessons:\s*([^\n]+)', t, re.I)
        if i_m: parsed['Abstract_Issue'] = i_m.group(1).strip()
        if p_m: parsed['Abstract_Problem'] = p_m.group(1).strip()
        if l_m: parsed['Abstract_Lessons'] = l_m.group(1).strip()
    
    # 1. Product / Process
    m_p = re.search(r'1\.\s*Product\s*/\s*Process\s*([\s\S]*?)(?=2\.\s*Problem|$)', bot_text, re.I)
    if m_p:
        t = m_p.group(1)
        pp = re.search(r'Product\s*/\s*Process:\s*([^\n]+)', t, re.I)
        cp = re.search(r'Component:\s*([^\n]+)', t, re.I)
        sc = re.search(r'Sub-Component:\s*([^\n]+)', t, re.I)
        if pp: parsed['Product_Process'] = pp.group(1).strip()
        if cp: parsed['Component'] = cp.group(1).strip()
        if sc: parsed['Sub_Component'] = sc.group(1).strip()
        
    # 2. Problem (Fundamental Problem)
    m_prob = re.search(r'2\.\s*Problem[^\n]*\n([\s\S]*?)(?=3\.\s*Lessons|$)', bot_text, re.I)
    if m_prob:
        parsed['Problem'] = m_prob.group(1).strip()
    
    # 3. Lessons (提取 Markdown 3列表格中的所有行)
    m_less = re.search(r'3\.\s*Lessons[^\n]*\n([\s\S]*?)(?=4\.\s*Potentially|$)', bot_text, re.I)
    if m_less:
        less_text = m_less.group(1).strip()
        lines = [l.strip() for l in less_text.split('\n') if '|' in l and '---' not in l]
        if len(lines) > 1:
            for l in lines[1:]: # 跳过表头行
                cells = [c.strip() for c in l.split('|')[1:-1]]
                if len(cells) >= 3:
                    parsed['Lessons_Rows'].append((cells[0], cells[1], cells[2]))
                elif len(cells) == 2:
                    parsed['Lessons_Rows'].append((cells[0], cells[1], ""))
        if not parsed['Lessons_Rows']:
            parsed['Lessons_Rows'].append((less_text, "", ""))
            
    # 4. Potentially affected (提取表格或文本键值)
    m_pot = re.search(r'4\.\s*Potentially affected[^\n]*\n([\s\S]*?)(?=5\.\s*Appendix|$)', bot_text, re.I)
    if m_pot:
        t = m_pot.group(1)
        w1 = re.search(r'What else[^\n|]*[\|\n]([^\n|]+)', t, re.I)
        w2 = re.search(r'Where can[^\n|]*[\|\n]([^\n|]+)', t, re.I)
        w3 = re.search(r'When can[^\n|]*[\|\n]([^\n|]+)', t, re.I)
        w4 = re.search(r'Who else[^\n|]*[\|\n]([^\n|]+)', t, re.I)
        if w1: parsed['What_Else'] = w1.group(1).strip()
        if w2: parsed['Where'] = w2.group(1).strip()
        if w3: parsed['When'] = w3.group(1).strip()
        if w4: parsed['Who'] = w4.group(1).strip()
        
    return parsed

# -----------------------------------------------------------------------------
# 5. 精准装配 Word 模板核心函数 (100% 对应单元格与段落)
# -----------------------------------------------------------------------------

def set_cell_formatted_text(cell, text):
    """安全清空单元格并设置为标准 Arial 10.5pt 格式"""
    cell.text = ""
    p = cell.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    run = p.add_run(str(text).strip())
    run.font.name = 'Arial'
    run.font.size = Pt(10.5)
    run.font.bold = False

def write_to_doc_section(doc, heading_kw, text_value):
    """向指定章节标题下方写入内容"""
    if not text_value: return
    paragraphs = doc.paragraphs
    for idx, p in enumerate(paragraphs):
        if heading_kw.lower() in p.text.lower() and len(p.text.strip()) < 50:
            target_idx = idx + 1
            while target_idx < len(paragraphs) and not paragraphs[target_idx].text.strip():
                target_idx += 1
            if target_idx < len(paragraphs):
                target_p = paragraphs[target_idx]
                target_p.text = ""
                r = target_p.add_run(str(text_value).strip())
                r.font.name = 'Arial'
                r.font.size = Pt(10.5)
                r.font.bold = False
            break

def populate_docx_exact_tables(template_source, bot_data, raw_row, ok_img=None, ng_img=None):
    """
    【表格定向精准装配引擎】
    - 0. Abstract -> 段落
    - 1. Product/Process -> 段落
    - 2. Problem -> 段落
    - Pictures -> 图片表格 (OK-Part / Not-OK-Part)
    - 3. Lessons -> 3列表格 (Lessons | Measures | Root Cause) 动态增行
    - 4. Potentially affected -> 2列4行表格 (What else / Where / When / Who)
    """
    doc = docx.Document(template_source)
    current_date_str = datetime.date.today().strftime('%b %d %Y')
    failure_mode_str = str(raw_row.get('Failure Mode', '')).strip()

    # 1. 替换页眉页脚老日期与主标题
    for p in doc.paragraphs:
        if 'May 24 2022' in p.text or 'May 24, 2022' in p.text:
            p.text = p.text.replace('May 24 2022', current_date_str).replace('May 24, 2022', current_date_str)
        if 'lesson learn' in p.text.lower() and len(p.text) < 60:
            if failure_mode_str and failure_mode_str not in p.text:
                p_clean = p.text.strip().rstrip("–-—: ").strip()
                p.text = f"{p_clean} – {failure_mode_str}"

    # 2. 填充 0. Abstract
    abs_issue = bot_data.get('Abstract_Issue') or raw_row.get('LL Brief Description', '')
    abs_prob = bot_data.get('Abstract_Problem') or raw_row.get('Failure Mode', '')
    abs_less = bot_data.get('Abstract_Lessons') or raw_row.get('Should or not to do', '')
    abs_full = f"Issue: {abs_issue}\nProblem: {abs_prob}\nLessons: {abs_less}"
    write_to_doc_section(doc, "Abstract", abs_full)

    # 3. 填充 1. Product / Process
    pp_val = bot_data.get('Product_Process') or raw_row.get('Related Material Field / Process', '')
    comp_val = bot_data.get('Component') or raw_row.get('Project/Part name', '')
    sub_val = bot_data.get('Sub_Component') or ''
    p1_full = f"Product / Process: {pp_val}\nComponent: {comp_val}\nSub-Component: {sub_val}"
    write_to_doc_section(doc, "1. Product / Process", p1_full)

    # 4. 填充 2. Problem (Fundamental Problem)
    prob_val = bot_data.get('Problem') or raw_row.get('LL Brief Description', '')
    write_to_doc_section(doc, "2. Problem", prob_val)

    # 5. 定向精准填充各个表格
    for table in doc.tables:
        t_header = "".join(cell.text for cell in table.rows[0].cells).lower()
        
        # A. 锁定 Pictures 表格
        if "ok-part" in t_header or "not-ok-part" in t_header or (len(table.columns) == 2 and len(table.rows) <= 2):
            for row in table.rows:
                for cell in row.cells:
                    if "ok-part" in cell.text.lower() and ok_img:
                        cell.text = "OK-Part\n"
                        p = cell.paragraphs[0]
                        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                        p.add_run().add_picture(io.BytesIO(ok_img), width=Inches(2.4))
                    elif "not-ok-part" in cell.text.lower() and ng_img:
                        cell.text = "Not-OK-Part\n"
                        p = cell.paragraphs[0]
                        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                        p.add_run().add_picture(io.BytesIO(ng_img), width=Inches(2.4))
                        
        # B. 锁定 3. Lessons (3列对策表)
        elif "lessons" in t_header and ("measures" in t_header or "root cause" in t_header):
            lessons_rows = bot_data.get('Lessons_Rows', [])
            if not lessons_rows:
                lessons_rows = [(
                    raw_row.get('Should or not to do', ''),
                    raw_row.get('Corrective Action', ''),
                    raw_row.get('Root Cause', '')
                )]
            # 保留表头，移除旧占位行
            while len(table.rows) > 1:
                tr = table.rows[-1]._tr
                table._tbl.remove(tr)
            # 动态根据条目增行并写入 3 列
            for row_tuple in lessons_rows:
                new_row = table.add_row()
                for c_idx in range(min(3, len(row_tuple))):
                    set_cell_formatted_text(new_row.cells[c_idx], row_tuple[c_idx])

        # C. 锁定 4. Potentially affected (2列 4行表格)
        elif "what else" in t_header or "potentially" in t_header or len(table.rows) == 4:
            w_map = {
                0: bot_data.get('What_Else') or raw_row.get('What else could be additionally affected?') or 'Similar PCB pattern plating processes.',
                1: bot_data.get('Where') or raw_row.get('Where can the problem additionally occur?') or 'Other production lines.',
                2: bot_data.get('When') or raw_row.get('When can the problem additionally appear?') or 'During parameter fluctuation or delayed maintenance.',
                3: bot_data.get('Who') or raw_row.get('Who else can be affected?') or 'PUQ-PQA, PQT, and relevant suppliers.'
            }
            for r_i, row in enumerate(table.rows):
                if len(row.cells) >= 2 and r_i in w_map:
                    set_cell_formatted_text(row.cells[1], w_map[r_i])

    return doc

def generate_eml_file(row_data, to_emails="", doc_bytes=None, doc_filename="LL_Template.docx"):
    """生成带附件、默认可编辑草稿的 Outlook EML 邮件"""
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
    msg.add_header('X-Unsent', '1') # 草稿可编辑
    
    alt_part = MIMEMultipart('alternative')
    alt_part.attach(MIMEText(html_body, 'html', 'utf-8'))
    msg.attach(alt_part)
    
    if doc_bytes:
        part_doc = MIMEBase('application', 'vnd.openxmlformats-officedocument.wordprocessingml.document')
        part_doc.set_payload(doc_bytes)
        encoders.encode_base64(part_doc)
        part_doc.add_header('Content-Disposition', f'attachment; filename="{doc_filename}"')
        msg.attach(part_doc)
        
    return msg.as_bytes()

# -----------------------------------------------------------------------------
# 6. 主交互流程
# -----------------------------------------------------------------------------

if excel_file is not None and template_file is not None:
    try:
        df, sheet_name, header_idx = load_excel_robust(excel_file)
        supplier_dict = load_supplier_emails(excel_file)
        
        # 过滤 LL Need or not == Y
        ll_need_col = next((c for c in df.columns if 'need or not' in str(c).lower()), 'LL Need or not')
        if ll_need_col in df.columns:
            df = df[df[ll_need_col].astype(str).str.strip().str.upper() == 'Y']
            st.success(f"🎉 成功载入 **{sheet_name}**：已过滤保留 `{ll_need_col} = 'Y'` 的 **{len(df)}** 条有效记录。")
            
        serial_no_col = next((c for c in df.columns if 'serial' in str(c).lower()), 'LL Serials No')
        supplier_scope_col = next((c for c in df.columns if 'scope' in str(c).lower() or 'task' in str(c).lower()), 'LL Supplier Scope')
        
        # 步骤 1：选择记录
        st.markdown("### 1️⃣ 步骤一：选择需要处理的台账记录")
        search_kw = st.text_input("🔍 搜索记录 (序列号/供应商/失效模式):")
        filtered_df = df.copy()
        if search_kw:
            filtered_df = df[df.astype(str).apply(lambda r: r.str.contains(search_kw, case=False).any(), axis=1)]
            
        selected_record_idx = st.selectbox(
            "👉 请选择台账记录：",
            options=filtered_df.index,
            format_func=lambda x: f"[{filtered_df.loc[x, serial_no_col]}] {filtered_df.loc[x, 'Failure Mode']} - {filtered_df.loc[x, 'Project/Part name']}"
        )
        
        selected_row = filtered_df.loc[selected_record_idx]
        ok_img, ng_img = get_images_for_row(excel_file, sheet_name, header_idx, selected_record_idx)
        
        # 构建完整嵌入表格的 FEBER Prompt (1:1 还原截图结构)
        prompt_content = f"""Please create me a short and precise lessons learned report out of the attached document in American English.
You are an honest engineer; you provide always links to the sources and name the original slide/page number.
Please stick to the facts. In case you have additional topics, supporting or additional useful information be creative, add them and highlight them in italic.

Please write the headings in bold. Use key words that a
