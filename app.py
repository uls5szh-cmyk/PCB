# -*- coding: utf-8 -*-
"""
=============================================================================
BOSCH | PCB Lesson Learn Quality Studio (Standard FEBER Edition)
- Blue Hints Auto-Wiper (Color & Semantic Level)
- Dynamic Table Row Expansion (3. Lessons & 4. Potentially affected)
- Exact Match with Teams M-PU Bot Engineering Prompt
=============================================================================
"""

import streamlit as st
import pandas as pd
import docx
from docx.shared import Inches, Pt, RGBColor
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
# 1. 页面基本配置与博世工业视觉设计 (Bosch CI Style)
# -----------------------------------------------------------------------------
st.set_page_config(
    page_title="Bosch | PCB Lesson Learn Studio",
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
        <p style="color: #525F6B; font-size: 0.9rem; margin: 0;">FEBER 规范深度对齐 · 蓝色提示词全自动清洗 · 表格自适应增行 · 邮件一键分发</p>
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
    
    if os.path.exists(excel_path):
        excel_file = excel_path
    else:
        st.sidebar.warning("⚠️ 未在指定路径找到 Excel，请手动上传。")
        
    if os.path.exists(template_path):
        template_file = template_path
    else:
        st.sidebar.warning("⚠️ 未在指定路径找到 Word 模板，请手动上传。")
else:
    up_excel = st.sidebar.file_uploader("上传 Master List (Excel):", type=["xlsx", "xlsm"])
    up_template = st.sidebar.file_uploader("上传 Word 模板 (.docx):", type=["docx"])
    if up_excel: excel_file = up_excel
    if up_template: template_file = up_template

# -----------------------------------------------------------------------------
# 3. 辅助解析工具函数
# -----------------------------------------------------------------------------

def load_supplier_emails(file_source):
    """从 Vendor code 表中读取供应商与邮箱"""
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
    """从 Excel 指定行提取 OK 和 NG 图片"""
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
    """加载 Excel 并定位表头行"""
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
# 4. 解析 Bot 润色结果与提示词深度清洗算法
# -----------------------------------------------------------------------------

def parse_bot_feber_response(bot_text):
    """深度解析 Teams M-PU Bot 的标准结构化输出"""
    parsed = {
        'Abstract_Issue': '',
        'Abstract_Problem': '',
        'Abstract_Lessons': '',
        'Product_Process': '',
        'Component': '',
        'Sub_Component': '',
        'Problem': '',
        'Lessons_Table_Rows': [],  # list of tuples: (Lesson, Measure, RootCause)
        'What_Else': '',
        'Where': '',
        'When': '',
        'Who': ''
    }
    
    # 0. Abstract
    abs_m = re.search(r'Abstract\s*([\s\S]*?)(?=1\.\s*Product|$)', bot_text, re.I)
    if abs_m:
        t = abs_m.group(1)
        i_m = re.search(r'Issue:\s*([^\n]+)', t, re.I)
        p_m = re.search(r'Problem:\s*([^\n]+)', t, re.I)
        l_m = re.search(r'Lessons:\s*([^\n]+)', t, re.I)
        if i_m: parsed['Abstract_Issue'] = i_m.group(1).strip()
        if p_m: parsed['Abstract_Problem'] = p_m.group(1).strip()
        if l_m: parsed['Abstract_Lessons'] = l_m.group(1).strip()

    # 1. Product / Process
    p_m = re.search(r'1\.\s*Product\s*/\s*Process\s*([\s\S]*?)(?=2\.\s*Problem|$)', bot_text, re.I)
    if p_m:
        t = p_m.group(1)
        pp = re.search(r'Product\s*/\s*Process:\s*([^\n]+)', t, re.I)
        comp = re.search(r'Component:\s*([^\n]+)', t, re.I)
        sub = re.search(r'Sub-Component:\s*([^\n]+)', t, re.I)
        parsed['Product_Process'] = pp.group(1).strip() if pp else ""
        parsed['Component'] = comp.group(1).strip() if comp else ""
        parsed['Sub_Component'] = sub.group(1).strip() if sub else ""

    # 2. Problem
    prob_m = re.search(r'2\.\s*Problem[^\n]*\n([\s\S]*?)(?=3\.\s*Lessons|$)', bot_text, re.I)
    if prob_m:
        parsed['Problem'] = prob_m.group(1).strip()

    # 3. Lessons (提取表格各行或列表项)
    less_m = re.search(r'3\.\s*Lessons[^\n]*\n([\s\S]*?)(?=4\.\s*Potentially|$)', bot_text, re.I)
    if less_m:
        less_text = less_m.group(1).strip()
        # 尝试按行解析 Markdown 表格
        table_lines = [l.strip() for l in less_text.split('\n') if '|' in l and '---' not in l]
        if len(table_lines) > 1:
            # 存在 Markdown 表格
            for l in table_lines[1:]: # 跳过表头
                parts = [c.strip() for c in l.split('|') if c.strip()]
                if len(parts) >= 3:
                    parsed['Lessons_Table_Rows'].append((parts[0], parts[1], parts[2]))
                elif len(parts) == 2:
                    parsed['Lessons_Table_Rows'].append((parts[0], parts[1], ""))
        if not parsed['Lessons_Table_Rows']:
            # 纯文本模式：提取整段
            parsed['Lessons_Table_Rows'].append((less_text, less_text, ""))

    # 4. Potentially affected
    pot_m = re.search(r'4\.\s*Potentially affected[^\n]*\n([\s\S]*?)(?=5\.\s*Appendix|$)', bot_text, re.I)
    if pot_m:
        t = pot_m.group(1)
        w1 = re.search(r'What else[^\n]*\n([^\n]+)', t, re.I)
        w2 = re.search(r'Where can[^\n]*\n([^\n]+)', t, re.I)
        w3 = re.search(r'When can[^\n]*\n([^\n]+)', t, re.I)
        w4 = re.search(r'Who else[^\n]*\n([^\n]+)', t, re.I)
        if w1: parsed['What_Else'] = w1.group(1).strip()
        if w2: parsed['Where'] = w2.group(1).strip()
        if w3: parsed['When'] = w3.group(1).strip()
        if w4: parsed['Who'] = w4.group(1).strip()

    return parsed

def wipe_all_blue_hints_and_prompts(doc):
    """
    【核心杀手锏：全域蓝色字体与说明性提示词清洗引擎】
    - 颜色级：检测 run 颜色，凡是属于蓝色系的提示词一律置空；
    - 语义级：匹配博世模板固有的 Note、Hint、Guideline 说明段落并清空。
    """
    # 固有的提示词文本特征库（包含中英文说明）
    KNOWN_HINT_PATTERNS = [
        r"^note:",
        r"^general hint:",
        r"^describe briefly",
        r"^briefly describe",
        r"^provide detailed",
        r"^concentrate on the actual",
        r"^keep it short",
        r"^support your content",
        r"^use key words",
        r"^document your lessons",
        r"^what would i suggest",
        r"^what measures did we",
        r"^sustainable solution –",
        r"^do not repeat the 8d",
        r"^if required add",
        r"^describe the main root",
        r"^if necessary for better",
        r"^if not, use the appendix",
        r"^if helpful, attach",
        r"^determine who else",
        r"\(similar applications",
        r"\(other production lines",
        r"\(new applications",
        r"\(other customers",
        r"^check if centers of competence",
        r"^refer to this appendix",
        r"^include reference numbers",
        r"^beware: this document will be uploaded",
        r"delete all hints in blue letters",
        r"^\.\.\.$",
        r"^…$"
    ]
    
    # 扫描正文段落
    for p in doc.paragraphs:
        p_text_strip = p.text.strip().lower()
        
        # 1. 语义特征清除
        if any(re.search(pat, p_text_strip) for pat in KNOWN_HINT_PATTERNS):
            p.text = ""
            continue
            
        # 2. 颜色特征清除 (扫描 Runs)
        for run in p.runs:
            is_blue = False
            # 检查 RGBColor
            if run.font.color and run.font.color.rgb:
                r, g, b = run.font.color.rgb
                if b > 150 and r < 100: # 判定为蓝色系
                    is_blue = True
            # 检查底层 XML 颜色
            rPr = run._r.get_or_add_rPr()
            color_elem = rPr.find(docx.oxml.ns.qn('w:color'))
            if color_elem is not None:
                val = color_elem.get(docx.oxml.ns.qn('w:val'), '').lower()
                if val in ['0070c0', '0000ff', '007bc0', '002060', '418ab3', '1f497d', '5b9bd5']:
                    is_blue = True
            if is_blue:
                run.text = ""
                
    # 扫描表格中的单元格段落
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    p_text_strip = p.text.strip().lower()
                    if any(re.search(pat, p_text_strip) for pat in KNOWN_HINT_PATTERNS):
                        p.text = ""
                        continue
                    for run in p.runs:
                        is_blue = False
                        if run.font.color and run.font.color.rgb:
                            r, g, b = run.font.color.rgb
                            if b > 150 and r < 100:
                                is_blue = True
                        rPr = run._r.get_or_add_rPr()
                        color_elem = rPr.find(docx.oxml.ns.qn('w:color'))
                        if color_elem is not None:
                            val = color_elem.get(docx.oxml.ns.qn('w:val'), '').lower()
                            if val in ['0070c0', '0000ff', '007bc0', '002060', '418ab3', '1f497d', '5b9bd5']:
                                is_blue = True
                        if is_blue:
                            run.text = ""

def write_clean_content_to_section(doc, heading_text, content_text):
    """定位指定标题，清空其下方的所有旧内容与提示词，并写入新文本"""
    if not content_text:
        return
    paragraphs = doc.paragraphs
    for idx, p in enumerate(paragraphs):
        if heading_text.lower() in p.text.lower() and len(p.text.strip()) < 45:
            # 找到下一个非空段落或标题
            target_idx = idx + 1
            while target_idx < len(paragraphs) and not paragraphs[target_idx].text.strip():
                target_idx += 1
            if target_idx < len(paragraphs):
                target_p = paragraphs[target_idx]
                target_p.text = ""
                run = target_p.add_run(str(content_text).strip())
                run.font.name = 'Arial'
                run.font.size = Pt(10.5)
                run.font.bold = False
            break

def populate_docx_master(template_source, bot_data, raw_row, ok_img=None, ng_img=None):
    """
    【最终交付模板装配引擎】
    - 清洗全部蓝色提示词；
    - 填充 0~5 章节；
    - `3. Lessons` 表格自适应行数动态扩展；
    - 图片表格自适应置入。
    """
    doc = docx.Document(template_source)
    current_date_str = datetime.date.today().strftime('%b %d %Y')
    failure_mode_str = str(raw_row.get('Failure Mode', '')).strip()

    # 1. 替换页眉页脚与老旧日期、追加主标题
    for p in doc.paragraphs:
        if 'May 24 2022' in p.text or 'May 24, 2022' in p.text:
            p.text = p.text.replace('May 24 2022', current_date_str).replace('May 24, 2022', current_date_str)
        if 'lesson learn' in p.text.lower() and len(p.text) < 60:
            if failure_mode_str and failure_mode_str not in p.text:
                p_clean = p.text.strip().rstrip("–-—: ").strip()
                p.text = f"{p_clean} – {failure_mode_str}"

    # 2. 先执行全域蓝色提示词物理清除
    wipe_all_blue_hints_and_prompts(doc)

    # 3. 填充 0. Abstract
    abs_full = ""
    if bot_data.get('Abstract_Issue'): abs_full += f"Issue: {bot_data['Abstract_Issue']}\n"
    if bot_data.get('Abstract_Problem'): abs_full += f"Problem: {bot_data['Abstract_Problem']}\n"
    if bot_data.get('Abstract_Lessons'): abs_full += f"Lessons: {bot_data['Abstract_Lessons']}\n"
    write_clean_content_to_section(doc, "Abstract", abs_full)

    # 4. 填充 1. Product / Process
    pp_content = ""
    if bot_data.get('Product_Process'): pp_content += f"Product / Process: {bot_data['Product_Process']}\n"
    elif raw_row.get('Related Material Field / Process'): pp_content += f"Product / Process: {raw_row.get('Related Material Field / Process')}\n"
    if bot_data.get('Component'): pp_content += f"Component: {bot_data['Component']}\n"
    elif raw_row.get('Project/Part name'): pp_content += f"Component: {raw_row.get('Project/Part name')}\n"
    if bot_data.get('Sub_Component'): pp_content += f"Sub-Component: {bot_data['Sub_Component']}\n"
    write_clean_content_to_section(doc, "1. Product / Process", pp_content)

    # 5. 填充 2. Problem (Fundamental Problem)
    prob_text = bot_data.get('Problem') or raw_row.get('LL Brief Description', '')
    write_clean_content_to_section(doc, "2. Problem", prob_text)

    # 6. 处理 3. Lessons 表格（自适应增行）
    lessons_rows = bot_data.get('Lessons_Table_Rows', [])
    if not lessons_rows:
        # 使用原表数据兜底
        lessons_rows = [(
            raw_row.get('Should or not to do', ''),
            raw_row.get('Corrective Action', ''),
            raw_row.get('Root Cause', '')
        )]
        
    for table in doc.tables:
        header_text = "".join(cell.text for cell in table.rows[0].cells).lower()
        if "lessons" in header_text and "measures" in header_text:
            # 找到 Lessons 表格，保留表头（Row 0），清空或替换后续行
            while len(table.rows) > 1:
                # 移除旧占位行
                tr = table.rows[-1]._tr
                table._tbl.remove(tr)
                
            # 动态根据记录行数增加行
            for row_tuple in lessons_rows:
                new_row = table.add_row()
                for c_idx in range(min(3, len(row_tuple))):
                    cell = new_row.cells[c_idx]
                    cell.text = ""
                    p = cell.paragraphs[0]
                    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
                    run = p.add_run(str(row_tuple[c_idx]).strip())
                    run.font.name = 'Arial'
                    run.font.size = Pt(10.5)
                    run.font.bold = False
            break

    # 7. 填充 4. Potentially affected
    pot_full = ""
    w1 = bot_data.get('What_Else') or raw_row.get('What else could be additionally affected?') or 'Similar PCB pattern plating and surface finish processes.'
    w2 = bot_data.get('Where') or raw_row.get('Where can the problem additionally occur?') or 'Other production lines with comparable specifications.'
    w3 = bot_data.get('When') or raw_row.get('When can the problem additionally appear?') or 'During parameter fluctuation or delayed equipment maintenance.'
    w4 = bot_data.get('Who') or raw_row.get('Who else can be affected?') or 'PQT, PQA, and relevant Tier-1 suppliers.'
    
    pot_full = f"""What else could be additionally affected?
{w1}

Where can the problem additionally occur?
{w2}

When can the problem additionally appear?
{w3}

Who else can be affected?
{w4}"""
    write_clean_content_to_section(doc, "4. Potentially affected", pot_full)

    # 8. 图片表格自适应置入
    for table in doc.tables:
        text = "".join(cell.text for row in table.rows for cell in row.cells)
        if "OK-Part" in text and "Not-OK-Part" in text:
            for row in table.rows:
                for cell in row.cells:
                    if "OK-Part" in cell.text and ok_img:
                        cell.text = "OK-Part\n"
                        p = cell.paragraphs[0]
                        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                        p.add_run().add_picture(io.BytesIO(ok_img), width=Inches(2.4))
                    elif "Not-OK-Part" in cell.text and ng_img:
                        cell.text = "Not-OK-Part\n"
                        p = cell.paragraphs[0]
                        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                        p.add_run().add_picture(io.BytesIO(ng_img), width=Inches(2.4))
            break

    return doc

def generate_eml_file(row_data, to_emails="", doc_bytes=None, doc_filename="LL_Template.docx"):
    """生成带附件、处于草稿编辑模式的 Outlook EML 邮件"""
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
    msg.add_header('X-Unsent', '1')
    
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
# 5. 主交互界面：闭环实战操作流
# -----------------------------------------------------------------------------

if excel_file is not None and template_file is not None:
    try:
        df, sheet_name, header_idx = load_excel_robust(excel_file)
        supplier_dict = load_supplier_emails(excel_file)
        
        # 过滤 LL Need or not == Y
        ll_need_col = next((c for c in df.columns if 'need or not' in str(c).lower()), 'LL Need or not')
        if ll_need_col in df.columns:
            orig_len = len(df)
            df = df[df[ll_need_col].astype(str).str.strip().str.upper() == 'Y']
            st.success(f"🎉 成功载入 **{sheet_name}**：已过滤保留 `{ll_need_col} = 'Y'` 的 **{len(df)}** 条有效记录。")
            
        serial_no_col = next((c for c in df.columns if 'serial' in str(c).lower()), 'LL Serials No')
        supplier_scope_col = next((c for c in df.columns if 'scope' in str(c).lower() or 'task' in str(c).lower()), 'LL Supplier Scope')
        
        # 步骤 1：选择记录
        st.markdown("### 1️⃣ 步骤一：在台账中选择一条记录")
        search_kw = st.text_input("🔍 搜索记录 (序列号/供应商/失效模式):")
        filtered_df = df.copy()
        if search_kw:
            filtered_df = df[df.astype(str).apply(lambda r: r.str.contains(search_kw, case=False).any(), axis=1)]
            
        selected_record_idx = st.selectbox(
            "👉 请选择要生成报告的台账记录：",
            options=filtered_df.index,
            format_func=lambda x: f"[{filtered_df.loc[x, serial_no_col]}] {filtered_df.loc[x, 'Failure Mode']} - {filtered_df.loc[x, 'Project/Part name']}"
        )
        
        selected_row = filtered_df.loc[selected_record_idx]
        ok_img, ng_img = get_images_for_row(excel_file, sheet_name, header_idx, selected_record_idx)
        
        # 构建与您标准图完全一致的专业 FEBER Prompt
        prompt_content = f"""Please create me a short and precise lessons learned report out of the attached document in American English.
You are an honest engineer; you provide always links to the sources and name the original slide/page number.
Please stick to the facts. In case you have additional topics, supporting or additional useful information be creative, add them and highlight them in italic.

Please write the headings in bold. Use key words that are understood by others in Bosch. Describe the report "user-friendly", so others can read it easily. One page for chapter 1-4 is appropriate. Delete all hints in blue letters.

If you are asked to create a lesson learned report, or to search for a lessons learned report, structure the answer as follows:
0. Abstract - write a short summary of the report with the structure - issue; problem; learnings; tags

Abstract
Issue: {selected_row.get('LL Brief Description', '')}
Problem: {selected_row.get('Failure Mode', '')}
Lessons: {selected_row.get('Should or not to do', '')}

1. Product / Process
Product / Process: {selected_row.get('Related Material Field / Process', '')}
Component: {selected_row.get('Project/Part name', '')}
Sub-Component: 

2. Problem (Fundamental Problem)
{selected_row.get('LL Brief Description', '')}

3. Lessons
| Lessons | Measures & Sustainable Solutions | Root Cause |
| --- | --- | --- |
| {selected_row.get('Should or not to do', '')} | {selected_row.get('Corrective Action', '')} | {selected_row.get('Root Cause', '')} |

4. Potentially affected
What else could be additionally affected?
{selected_row.get('What else could be additionally affected?', 'Similar PCB pattern plating processes.')}

Where can the problem additionally occur?
{selected_row.get('Where can the problem additionally occur?', 'Other production lines.')}

When can the problem additionally appear?
{selected_row.get('When can the problem additionally appear?', 'During parameter fluctuations.')}

Who else can be affected?
{selected_row.get('Who else can be affected?', 'PUQ-PQA, PQT, and relevant suppliers.')}

5. Appendix (Optional)"""

        # 步骤 2：复制 Prompt 并在 Teams M-PU Bot 中润色
        st.markdown("---")
        st.markdown("### 2️⃣ 步骤二：复制专属 Prompt 并发给 Teams M-PU ChatGPT Bot")
        
        c_p, c_b = st.columns([3, 1])
        with c_p:
            st.text_area("📋 已按博世标准结构组织的完整 Prompt:", prompt_content, height=240)
        with c_b:
            st.markdown("<br>", unsafe_allow_html=True)
            st.link_button("🚀 一键直达 Teams M-PU Bot", TEAMS_BOT_URL, use_container_width=True)
            st.caption("💡 复制左侧 Prompt 并在 Teams Bot 窗口粘贴发送。")

        # 步骤 3：粘贴 Bot 输出并一键生成
        st.markdown("---")
        st.markdown("### 3️⃣ 步骤三：粘贴 M-PU Bot 输出文本并生成交付件")
        
        col_in, col_sup = st.columns([3, 2])
        with col_in:
            bot_reply = st.text_area(
                "📥 在此粘贴 M-PU Bot 润色后的完整回复：",
                height=220,
                placeholder="粘贴 Bot 输出的包含 0. Abstract, 1. Product/Process, 2. Problem, 3. Lessons, 4. Potentially affected 的完整文本..."
            )
        with col_sup:
            selected_sups = st.multiselect("👥 选择收件供应商 (自动读取 Vendor code 邮箱):", options=list(supplier_dict.keys()))
            to_emails_list = []
            for s in selected_sups:
                to_emails_list.extend(supplier_dict[s])
            to_emails_str = "; ".join(list(set(to_emails_list)))
            if to_emails_str:
                st.info(f"📧 **自动收件人:**\n`{to_emails_str}`")

        if st.button("🚀 立即安全生成 Word 报告与 Outlook 邮件草稿", type="primary", use_container_width=True):
            if template_file is None:
                st.error("❌ 未检测到 Word 模板，请在侧边栏确认路径。")
            else:
                with st.spinner("正在彻底抹除蓝色提示词、动态构建表格并嵌入图片..."):
                    bot_data = parse_bot_feber_response(bot_reply) if bot_reply.strip() else {}
                    
                    # 生成装配好的 Word
                    doc = populate_docx_master(template_file, bot_data, selected_row, ok_img, ng_img)
                    bio = io.BytesIO()
                    doc.save(bio)
                    doc_bytes = bio.getvalue()
                    
                    serial_str = str(selected_row.get(serial_no_col, 'LL-Export'))
                    doc_filename = f"LL_Template_{serial_str}.docx"
                    eml_bytes = generate_eml_file(selected_row, to_emails_str, doc_bytes, doc_filename)
                    
                    st.success("🎉 生成成功！所有蓝色提示词已彻底清除，3. Lessons 表格已按需动态增行，图片已嵌入完毕。")
                    
                    c_d1, c_d2 = st.columns(2)
                    with c_d1:
                        st.download_button(
                            f"📥 下载 Word 报告: {doc_filename}",
                            doc_bytes,
                            doc_filename,
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                            use_container_width=True
                        )
                    with c_d2:
                        st.download_button(
                            f"📧 下载 Outlook 草稿 (已带Word附件): Email_Draft_{serial_str}.eml",
                            eml_bytes,
                            f"Email_Draft_{serial_str}.eml",
                            mime="message/rfc822",
                            use_container_width=True
                        )

    except Exception as e:
        st.error(f"❌ 运行异常: {e}")
else:
    st.info("ℹ️ 请在侧边栏确认 Master List (Excel) 和 Word 模板的文件路径。")
