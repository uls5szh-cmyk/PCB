# -*- coding: utf-8 -*-
"""
=============================================================================
BOSCH | PCB Lesson Learn Quality Studio (Exact Table & Picture Match Edition)
- Automatic Extraction from Excel 'Picture' / 'NG Picture' Column
- Direct Injection into Word Table Cell containing 'Picture' (Case-Insensitive)
- 1:1 Mirror Prompt with 3-Column Tables & Dynamic Row Expansion
- Recipient-Ready Outlook EML Draft Generation with Word Attachment
=============================================================================
"""

import streamlit as st
import pandas as pd
import docx
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT
from docx.oxml import OxmlElement
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
    .bds-card {
        background: #FFFFFF;
        border: 1px solid #DDE3EA;
        border-radius: 8px;
        padding: 20px;
        margin-bottom: 20px;
        box-shadow: 0 4px 12px rgba(0, 40, 80, 0.04);
    }
    .bds-step-badge {
        display: inline-block;
        background: var(--bosch-blue);
        color: #FFFFFF;
        font-size: 0.8rem;
        font-weight: 700;
        padding: 4px 12px;
        border-radius: 20px;
        margin-bottom: 10px;
    }
    .stButton>button {
        background-color: var(--bosch-blue) !important;
        color: white !important;
        border-radius: 4px !important;
        font-weight: 600 !important;
        border: none !important;
        padding: 8px 20px !important;
    }
    .stButton>button:hover {
        background-color: var(--bosch-light-blue) !important;
        color: white !important;
    }
</style>
<div class="bosch-top-bar"></div>
"""
st.markdown(BOSCH_UI_STYLE, unsafe_allow_html=True)

st.markdown("""
<div style="display: flex; align-items: center; justify-content: space-between; margin-bottom: 15px;">
    <div>
        <h2 style="color: #005691; margin: 0; font-weight: 700;">🔴 BOSCH | PCB Lesson Learn 协同工作台</h2>
        <p style="color: #525F6B; font-size: 0.95rem; margin: 4px 0 0 0;">FEBER 质量报告规范 · Picture 单元格图片自动注入 · 3列表格动态增行 · 邮件草稿一键闭环</p>
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
# 3. 辅助解析函数（高可靠性 Picture 提取）
# -----------------------------------------------------------------------------

def load_supplier_emails(file_source):
    """从 Vendor code 表中读取供应商与邮箱映射"""
    try:
        if hasattr(file_source, 'seek'): file_source.seek(0)
        xl = pd.ExcelFile(file_source)
        if 'Vendor code' not in xl.sheet_names:
            return {}
        if hasattr(file_source, 'seek'): file_source.seek(0)
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

def get_picture_for_row_robust(file_source, sheet_name, header_idx, target_df_loc):
    """
    【高可靠性图片提取】
    - 自动检索包含 'picture' 的列；
    - 结合行号容差匹配最靠近该数据行的图片。
    """
    try:
        if hasattr(file_source, 'seek'): file_source.seek(0)
        wb = openpyxl.load_workbook(file_source, data_only=True)
        ws = wb[sheet_name] if sheet_name in wb.sheetnames else wb.active
        
        # 寻找 Picture 列
        col_pic = -1
        for col_idx in range(1, ws.max_column + 1):
            for check_r in range(1, header_idx + 3):
                val = ws.cell(row=check_r, column=col_idx).value
                if val and 'picture' in str(val).strip().lower():
                    col_pic = col_idx - 1
                    break
            if col_pic != -1:
                break
                
        # 目标 Excel 绝对行号 (0-based)
        excel_target_row = (header_idx + 1) + target_df_loc
        
        candidates = []
        for img in getattr(ws, '_images', []):
            r = None
            c = None
            anchor = getattr(img, 'anchor', None)
            if hasattr(anchor, '_from'):
                r = anchor._from.row
                c = anchor._from.col
            elif isinstance(anchor, str):
                from openpyxl.utils.cell import coordinate_to_tuple
                r_tuple = coordinate_to_tuple(anchor)
                r, c = r_tuple[0] - 1, r_tuple[1] - 1
                
            if r is not None:
                diff = abs(r - excel_target_row)
                candidates.append((diff, img))
                if diff <= 2 and (col_pic == -1 or (c is not None and abs(c - col_pic) <= 3)):
                    return img._data()
                    
        # 兜底：若有候选图片，返回行距最近的
        if candidates:
            candidates.sort(key=lambda x: x[0])
            if candidates[0][0] <= 5:
                return candidates[0][1]._data()
    except Exception as e:
        print(f"图片提取异常: {e}")
    return None

def load_excel_robust(file_source):
    """加载 Excel 并定位表头"""
    if hasattr(file_source, 'seek'): file_source.seek(0)
    xl = pd.ExcelFile(file_source)
    sheet_names = xl.sheet_names
    
    target_sheet = next((s for s in sheet_names if "PUQ3 LL Overall list" in s), None)
    if not target_sheet:
        target_sheet = next((s for s in sheet_names if "Overall" in s or "LL" in s), sheet_names[0])
        
    if hasattr(file_source, 'seek'): file_source.seek(0)
    df_temp = pd.read_excel(file_source, sheet_name=target_sheet, nrows=10, header=None)
    header_idx = 1 
    for idx, row in df_temp.iterrows():
        row_str = [str(val).strip().lower() for val in row.tolist()]
        if any('serials' in val or 'project/part' in val or 'failure mode' in val for val in row_str):
            header_idx = idx
            break
            
    if hasattr(file_source, 'seek'): file_source.seek(0)
    df = pd.read_excel(file_source, sheet_name=target_sheet, header=header_idx)
    df.columns = df.columns.astype(str).str.strip()
    return df, target_sheet, header_idx

# -----------------------------------------------------------------------------
# 4. 高鲁棒性 Bot 输出解析器 (支持制表符 \t 与 Markdown 表格双模式)
# -----------------------------------------------------------------------------

def parse_bot_feber_response(bot_text):
    """解析 Bot 按照 FEBER 规范输出的结构化文本与 3 列表格"""
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
        i_m = re.search(r'Issue:\s*([\s\S]*?)(?=Problem:|$)', t, re.I)
        p_m = re.search(r'Problem:\s*([\s\S]*?)(?=Lessons:|$)', t, re.I)
        l_m = re.search(r'Lessons:\s*([\s\S]*?)(?=Tags:|Picture|1\.\s*Product|$)', t, re.I)
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
    
    # 3. Lessons (支持制表符 \t 与 Markdown | 双模式)
    m_less = re.search(r'3\.\s*Lessons[^\n]*\n([\s\S]*?)(?=4\.\s*Potentially|$)', bot_text, re.I)
    if m_less:
        less_text = m_less.group(1).strip()
        raw_lines = [l.strip() for l in less_text.split('\n') if l.strip()]
        for line in raw_lines:
            if "measures & sustainable solutions" in line.lower() or "---" in line or line.startswith("| :---") or line.startswith("Lessons\t"):
                continue
            if '\t' in line:
                parts = [p.strip() for p in line.split('\t')]
                if len(parts) >= 3: parsed['Lessons_Rows'].append((parts[0], parts[1], parts[2]))
                elif len(parts) == 2: parsed['Lessons_Rows'].append((parts[0], parts[1], ""))
                elif len(parts) == 1: parsed['Lessons_Rows'].append((parts[0], "", ""))
            elif '|' in line:
                parts = [p.strip() for p in line.split('|')[1:-1]]
                if len(parts) >= 3: parsed['Lessons_Rows'].append((parts[0], parts[1], parts[2]))
                elif len(parts) == 2: parsed['Lessons_Rows'].append((parts[0], parts[1], ""))
        if not parsed['Lessons_Rows']:
            parsed['Lessons_Rows'].append((less_text, "", ""))
            
    # 4. Potentially affected
    m_pot = re.search(r'4\.\s*Potentially affected[^\n]*\n([\s\S]*?)(?=5\.\s*Appendix|$)', bot_text, re.I)
    if m_pot:
        t = m_pot.group(1)
        w1 = re.search(r'What else[^\n\t|]*[\t\|\n]([^\n|]+)', t, re.I)
        w2 = re.search(r'Where can[^\n\t|]*[\t\|\n]([^\n|]+)', t, re.I)
        w3 = re.search(r'When can[^\n\t|]*[\t\|\n]([^\n|]+)', t, re.I)
        w4 = re.search(r'Who else[^\n\t|]*[\t\|\n]([^\n|]+)', t, re.I)
        if w1: parsed['What_Else'] = w1.group(1).strip()
        if w2: parsed['Where'] = w2.group(1).strip()
        if w3: parsed['When'] = w3.group(1).strip()
        if w4: parsed['Who'] = w4.group(1).strip()
        
    return parsed

# -----------------------------------------------------------------------------
# 5. 精准装配 Word 模板核心函数 (直接命中 Picture 单元格)
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

def set_aligned_field_paragraph(p, label, value, indent_inches=1.0):
    """为 Abstract 区域设置悬挂缩进（Hanging Indent），使冒号后的文字严格垂直对齐"""
    p.text = ""
    p.paragraph_format.left_indent = Inches(indent_inches)
    p.paragraph_format.first_line_indent = Inches(-indent_inches)
    p.paragraph_format.tab_stops.add_tab_stop(Inches(indent_inches), WD_TAB_ALIGNMENT.LEFT)
    
    r_label = p.add_run(label)
    r_label.font.name = 'Arial'
    r_label.font.size = Pt(10.5)
    r_label.font.bold = True
    
    p.add_run("\t")
    
    r_val = p.add_run(str(value).strip())
    r_val.font.name = 'Arial'
    r_val.font.size = Pt(10.5)
    r_val.font.bold = False

def replace_field_value_in_doc(doc, field_label, new_value, is_abstract=False):
    """精准查找字段标签并更新值"""
    if not new_value: return
    for p in doc.paragraphs:
        if field_label.lower() in p.text.lower():
            if is_abstract:
                set_aligned_field_paragraph(p, field_label, new_value, indent_inches=0.9)
            else:
                p.text = ""
                r_label = p.add_run(f"{field_label} ")
                r_label.font.name = 'Arial'
                r_label.font.size = Pt(10.5)
                r_label.font.bold = True
                
                r_val = p.add_run(str(new_value).strip())
                r_val.font.name = 'Arial'
                r_val.font.size = Pt(10.5)
                r_val.font.bold = False
            return

def insert_content_under_heading(doc, heading_kw, text_value):
    """在指定标题下方安全插入内容段落（确保 2 Problem 100% 写入不遗漏）"""
    if not text_value: return
    for idx, p in enumerate(doc.paragraphs):
        p_txt = p.text.strip().lower()
        if heading_kw.lower() in p_txt and len(p_txt) < 50:
            new_p_elem = OxmlElement('w:p')
            p._element.addnext(new_p_elem)
            new_p = docx.text.paragraph.Paragraph(new_p_elem, doc)
            new_p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            r = new_p.add_run(str(text_value).strip())
            r.font.name = 'Arial'
            r.font.size = Pt(10.5)
            r.font.bold = False
            return

def populate_docx_exact_tables(template_source, bot_data, raw_row, picture_img=None):
    """
    【表格与图文定向精准装配引擎】
    - 0. Abstract -> 对齐分离回填 Issue / Problem / Lessons
    - Picture 单元格 -> 识别包含 'picture' 的单元格并居中置入图片
    - 1. Product/Process -> 分离回填 Product/Process / Component / Sub-Component
    - 2. Problem -> 精准注入正文段落
    - 3. Lessons -> 3列表格 (Lessons | Measures & Sustainable Solutions | Root Cause) 逐行动态增行
    - 4. Potentially affected -> 2列4行表格 (What else / Where / When / Who)
    """
    if hasattr(template_source, 'seek'):
        template_source.seek(0)
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

    # 2. 精准对齐填充 0. Abstract (整齐悬挂缩进)
    abs_issue = bot_data.get('Abstract_Issue') or raw_row.get('LL Brief Description', '')
    abs_prob = bot_data.get('Abstract_Problem') or raw_row.get('Failure Mode', '')
    abs_less = bot_data.get('Abstract_Lessons') or raw_row.get('Should or not to do', '')
    
    replace_field_value_in_doc(doc, "Issue:", abs_issue, is_abstract=True)
    replace_field_value_in_doc(doc, "Problem:", abs_prob, is_abstract=True)
    replace_field_value_in_doc(doc, "Lessons:", abs_less, is_abstract=True)

    # 3. 精准独立填充 1. Product / Process
    pp_val = bot_data.get('Product_Process') or raw_row.get('Related Material Field / Process', '')
    comp_val = bot_data.get('Component') or raw_row.get('Project/Part name', '')
    sub_val = bot_data.get('Sub_Component') or 'Not specified'
    
    replace_field_value_in_doc(doc, "Product / Process:", pp_val)
    replace_field_value_in_doc(doc, "Component:", comp_val)
    replace_field_value_in_doc(doc, "Sub-Component:", sub_val)

    # 4. 填充 2. Problem (Fundamental Problem) -> 采用 addnext 确保 100% 写入
    prob_val = bot_data.get('Problem') or raw_row.get('LL Brief Description', '')
    insert_content_under_heading(doc, "Problem (Fundamental Problem)", prob_val)

    # 5. 定向精准填充各个表格 (Picture 单元格 / 3. Lessons / 4. Potentially affected)
    picture_inserted = False
    for table in doc.tables:
        t_header = "".join(cell.text for cell in table.rows[0].cells).lower()
        
        # A. 扫描任意表格中包含 'picture' 的单元格并置入图片
        for row in table.rows:
            for cell in row.cells:
                c_txt = cell.text.lower().replace(" ", "")
                if "picture" in c_txt:
                    if picture_img:
                        cell.text = ""
                        p = cell.paragraphs[0]
                        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                        p.add_run().add_picture(io.BytesIO(picture_img), width=Inches(3.0))
                        picture_inserted = True

        # B. 锁定 3. Lessons (3列对策表：Lessons | Measures & Sustainable Solutions | Root Cause)
        if "lessons" in t_header and ("measures" in t_header or "root cause" in t_header):
            lessons_rows = bot_data.get('Lessons_Rows', [])
            if not lessons_rows:
                lessons_rows = [(
                    raw_row.get('Should or not to do', ''),
                    raw_row.get('Corrective Action', ''),
                    raw_row.get('Root Cause', '')
                )]
            # 清除旧占位行（只保留表头 Row 0）
            while len(table.rows) > 1:
                tr = table.rows[-1]._tr
                table._tbl.remove(tr)
            # 动态根据条目多行写入 3 个独立单元格
            for row_tuple in lessons_rows:
                new_row = table.add_row()
                for c_idx in range(min(3, len(row_tuple))):
                    set_cell_formatted_text(new_row.cells[c_idx], row_tuple[c_idx])

        # C. 锁定 4. Potentially affected (2列 4行评估表格)
        elif "what else" in t_header or "potentially" in t_header or len(table.rows) == 4:
            w_map = {
                0: bot_data.get('What_Else') or raw_row.get('What else could be additionally affected?') or 'Other PCB suppliers manufacturing multilayer boards using similar pattern plating processes.',
                1: bot_data.get('Where') or raw_row.get('Where can the problem additionally occur?') or 'Other production plants or lines employing similar pattern plating technologies worldwide.',
                2: bot_data.get('When') or raw_row.get('When can the problem additionally appear?') or 'Future PCB product lines or applications with comparable design/process constraints.',
                3: bot_data.get('Who') or raw_row.get('Who else can be affected?') or 'Downstream customers, assembly plants, and Bosch divisions reliant on affected PCBs.'
            }
            for r_i, row in enumerate(table.rows):
                if len(row.cells) >= 2 and r_i in w_map:
                    set_cell_formatted_text(row.cells[1], w_map[r_i])

    # 兜底：如果表格里没找到，检查正文段落
    if not picture_inserted and picture_img:
        for p in doc.paragraphs:
            p_txt_clean = p.text.lower().replace(" ", "")
            if "picture" in p_txt_clean and len(p_txt_clean) < 40:
                p.text = ""
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p.add_run().add_picture(io.BytesIO(picture_img), width=Inches(3.0))
                break

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
# 6. 主交互流程 (BDS 2.0 工业级卡片流)
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
        
        # ==================== STEP 1 ====================
        st.markdown('<div class="bds-card">', unsafe_allow_html=True)
        st.markdown('<span class="bds-step-badge">STEP 1</span> <h4 style="display:inline; margin-left:8px; color:#005691;">选择台账记录并检查图片</h4>', unsafe_allow_html=True)
        
        search_kw = st.text_input("🔍 搜索记录 (序列号/供应商/失效模式):", placeholder="输入关键字实时过滤...")
        filtered_df = df.copy()
        if search_kw:
            filtered_df = df[df.astype(str).apply(lambda r: r.str.contains(search_kw, case=False).any(), axis=1)]
            
        selected_record_idx = st.selectbox(
            "👉 请选择台账记录：",
            options=filtered_df.index,
            format_func=lambda x: f"[{filtered_df.loc[x, serial_no_col]}] {filtered_df.loc[x, 'Failure Mode']} - {filtered_df.loc[x, 'Project/Part name']}"
        )
        
        selected_row = filtered_df.loc[selected_record_idx]
        target_df_loc = filtered_df.index.get_loc(selected_record_idx)
        
        # 从 Excel 提取 Picture 列图片
        picture_img = get_picture_for_row_robust(excel_file, sheet_name, header_idx, target_df_loc)
        
        if picture_img:
            st.info("✅ 已成功从 Excel 的 Picture 列提取到不良图片！")
            st.image(picture_img, caption="提取到的 Picture 预览", width=220)
        else:
            st.warning("⚠️ 未在当前行的 Picture 列中检测到图片，生成时图片位置将保持空白。")
            
        st.markdown('</div>', unsafe_allow_html=True)
        
        # 动态生成纯净的事实清单
        raw_facts_list = []
        for col_name in df.columns:
            if col_name != '选择 (Select)':
                val_str = str(selected_row.get(col_name, '')).strip()
                if val_str and val_str not in ['nan', 'None']:
                    raw_facts_list.append(f"{col_name}: {val_str}")
        raw_facts_block = "\n".join(raw_facts_list)
        
        # 1:1 还原包含 3 列表格原型的标准 FEBER Prompt
        prompt_content = f"""Please create me a short and precise lessons learned report out of the attached document in American English.
You are an honest engineer; you provide always links to the sources and name the original slide/page number.
Please stick to the facts. In case you have additional topics, supporting or additional useful information be creative, add them and highlight them in italic.

Please write the headings in bold. Use key words that are understood by others in Bosch. Describe the report "user-friendly", so others can read it easily. One page for chapter 1-4 is appropriate. Delete all hints in blue letters.

If you are asked to create a lesson learned report, or to search for a lessons learned report, structure the answer as follows:
0. Abstract - write a short summary of the report with the structure - issue; problem; learnings; tags

Abstract
Issue: Describe briefly. Do not use abbreviations that are not commonly known.
Problem: Briefly describe the main problem using key words.
Lessons: Concentrate on the actual lessons learned. What is new? How can we prevent this issue in the future?

Picture – Product – Defect
General Hint:
• Keep it short. Two pages for chapter 1-4 should be sufficient.
• Support your content with pictures where appropriate.
• Use key words that are understood by others in Bosch and not only in your area of expertise.

1. Product / Process
Product / Process:
Component:
Sub-Component:

2. Problem (Fundamental Problem)
Note:
Briefly describe the fundamental problem. Do not use technical root causes (TRC) or managerial root causes (MRC). Use pictures or graphs to visualize the problem.

3. Lessons
Document your lessons in the table.
| Lessons | Measures & Sustainable Solutions | Root Cause |
| :--- | :--- | :--- |
| What do you suggest could be done differently next time. | What measures did we implement in our area that might also be useful to others? | Describe the main root cause only. |

4. Potentially affected.
Determine who else might find this information useful to the best of your knowledge:
| Aspect | Description |
| :--- | :--- |
| What else could be additionally be affected? | (Similar applications, products, processes, …) or Not applicable |
| Where can the problem additionally occur? | (Other production lines, plants, regions, …) or Not applicable |
| When can the problem additionally appear? | (New applications, usage of products, …) or Not applicable |
| Who else can be affected? | (Other customers, suppliers, associates, …) or Not applicable |

Check if Centers of Competence (CoC) or BEO working groups should be informed: https://connect.bosch.com/communities/community/BEO

5. Appendix (Optional)

==================== [Raw Master List Facts] ====================
{raw_facts_block}"""

        # ==================== STEP 2 ====================
        st.markdown('<div class="bds-card">', unsafe_allow_html=True)
        st.markdown('<span class="bds-step-badge">STEP 2</span> <h4 style="display:inline; margin-left:8px; color:#005691;">一键复制 Prompt 并在 Teams M-PU Bot 润色</h4>', unsafe_allow_html=True)
        
        c_p, c_b = st.columns([3, 1])
        with c_p:
            st.text_area("📋 已完整内嵌 3 列表格原型的工程 Prompt (点击右上角图标复制):", prompt_content, height=220)
        with c_b:
            st.markdown("<br>", unsafe_allow_html=True)
            st.link_button("🚀 一键直达 Teams M-PU Bot", TEAMS_BOT_URL, use_container_width=True)
            st.caption("💡 操作提示：复制左侧带有 3 列表格的完整 Prompt，在 Teams 窗口中发送给 Bot。")
        st.markdown('</div>', unsafe_allow_html=True)

        # ==================== STEP 3 ====================
        st.markdown('<div class="bds-card">', unsafe_allow_html=True)
        st.markdown('<span class="bds-step-badge">STEP 3</span> <h4 style="display:inline; margin-left:8px; color:#005691;">粘贴 Bot 回复并一键生成最终交付包</h4>', unsafe_allow_html=True)
        
        col_in, col_sup = st.columns([3, 2])
        with col_in:
            bot_reply = st.text_area(
                "📥 在此粘贴 M-PU Bot 润色后的完整回复：",
                height=220,
                placeholder="粘贴 Bot 输出的包含 0. Abstract, 1. Product/Process, 2. Problem, 3. Lessons (包含3列表格), 4. Potentially affected 的完整文本..."
            )
        with col_sup:
            selected_sups = st.multiselect("👥 选择收件供应商 (自动读取 Vendor code 邮箱):", options=list(supplier_dict.keys()))
            to_emails_list = []
            for s in selected_sups:
                to_emails_list.extend(supplier_dict[s])
            to_emails_str = "; ".join(list(set(to_emails_list)))
            if to_emails_str:
                st.info(f"📧 **自动收件人:**\n`{to_emails_str}`")

        if st.button("🚀 立即生成标准化 Word 报告与 Outlook 邮件草稿", type="primary", use_container_width=True):
            if template_file is None:
                st.error("❌ 未检测到 Word 模板，请在侧边栏确认路径。")
            else:
                with st.spinner("正在定向装配表格、插入 Picture 图片并生成邮件附件..."):
                    bot_data = parse_bot_feber_response(bot_reply) if bot_reply.strip() else {}
                    
                    # 定向装配 Word 模板
                    doc = populate_docx_exact_tables(template_file, bot_data, selected_row, picture_img)
                    bio = io.BytesIO()
                    doc.save(bio)
                    doc_bytes = bio.getvalue()
                    
                    serial_str = str(selected_row.get(serial_no_col, 'LL-Export'))
                    doc_filename = f"LL_Template_{serial_str}.docx"
                    eml_bytes = generate_eml_file(selected_row, to_emails_str, doc_bytes, doc_filename)
                    
                    st.success("🎉 生成成功！Abstract 已严格对齐，2. Problem 与 Picture 图片已正常写入，3. Lessons 3列表格已按行回填。")
                    
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
        st.markdown('</div>', unsafe_allow_html=True)

    except Exception as e:
        st.error(f"❌ 运行异常: {e}")
else:
    st.info("ℹ️ 请在侧边栏确认 Master List (Excel) 和 Word 模板的文件路径。")
