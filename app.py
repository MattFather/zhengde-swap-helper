import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Cm, Pt
from docx.enum.section import WD_ORIENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import io
import datetime
import subprocess
import tempfile
import os
import streamlit.components.v1 as components
import base64

# ================= 1. 頁面基本設定與 JS 快捷鍵 =================
st.set_page_config(
    page_title="正德調課與列印整合系統",
    page_icon="🏫",
    layout="wide"
)

components.html(
    """
    <script>
    const doc = window.parent.document;
    doc.addEventListener('keydown', function(event) {
        if (event.key.toLowerCase() === 'c') {
            const activeElement = doc.activeElement;
            const isInput = activeElement.tagName === 'INPUT' || activeElement.tagName === 'TEXTAREA';
            if (!isInput) {
                event.preventDefault();
                event.stopPropagation();
                event.stopImmediatePropagation();
            }
        }
    }, true); 
    </script>
    """,
    height=0,
    width=0,
)

# ================= 2. 核心資料載入 =================
@st.cache_data
def load_data():
    return pd.read_csv("schedule.csv")

try:
    df = load_data()
except FileNotFoundError:
    st.error("找不到 schedule.csv 檔案。")
    st.stop()

# ================= 3. 智慧混合調課引擎 (雙人+三角) =================
def analyze_swap_options(df, my_name, target_class, my_day, my_period):
    class_schedule = df[df['Class'] == target_class]
    direct_swaps = {}
    triangle_swaps = {}
    
    for _, row_c in class_schedule.iterrows():
        teacher_c = row_c['Teacher']
        day_c = row_c['Day']
        period_c = row_c['Period']
        subj_c = row_c['Subject']
        
        if teacher_c == my_name: continue
        
        # 條件 1: 我在 C 的時段必須有空 (我才能去上他的課)
        if not df[(df['Teacher'] == my_name) & (df['Day'] == day_c) & (df['Period'] == period_c)].empty:
            continue
            
        # 檢查是否能「直接雙人互換」：C 在我的時段必須有空
        c_busy_t1 = not df[(df['Teacher'] == teacher_c) & (df['Day'] == my_day) & (df['Period'] == my_period)].empty
        
        if not c_busy_t1:
            direct_swaps[(day_c, period_c)] = {'Teacher': teacher_c, 'Subject': subj_c}
        else:
            # 無法直接互換，啟動「三角調課」搜索，尋找橋樑 B 老師
            valid_b_paths = []
            for _, row_b in class_schedule.iterrows():
                teacher_b = row_b['Teacher']
                day_b = row_b['Day']
                period_b = row_b['Period']
                subj_b = row_b['Subject']
                
                if teacher_b in (my_name, teacher_c): continue
                
                # 橋樑 B 必須在我的時段有空 (B 來幫我上課)
                b_busy_t1 = not df[(df['Teacher'] == teacher_b) & (df['Day'] == my_day) & (df['Period'] == my_period)].empty
                if b_busy_t1: continue
                
                # 目標 C 必須在橋樑 B 的時段有空 (C 去幫 B 上課)
                c_busy_t3 = not df[(df['Teacher'] == teacher_c) & (df['Day'] == day_b) & (df['Period'] == period_b)].empty
                if c_busy_t3: continue
                
                valid_b_paths.append({
                    'Teacher': teacher_b, 'Day': day_b, 'Period': period_b, 'Subject': subj_b
                })
                
            if valid_b_paths:
                triangle_swaps[(day_c, period_c)] = {
                    'Teacher': teacher_c, 'Subject': subj_c, 'Paths': valid_b_paths
                }
                
    return direct_swaps, triangle_swaps

def create_schedule_grid(df, teacher_name):
    t_df = df[df['Teacher'] == teacher_name].copy()
    if t_df.empty: return pd.DataFrame()
    t_df = t_df.drop_duplicates(subset=['Period', 'Day'])
    t_df['Cell'] = t_df['Class'].astype(str) + "班\n" + t_df['Subject']
    grid = t_df.pivot(index='Period', columns='Day', values='Cell')
    all_days = ['Mon', 'Tue', 'Wed', 'Thu', 'Fri']
    all_periods = list(range(1, 8)) 
    grid = grid.reindex(index=all_periods, columns=all_days).fillna("")
    grid.index = [f"第 {i} 節" for i in all_periods]
    grid.columns = ['星期一', '星期二', '星期三', '星期四', '星期五']
    return grid

def get_next_weekday(day_zh):
    day_map = {'星期一': 0, '星期二': 1, '星期三': 2, '星期四': 3, '星期五': 4}
    target_wd = day_map.get(day_zh, 0)
    today = datetime.date.today()
    days_ahead = target_wd - today.weekday()
    if days_ahead <= 0: days_ahead += 7
    return today + datetime.timedelta(days_ahead)

# ================= 4. 列印系統核心 (含字體微調) =================
def docx_to_pdf(docx_bytes):
    with tempfile.TemporaryDirectory() as tmpdir:
        docx_path = os.path.join(tmpdir, "temp.docx")
        pdf_path = os.path.join(tmpdir, "temp.pdf")
        with open(docx_path, "wb") as f: f.write(docx_bytes)
        try:
            subprocess.run(["libreoffice", "--headless", "--convert-to", "pdf", "--outdir", tmpdir, docx_path], check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
            if os.path.exists(pdf_path):
                with open(pdf_path, "rb") as f: return f.read()
            return None
        except Exception: return None

def set_cell_border(cell, **kwargs):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcBorders = tcPr.find(qn('w:tcBorders'))
    if tcBorders is None:
        tcBorders = OxmlElement('w:tcBorders')
        tcPr.append(tcBorders)
    for side in ["top", "left", "bottom", "right"]:
        if side in kwargs:
            tag = 'w:{}'.format(side)
            element = tcBorders.find(qn(tag))
            if element is not None: tcBorders.remove(element)
            element = OxmlElement(tag)
            for key, val in kwargs[side].items():
                element.set(qn('w:{}'.format(key)), str(val))
            tcBorders.append(element)

def set_chinese_font(doc, font_name='標楷體'):
    doc.styles['Normal'].font.name = font_name
    doc.styles['Normal']._element.rPr.rFonts.set(qn('w:eastAsia'), font_name)

def add_run_with_autofit(paragraph, text, default_pt=9):
    text_len = len(text)
    run = paragraph.add_run(text)
    if text_len <= 4: run.font.size = Pt(default_pt)
    elif text_len <= 6: run.font.size = Pt(default_pt - 1.5)
    else: run.font.size = Pt(default_pt - 2.5)
    run.bold = True
    return run

def generate_timetable_block(container_cell, title_suffix, sch_year, sch_term, issue_unit, class_label, filtered_df, is_teacher_side=True, teacher_name=""):
    p_header = container_cell.paragraphs[0]
    p_header.paragraph_format.space_before = Pt(0)
    p_header.paragraph_format.space_after = Pt(0)
    p_header.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_h = p_header.add_run(f"新北市立正德國民中學 {sch_year}學年度第{sch_term}學期\n調/代 課單")
    run_h.bold = True
    run_h.font.size = Pt(14) 

    p_sub = container_cell.add_paragraph()
    p_sub.paragraph_format.space_before = Pt(0)
    p_sub.paragraph_format.space_after = Pt(0)
    tab_stops = p_sub.paragraph_format.tab_stops
    tab_stops.add_tab_stop(Cm(13.32), WD_TAB_ALIGNMENT.RIGHT)
    left_text = f"教師：{teacher_name}" if teacher_name else ""
    run_sub = p_sub.add_run(f"{left_text}\t班級：{class_label}")
    run_sub.bold = True
    run_sub.font.size = Pt(12) 

    inner_table = container_cell.add_table(rows=9, cols=6)
    inner_table.style = 'Table Grid'
    inner_table.autofit = False 
    inner_widths = [Cm(2.22), Cm(2.22), Cm(2.22), Cm(2.22), Cm(2.22), Cm(2.22)]
    for j, width in enumerate(inner_widths):
        inner_table.columns[j].width = width
        for cell in inner_table.columns[j].cells: cell.width = width
    
    weekdays_list = ["一", "二", "三", "四", "五"]
    h_cells = inner_table.rows[0].cells
    p_tl = h_cells[0].paragraphs[0]
    p_tl.text = ""
    run_tl = p_tl.add_run("節次/星期")
    run_tl.font.size = Pt(11)
    
    for i, day in enumerate(weekdays_list):
        p_day = h_cells[i+1].paragraphs[0]
        p_day.text = ""
        run_day = p_day.add_run(day)
        run_day.font.size = Pt(12)

    periods_list = ["1", "2", "3", "4", "5", "6", "7", "8"]
    times_list = ["08:20-09:05", "09:15-10:00", "10:10-10:55", "11:05-11:50", "13:00-13:45", "13:55-14:40", "15:00-15:45", "15:55-16:40"]
    for r_idx in range(8):
        row = inner_table.rows[r_idx + 1]
        row.height = Cm(1.5)
        cell_p = row.cells[0].paragraphs[0]
        cell_p.paragraph_format.space_after = Pt(0)
        run_num = cell_p.add_run(periods_list[r_idx])
        run_num.bold = True
        run_num.font.size = Pt(12) 
        cell_p.add_run("\n")
        run_time = cell_p.add_run(times_list[r_idx])
        run_time.font.size = Pt(9) 

    day_map = {0: 1, 1: 2, 2: 3, 3: 4, 4: 5} 
    cell_records = {}
    
    for _, row_data in filtered_df.iterrows():
        if pd.notnull(row_data["日期"]) and row_data["日期"] != "":
            d_idx = day_map.get(pd.to_datetime(row_data["日期"]).weekday())
            try: p_idx = int(str(row_data["節次"]).split()[1])
            except: continue
            if d_idx and p_idx:
                k = (p_idx, d_idx)
                if k not in cell_records: cell_records[k] = []
                cell_records[k].append(row_data)
                
    for (p_idx, d_idx), records in cell_records.items():
        cell = inner_table.rows[p_idx].cells[d_idx]
        cell.text = "" 
        
        actual_classes = [r for r in records if str(r.get("調/代課")) != "空堂X"]
        x_marks = [r for r in records if str(r.get("調/代課")) == "空堂X"]
        
        if actual_classes:
            for idx, row_data in enumerate(actual_classes):
                if idx == 0: p1 = cell.paragraphs[0]
                else:
                    cell.add_paragraph()
                    p1 = cell.add_paragraph()
                    
                c_name = str(row_data["班級"]).strip() if pd.notnull(row_data["班級"]) and row_data["班級"] != "" else ""
                s_name = str(row_data["科目"]).strip() if pd.notnull(row_data["科目"]) and row_data["科目"] != "" else ""
                
                p1.paragraph_format.space_after = Pt(0)
                run_date_cell = p1.add_run(pd.to_datetime(row_data["日期"]).strftime("%m/%d"))
                run_date_cell.font.size = Pt(9)
                run_date_cell.bold = True
                
                p2 = cell.add_paragraph()
                p2.paragraph_format.space_after = Pt(0)
                
                if is_teacher_side and c_name:
                    run_c = p2.add_run(f"{c_name} ")
                    run_c.font.size = Pt(9)
                    run_c.bold = True
                
                if s_name:
                    run_s = p2.add_run(s_name)
                    run_s.bold = True
                    if len(s_name) > 4: run_s.font.size = Pt(7.5) 
                    elif len(s_name) == 4: run_s.font.size = Pt(8.0) 
                    else: run_s.font.size = Pt(9.0) 
                
                p3 = cell.add_paragraph()
                p3.paragraph_format.space_after = Pt(0)
                run_teacher = p3.add_run(str(row_data["老師"]))
                run_teacher.font.size = Pt(9)
                run_teacher.bold = True
                
                p4 = cell.add_paragraph()
                p4.paragraph_format.space_after = Pt(0)
                pair_id = str(row_data.get("配對編號", "")).strip()
                
                if str(row_data["調/代課"]) == "代課":
                    run_type = p4.add_run("[代課]")
                    run_type.font.size = Pt(8)
                else:
                    if title_suffix == "存查聯" and pair_id: run_type = p4.add_run(f"[{pair_id}]")
                    else: run_type = p4.add_run("[調課]")
                    run_type.font.size = Pt(8)
                            
        elif x_marks and "教師通知聯" in title_suffix:
            row_data = x_marks[0]
            p1 = cell.paragraphs[0]
            p1.paragraph_format.space_after = Pt(0)
            run_date_cell = p1.add_run(pd.to_datetime(row_data["日期"]).strftime("%m/%d"))
            run_date_cell.font.size = Pt(9)
            run_date_cell.bold = True
            
            p2 = cell.add_paragraph()
            p2.paragraph_format.space_after = Pt(0)
            run_x = p2.add_run("✖")
            run_x.font.size = Pt(14)
            run_x.bold = True
            
            p3 = cell.add_paragraph()
            p3.paragraph_format.space_after = Pt(0)
            target_info = str(row_data.get("原資訊", "")).strip()
            run_text = p3.add_run(target_info if target_info else "(已調走)")
            run_text.font.size = Pt(8)
            run_text.bold = True

    for r in range(9):
        for c in range(6):
            curr_cell = inner_table.rows[r].cells[c]
            curr_cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            for para in curr_cell.paragraphs:
                para.alignment = WD_ALIGN_PARAGRAPH.CENTER
                para.paragraph_format.line_spacing = 1.0
                para.paragraph_format.space_before = Pt(0)
                para.paragraph_format.space_after = Pt(0)
            if r == 4: set_cell_border(curr_cell, bottom={"sz": 24, "val": "single", "color": "000000"})
            if r == 5: set_cell_border(curr_cell, top={"sz": 24, "val": "single", "color": "000000"})

    print_p = container_cell.paragraphs[-1] 
    print_p.text = ""
    print_p.paragraph_format.space_before = Pt(0) 
    print_p.paragraph_format.space_after = Pt(0)
    print_p.alignment = WD_ALIGN_PARAGRAPH.LEFT 
    tab_stops_print = print_p.paragraph_format.tab_stops
    tab_stops_print.add_tab_stop(Cm(13.32), WD_TAB_ALIGNMENT.RIGHT)
    run_issue = print_p.add_run(f"發放單位：{issue_unit}")
    run_issue.font.size = Pt(10)
    run_date = print_p.add_run(f"\t列印：{datetime.date.today().strftime('%Y/%m/%d')}")
    run_date.font.size = Pt(10)

    footer_p = container_cell.add_paragraph()
    footer_p.paragraph_format.space_before = Pt(0)
    footer_p.paragraph_format.space_after = Pt(0)
    footer_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_f = footer_p.add_run(f"({title_suffix})")
    run_f.bold = True
    run_f.font.size = Pt(10)

def process_swap_logic(df):
    df_result = []
    subs = df[df["調/代課"] == "代課"].copy()
    for _, r in subs.iterrows(): df_result.append(r)
        
    swaps = df[df["調/代課"] == "調課"].copy()
    pair_ids = [p for p in swaps["配對編號"].unique() if pd.notnull(p) and str(p).strip() != ""]
    
    for pid in pair_ids:
        rows = swaps[swaps["配對編號"] == pid].sort_index()
        n = len(rows)
        if n >= 2:
            orig_dates = rows["日期"].tolist()
            orig_periods = rows["節次"].tolist()
            shifted_dates = orig_dates[1:] + [orig_dates[0]]
            shifted_periods = orig_periods[1:] + [orig_periods[0]]
            
            for i in range(n):
                new_row = rows.iloc[i].copy()
                new_row["日期"] = shifted_dates[i]
                new_row["節次"] = shifted_periods[i]
                df_result.append(new_row)
                
                x_row = rows.iloc[i].copy()
                x_row["日期"] = orig_dates[i]
                x_row["節次"] = orig_periods[i]
                x_row["調/代課"] = "空堂X"
                target_date = shifted_dates[i]
                target_period = shifted_periods[i]
                try:
                    t_date_str = pd.to_datetime(target_date).strftime('%m/%d') if pd.notnull(target_date) and target_date != "" else ""
                    t_p_num = "".join(filter(str.isdigit, str(target_period)))
                    if t_date_str and t_p_num: x_row["原資訊"] = f"調 {t_date_str}[{t_p_num}]"
                    else: x_row["原資訊"] = "(已調走)"
                except: x_row["原資訊"] = "(已調走)"
                df_result.append(x_row)
        else:
            for _, r in rows.iterrows(): df_result.append(r)
    
    no_id = swaps[swaps["配對編號"].isna() | (swaps["配對編號"] == "")]
    for _, r in no_id.iterrows(): df_result.append(r)
    return pd.DataFrame(df_result)

def create_docx(sch_year, sch_term, issue_unit, edited_df):
    doc = Document()
    section = doc.sections[0]
    section.orient = WD_ORIENT.LANDSCAPE
    section.page_width = Cm(29.7)
    section.page_height = Cm(21.0)
    section.left_margin = Cm(0.8)
    section.right_margin = Cm(0.5)
    section.top_margin = section.bottom_margin = Cm(0.5)
    set_chinese_font(doc, '標楷體')

    df_raw = edited_df[edited_df["勾選列印資料"] == True].copy()
    if df_raw.empty: return None
    
    df_raw["配對編號"] = df_raw["配對編號"].fillna("").astype(str).str.strip()
    df_raw["班級"] = df_raw["班級"].fillna("").astype(str).str.strip()
    df_raw["老師"] = df_raw["老師"].fillna("").astype(str).str.strip()
    df_processed = process_swap_logic(df_raw)

    all_blocks = []
    classes = sorted(list(set([c for c in df_processed["班級"] if c != ""])))
    all_blocks.append({"suffix": "存查聯", "label": ", ".join(classes), "df": df_processed, "is_teacher": True, "teacher_name": ""})

    teachers = sorted(list(set([t for t in df_processed["老師"] if t != ""])))
    for t in teachers:
        df_t = df_processed[df_processed["老師"] == t]
        t_classes = sorted(list(set([c for c in df_t["班級"] if c != ""])))
        all_blocks.append({"suffix": "教師通知聯", "label": ", ".join(t_classes), "df": df_t, "is_teacher": True, "teacher_name": f"{t}老師"})

    for c in classes:
        df_c = df_processed[df_processed["班級"] == c]
        all_blocks.append({"suffix": "班級公告聯", "label": c, "df": df_c, "is_teacher": False, "teacher_name": ""})

    for i in range(0, len(all_blocks), 2):
        if i > 0: doc.add_page_break()
        table = doc.add_table(rows=1, cols=4)
        table.autofit = False
        col_widths = [Cm(13.7), Cm(0.5), Cm(0.5), Cm(13.7)]
        for j in range(4):
            table.columns[j].width = col_widths[j]
            for cell in table.columns[j].cells: cell.width = col_widths[j]
        set_cell_border(table.cell(0, 1), right={"sz": 6, "val": "dashed", "color": "808080"})

        b1 = all_blocks[i]
        generate_timetable_block(table.cell(0, 0), b1["suffix"], sch_year, sch_term, issue_unit, b1["label"], b1["df"], is_teacher_side=b1["is_teacher"], teacher_name=b1["teacher_name"])
        if i + 1 < len(all_blocks):
            b2 = all_blocks[i+1]
            generate_timetable_block(table.cell(0, 3), b2["suffix"], sch_year, sch_term, issue_unit, b2["label"], b2["df"], is_teacher_side=b2["is_teacher"], teacher_name=b2["teacher_name"])

    bio = io.BytesIO()
    doc.save(bio)
    return bio.getvalue()


# ================= 5. 系統狀態記憶與初始化 =================
state_keys = [
    "last_user_name", "source_class", "source_subject", "source_period", "source_day_en", "source_day_zh", 
    "target_teacher", "target_subject", "target_period", "target_day_en", "target_day_zh", "last_clicked_cell",
    "swap_mode", "triangle_paths"
]
for key in state_keys:
    if key not in st.session_state: st.session_state[key] = None

if 'res_data' not in st.session_state:
    st.session_state.res_data = pd.DataFrame({
        "勾選列印資料": pd.Series(dtype='bool'),
        "配對編號": pd.Series(dtype='str'),
        "班級": pd.Series(dtype='str'),
        "日期": pd.Series(dtype='datetime64[ns]'),
        "節次": pd.Series(dtype='str'),
        "科目": pd.Series(dtype='str'),
        "老師": pd.Series(dtype='str'),
        "調/代課": pd.Series(dtype='str')
    })

def check_source_conflict(df, teacher, date_val, period):
    if df.empty: return False
    date_str = pd.to_datetime(date_val).strftime('%Y-%m-%d')
    df_dates = pd.to_datetime(df['日期'], errors='coerce').dt.strftime('%Y-%m-%d')
    conflict = df[(df['老師'] == teacher) & (df_dates == date_str) & (df['節次'] == period)]
    return not conflict.empty

def check_destination_conflict(df, teacher, date_val, period):
    if df.empty: return False
    processed_df = process_swap_logic(df)
    if processed_df.empty: return False
    date_str = pd.to_datetime(date_val).strftime('%Y-%m-%d')
    p_dates = pd.to_datetime(processed_df['日期'], errors='coerce').dt.strftime('%Y-%m-%d')
    conflict = processed_df[
        (processed_df['老師'] == teacher) & 
        (p_dates == date_str) & 
        (processed_df['節次'] == period) &
        (processed_df['調/代課'] != '空堂X')
    ]
    return not conflict.empty

def style_my_grid(val):
    val_str = str(val)
    if "🌟[可互調]" in val_str: return "color: #0066cc; background-color: #f0f8ff;" 
    elif "🔺[三角調]" in val_str: return "color: #cc6600; background-color: #fff3e6;" 
    elif "🔄[欲調走]" in val_str: return "color: #d9534f; font-weight: bold; background-color: #fdf5f5;"
    return ""

def style_target_grid(val):
    val_str = str(val)
    if "🌟[您去上]" in val_str: return "color: #0066cc; font-weight: bold; background-color: #f0f8ff;"
    elif "🔄[來代課]" in val_str: return "color: #d9534f; font-weight: bold; background-color: #fdf5f5;"
    elif "🔄[去代課]" in val_str: return "color: #d9534f; font-weight: bold; background-color: #fdf5f5;"
    elif "🔺[去代課]" in val_str: return "color: #cc6600; font-weight: bold; background-color: #fff3e6;"
    elif "🔺[來代課]" in val_str: return "color: #cc6600; font-weight: bold; background-color: #fff3e6;"
    return ""

# ================= 6. UI 版面佈局 =================
st.title("🏫 正德調課小幫手 ＆ 列印整合系統")

# 【完美統一版】回歸兩個分頁
tab_visual, tab_print = st.tabs(["🔄 第一步：智慧調課與配對", "🖨️ 第二步：列印單據與輸出"])

# ----------------- Tab 1: 第一步：智慧調課與配對 -----------------
with tab_visual:
    all_teachers = sorted(df['Teacher'].dropna().unique())
    my_name = st.selectbox("🙋‍♂️ 請輸入您的名字：", all_teachers, index=None, placeholder="請選擇您的名字...")
    
    st.markdown("---")

    if my_name:
        if my_name != st.session_state.last_user_name:
            for k in state_keys: st.session_state[k] = None
            st.session_state.last_user_name = my_name

        day_en = ['Mon', 'Tue', 'Wed', 'Thu', 'Fri']
        day_zh = ['星期一', '星期二', '星期三', '星期四', '星期五']
        day_map_rev = dict(zip(day_zh, day_en))

        original_my_grid = create_schedule_grid(df, my_name)
        display_grid = original_my_grid.copy()
        
        # 預先計算可用方案
        direct_swaps, triangle_swaps = {}, {}
        if st.session_state.source_class:
            direct_swaps, triangle_swaps = analyze_swap_options(
                df, my_name, st.session_state.source_class, 
                st.session_state.source_day_en, st.session_state.source_period
            )
            
            # 標示自己準備調走的課
            source_r = st.session_state.source_period - 1
            source_c = day_zh.index(st.session_state.source_day_zh)
            orig_val = original_my_grid.iloc[source_r, source_c]
            display_grid.iloc[source_r, source_c] = f"🔄[欲調走]\n{orig_val}"
            
            # 標示可以互換的目標 (藍色)
            for (d_day, d_period), info in direct_swaps.items():
                r = d_period - 1
                c = day_en.index(d_day)
                display_grid.iloc[r, c] = f"🌟[可互調]\n{info['Teacher']}"
                
            # 標示只能三角換的目標 (橘色)
            for (t_day, t_period), info in triangle_swaps.items():
                if (t_day, t_period) not in direct_swaps:
                    r = t_period - 1
                    c = day_en.index(t_day)
                    display_grid.iloc[r, c] = f"🔺[三角調]\n{info['Teacher']}"

        col_left, col_right = st.columns([1, 1], gap="large")

        with col_left:
            st.subheader(f"📅 【{my_name}老師】的課表")
            st.info("🎯 **操作說明：**\n1️⃣ 點擊有課的格子 -> 選擇要調走的課。\n2️⃣ 點擊空堂的格子 -> 選擇目標落點 (藍色=可互調，橘色=可三角調)。")
            
            try: styled_display_grid = display_grid.style.map(style_my_grid)
            except AttributeError: styled_display_grid = display_grid.style.applymap(style_my_grid)

            event = st.dataframe(
                styled_display_grid,
                use_container_width=True,
                height=400,
                on_select="rerun",
                selection_mode="single-cell",
                key="my_schedule_grid"
            )

            selection = event.selection.cells
            if selection:
                cell = selection[0]
                try:
                    r_val, c_val = (cell[0], cell[1]) if isinstance(cell, (tuple, list)) else (cell.get('row', 0) if isinstance(cell, dict) else getattr(cell, 'row', 0), cell.get('column', 0) if isinstance(cell, dict) else getattr(cell, 'column', 0))
                    c_idx = day_zh.index(str(c_val)) if str(c_val) in day_zh else int(c_val)
                    t_day_zh = day_zh[c_idx]
                    t_day_en = day_map_rev.get(t_day_zh, "Mon")
                    
                    if isinstance(r_val, str) and "第" in r_val:
                        t_period = int(r_val.replace("第 ", "").replace(" 節", ""))
                        r_idx = t_period - 1
                    else:
                        r_idx = int(r_val)
                        t_period = r_idx + 1

                    clicked_id = f"{t_day_en}_{t_period}"

                    if st.session_state.last_clicked_cell != clicked_id:
                        orig_content = original_my_grid.iloc[r_idx, c_idx]
                        if orig_content != "":
                            # 第一階段：點了有課的格子，設定來源
                            match_data = df[(df['Teacher'] == my_name) & (df['Day'] == t_day_en) & (df['Period'] == t_period)]
                            if not match_data.empty:
                                st.session_state.source_class = match_data.iloc[0]['Class']
                                st.session_state.source_subject = match_data.iloc[0]['Subject']
                                st.session_state.source_period = t_period
                                st.session_state.source_day_en = t_day_en
                                st.session_state.source_day_zh = t_day_zh
                                st.session_state.swap_mode = None
                                st.session_state.target_teacher = None
                                st.session_state.last_clicked_cell = clicked_id
                                st.rerun()
                        else:
                            # 第二階段：點了空堂格子，確認落點方案
                            if st.session_state.source_class:
                                if (t_day_en, t_period) in direct_swaps:
                                    st.session_state.swap_mode = 'direct'
                                    st.session_state.target_teacher = direct_swaps[(t_day_en, t_period)]['Teacher']
                                    st.session_state.target_subject = direct_swaps[(t_day_en, t_period)]['Subject']
                                    st.session_state.target_period = t_period
                                    st.session_state.target_day_en = t_day_en
                                    st.session_state.target_day_zh = t_day_zh
                                    st.session_state.last_clicked_cell = clicked_id
                                    st.rerun()
                                elif (t_day_en, t_period) in triangle_swaps:
                                    st.session_state.swap_mode = 'triangle'
                                    st.session_state.target_teacher = triangle_swaps[(t_day_en, t_period)]['Teacher']
                                    st.session_state.target_subject = triangle_swaps[(t_day_en, t_period)]['Subject']
                                    st.session_state.triangle_paths = triangle_swaps[(t_day_en, t_period)]['Paths']
                                    st.session_state.target_period = t_period
                                    st.session_state.target_day_en = t_day_en
                                    st.session_state.target_day_zh = t_day_zh
                                    st.session_state.last_clicked_cell = clicked_id
                                    st.rerun()
                except Exception: pass
            else: st.session_state.last_clicked_cell = None

        with col_right:
            if st.session_state.swap_mode == 'direct':
                st.subheader(f"✨ 雙人互調方案 (目標: {st.session_state.target_teacher}老師)")
                target_grid = create_schedule_grid(df, st.session_state.target_teacher)
                
                tr_source = st.session_state.source_period - 1
                tc_source = day_en.index(st.session_state.source_day_en)
                target_grid.iloc[tr_source, tc_source] = f"🔄[來代課]\n{st.session_state.source_class}班\n{st.session_state.source_subject}"
                
                tr_target = st.session_state.target_period - 1
                tc_target = day_en.index(st.session_state.target_day_en)
                target_grid.iloc[tr_target, tc_target] = f"🌟[您去上]\n{st.session_state.source_class}班\n{st.session_state.source_subject}"
                
                try: styled_target_grid = target_grid.style.map(style_target_grid)
                except AttributeError: styled_target_grid = target_grid.style.applymap(style_target_grid)
                st.dataframe(styled_target_grid, use_container_width=True, height=280)
                
                col_d1, col_d2, col_btn = st.columns([2, 2, 1.5])
                with col_d1: date_mine = st.date_input(f"您的原上課日 ({st.session_state.source_day_zh})", value=get_next_weekday(st.session_state.source_day_zh))
                with col_d2: date_target = st.date_input(f"對方原上課日 ({st.session_state.target_day_zh})", value=get_next_weekday(st.session_state.target_day_zh))
                with col_btn:
                    st.markdown("<br>", unsafe_allow_html=True)
                    if st.button("➕ 一鍵加入", type="primary", use_container_width=True):
                        source_period_str = f"第 {st.session_state.source_period} 節"
                        target_period_str = f"第 {st.session_state.target_period} 節"
                        if check_source_conflict(st.session_state.res_data, my_name, date_mine, source_period_str):
                            st.error(f"⚠️ 衝堂：您 {date_mine} 的課已加入過清單！")
                        elif check_source_conflict(st.session_state.res_data, st.session_state.target_teacher, date_target, target_period_str):
                            st.error(f"⚠️ 衝堂：對方 {date_target} 的課已加入過清單！")
                        elif check_destination_conflict(st.session_state.res_data, my_name, date_target, target_period_str):
                            st.error(f"⚠️ 目標衝堂：您在 {date_target} {target_period_str} 已有排定其他課程！")
                        elif check_destination_conflict(st.session_state.res_data, st.session_state.target_teacher, date_mine, source_period_str):
                            st.error(f"⚠️ 目標衝堂：對方在 {date_mine} {source_period_str} 已有排定其他課程！")
                        else:
                            current_ids = pd.to_numeric(st.session_state.res_data["配對編號"], errors='coerce').dropna()
                            next_id = str(int(current_ids.max() + 1)) if not current_ids.empty else "1"
                            new_rows = pd.DataFrame([
                                {"勾選列印資料": True, "配對編號": next_id, "班級": st.session_state.source_class, 
                                 "日期": pd.to_datetime(date_mine), "節次": source_period_str, 
                                 "科目": str(st.session_state.source_subject).strip(), "老師": my_name, "調/代課": "調課"},
                                {"勾選列印資料": True, "配對編號": next_id, "班級": st.session_state.source_class, 
                                 "日期": pd.to_datetime(date_target), "節次": target_period_str, 
                                 "科目": str(st.session_state.target_subject).strip(), "老師": st.session_state.target_teacher, "調/代課": "調課"}
                            ])
                            st.session_state.res_data = pd.concat([st.session_state.res_data, new_rows], ignore_index=True)
                            st.success("✅ 已加入清單！請至第二頁查看。")

            elif st.session_state.swap_mode == 'triangle':
                paths = st.session_state.triangle_paths
                st.subheader(f"✨ 三角調課方案 (目標: {st.session_state.target_teacher}老師的時段)")
                
                path_options = {f"由 {p['Teacher']} 老師代上您的課 (橋樑)": p for p in paths}
                selected_opt = st.selectbox("🎯 系統為您找到以下橋樑老師，請選擇方案：", list(path_options.keys()))
                bridge = path_options[selected_opt]
                bridge_day_zh = [k for k, v in day_map_rev.items() if v == bridge['Day']][0]
                
                # 顯示目標 C 的課表
                grid_c = create_schedule_grid(df, st.session_state.target_teacher)
                r_c_gives = st.session_state.target_period - 1
                c_c_gives = day_en.index(st.session_state.target_day_en)
                grid_c.iloc[r_c_gives, c_c_gives] = f"🌟[您去上]\n(給 {my_name})"
                
                r_c_takes = bridge['Period'] - 1
                c_c_takes = day_en.index(bridge['Day'])
                grid_c.iloc[r_c_takes, c_c_takes] = f"🔄[去代課]\n(替 {bridge['Teacher']})"
                
                # 顯示橋樑 B 的課表
                grid_b = create_schedule_grid(df, bridge['Teacher'])
                r_b_gives = bridge['Period'] - 1
                c_b_gives = day_en.index(bridge['Day'])
                grid_b.iloc[r_b_gives, c_b_gives] = f"🔺[來代課]\n(給 {st.session_state.target_teacher})"
                
                r_b_takes = st.session_state.source_period - 1
                c_b_takes = day_en.index(st.session_state.source_day_en)
                grid_b.iloc[r_b_takes, c_b_takes] = f"🔄[去代課]\n(替 {my_name})"

                try: 
                    sc = grid_c.style.map(style_target_grid)
                    sb = grid_b.style.map(style_target_grid)
                except AttributeError: 
                    sc = grid_c.style.applymap(style_target_grid)
                    sb = grid_b.style.applymap(style_target_grid)
                
                tc1, tc2 = st.columns(2)
                with tc1:
                    st.caption(f"👀 對方({st.session_state.target_teacher}) 變化")
                    st.dataframe(sc, use_container_width=True, height=250)
                with tc2:
                    st.caption(f"👀 橋樑({bridge['Teacher']}) 變化")
                    st.dataframe(sb, use_container_width=True, height=250)
                
                st.markdown("---")
                col_t1, col_t2, col_t3, col_tbtn = st.columns([1, 1, 1, 1])
                with col_t1: date_mine = st.date_input("您的原上課日", value=get_next_weekday(st.session_state.source_day_zh))
                with col_t2: date_c = st.date_input(f"對方({st.session_state.target_teacher}) 原上課日", value=get_next_weekday(st.session_state.target_day_zh))
                with col_t3: date_b = st.date_input(f"橋樑({bridge['Teacher']}) 原上課日", value=get_next_weekday(bridge_day_zh))
                with col_tbtn:
                    st.markdown("<br>", unsafe_allow_html=True)
                    if st.button("➕ 三角一鍵加入", type="primary", use_container_width=True):
                        p_mine_str = f"第 {st.session_state.source_period} 節"
                        p_c_str = f"第 {st.session_state.target_period} 節"
                        p_b_str = f"第 {bridge['Period']} 節"
                        
                        if check_source_conflict(st.session_state.res_data, my_name, date_mine, p_mine_str) or \
                           check_source_conflict(st.session_state.res_data, st.session_state.target_teacher, date_c, p_c_str) or \
                           check_source_conflict(st.session_state.res_data, bridge['Teacher'], date_b, p_b_str):
                            st.error("⚠️ 衝堂：此方案有老師的時段已加入清單！")
                        elif check_destination_conflict(st.session_state.res_data, my_name, date_c, p_c_str) or \
                             check_destination_conflict(st.session_state.res_data, st.session_state.target_teacher, date_b, p_b_str) or \
                             check_destination_conflict(st.session_state.res_data, bridge['Teacher'], date_mine, p_mine_str):
                            st.error("⚠️ 目標衝堂：目標日調入後會發生衝堂！")
                        else:
                            # 加入順序必須是 A(使用者) -> C(目標) -> B(橋樑)，列印系統才能正確將C的課給A，B的課給C，A的課給B
                            current_ids = pd.to_numeric(st.session_state.res_data["配對編號"], errors='coerce').dropna()
                            next_id = str(int(current_ids.max() + 1)) if not current_ids.empty else "1"
                            new_rows = pd.DataFrame([
                                {"勾選列印資料": True, "配對編號": next_id, "班級": st.session_state.source_class, 
                                 "日期": pd.to_datetime(date_mine), "節次": p_mine_str, 
                                 "科目": str(st.session_state.source_subject).strip(), "老師": my_name, "調/代課": "調課"},
                                {"勾選列印資料": True, "配對編號": next_id, "班級": st.session_state.source_class, 
                                 "日期": pd.to_datetime(date_c), "節次": p_c_str, 
                                 "科目": str(st.session_state.target_subject).strip(), "老師": st.session_state.target_teacher, "調/代課": "調課"},
                                {"勾選列印資料": True, "配對編號": next_id, "班級": st.session_state.source_class, 
                                 "日期": pd.to_datetime(date_b), "節次": p_b_str, 
                                 "科目": str(bridge['Subject']).strip(), "老師": bridge['Teacher'], "調/代課": "調課"}
                            ])
                            st.session_state.res_data = pd.concat([st.session_state.res_data, new_rows], ignore_index=True)
                            st.success("✅ 三角方案已成功加入！")

            elif st.session_state.source_class:
                st.info("👈 請在左側點擊 🌟 或 🔺 標記的空堂來選擇對象。")
            else:
                st.info("👈 準備好了嗎？請先在左側課表點選一堂您想調走的課。")

# ----------------- Tab 2: 🖨️ 第二步：列印單據與輸出 -----------------
with tab_print:
    c1, c2, c3 = st.columns(3)
    with c1: sch_year = st.text_input("學年度", value="114")
    with c2: sch_term = st.selectbox("學期", ["一", "二"], index=1)
    with c3: issue_unit = st.text_input("發放單位", value="ＯＯＯ老師")

    df_subs = df['Subject'].dropna().astype(str).str.strip().unique().tolist()
    base_subs = ["", "國文", "英文", "數學", "生物", "理化", "地科", "地理", "歷史", "公民", 
                 "體育", "健康", "視藝", "表藝", "音樂", "家政", "童軍", "輔導", "資訊", "生科", "本土語"]
    subject_list = list(dict.fromkeys(base_subs + df_subs))

    st.markdown("#### 📝 待列印清單編輯區")
    
    if not st.session_state.res_data.empty:
        st.session_state.res_data["日期"] = pd.to_datetime(st.session_state.res_data["日期"], errors='coerce')
        st.session_state.res_data["勾選列印資料"] = st.session_state.res_data["勾選列印資料"].astype(bool)
        for col in ["配對編號", "班級", "節次", "科目", "老師", "調/代課"]:
            st.session_state.res_data[col] = st.session_state.res_data[col].fillna("").astype(str)
            st.session_state.res_data.loc[st.session_state.res_data[col] == "nan", col] = ""
    else:
        st.session_state.res_data = pd.DataFrame({
            "勾選列印資料": pd.Series(dtype='bool'),
            "配對編號": pd.Series(dtype='str'),
            "班級": pd.Series(dtype='str'),
            "日期": pd.Series(dtype='datetime64[ns]'),
            "節次": pd.Series(dtype='str'),
            "科目": pd.Series(dtype='str'),
            "老師": pd.Series(dtype='str'),
            "調/代課": pd.Series(dtype='str')
        })

    edited_df = st.data_editor(
        st.session_state.res_data,
        column_config={
            "勾選列印資料": st.column_config.CheckboxColumn("勾選"),
            "配對編號": st.column_config.TextColumn("配對編號"),
            "班級": st.column_config.TextColumn("班級"),
            "日期": st.column_config.DateColumn("日期", format="MM/DD"),
            "節次": st.column_config.SelectboxColumn("節次", options=[f"第 {i} 節" for i in range(1, 9)]),
            "科目": st.column_config.SelectboxColumn("科目", options=subject_list),
            "老師": st.column_config.TextColumn("老師"),
            "調/代課": st.column_config.SelectboxColumn("調/代課", options=["調課", "代課"]),
        },
        num_rows="dynamic",
        use_container_width=True,
        hide_index=True,
        column_order=("勾選列印資料", "配對編號", "班級", "日期", "節次", "科目", "老師", "調/代課")
    )
    
    st.session_state.res_data = edited_df

    c_download, _ = st.columns([2, 8])
    with c_download:
        csv_bytes = edited_df.to_csv(index=False).encode('utf-8-sig')
        st.download_button(
            label="💾 下載暫存檔",
            data=csv_bytes,
            file_name=f"調代課暫存_{datetime.date.today().strftime('%Y%m%d')}.csv",
            mime="text/csv",
            use_container_width=True
        )

    st.divider()

    st.markdown("#### 🔒 本機資料恢復 (選填)")
    uploaded_file = st.file_uploader("📂 若有先前下載的暫存檔 (.csv)，請在此上傳恢復：", type=["csv"])
    if uploaded_file is not None:
        if 'last_uploaded_id' not in st.session_state or st.session_state.last_uploaded_id != uploaded_file.file_id:
            try:
                df_upload = pd.read_csv(uploaded_file, keep_default_na=False, dtype=str)
                if "日期" in df_upload.columns:
                    df_upload["日期"] = pd.to_datetime(df_upload["日期"], errors='coerce').dt.date
                if "勾選列印資料" in df_upload.columns:
                    df_upload["勾選列印資料"] = df_upload["勾選列印資料"].astype(str).str.lower() == 'true'
                for col in ["配對編號", "班級", "節次", "科目", "老師", "調/代課"]:
                    if col in df_upload.columns: df_upload[col] = df_upload[col].astype(str)
                st.session_state.res_data = df_upload
                st.session_state.last_uploaded_id = uploaded_file.file_id
                st.rerun() 
            except Exception as e:
                st.error(f"❌ 檔案讀取失敗: {e}")

    st.divider()

    if issue_unit.strip() == "ＯＯＯ老師":
        st.error("⚠️ 提醒：請在最上方修改「發放單位」(預設為ＯＯＯ老師) 後，即可解鎖列印與下載功能。")
    else:
        # ================= 列印與轉換輸出 =================
        data_docx = create_docx(sch_year, sch_term, issue_unit, edited_df)

        if data_docx:
            col_word, col_pdf = st.columns([1, 1])
            with col_word:
                st.download_button(
                    label="📥 下載 Word 檔 (可編輯)",
                    data=data_docx,
                    file_name=f"正德調代課單_{datetime.date.today().strftime('%Y%m%d')}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True
                )
                
            with col_pdf:
                if st.button("📥 轉換並下載 PDF (手機建議)", use_container_width=True, type="primary"):
                    with st.spinner("🚀 伺服器正在努力轉換中 (約需 5~10 秒，請耐心等候)..."):
                        pdf_data = docx_to_pdf(data_docx)
                        if pdf_data:
                            st.success("✅ 轉換成功！檔案已自動下載。")
                            b64_pdf = base64.b64encode(pdf_data).decode('utf-8')
                            pdf_filename = f"正德調代課單_{datetime.date.today().strftime('%Y%m%d')}.pdf"
                            
                            auto_download_js = f"""
                                <script>
                                    setTimeout(function() {{
                                        const parentDoc = window.parent.document;
                                        const link = parentDoc.createElement('a');
                                        link.href = 'data:application/octet-stream;base64,{b64_pdf}';
                                        link.download = '{pdf_filename}';
                                        parentDoc.body.appendChild(link);
                                        link.click();
                                        parentDoc.body.removeChild(link);
                                    }}, 300);
                                </script>
                            """
                            components.html(auto_download_js, height=0, width=0)
                            
                            st.download_button(
                                label="備用：若未自動下載請點此",
                                data=pdf_data,
                                file_name=pdf_filename,
                                mime="application/pdf",
                                use_container_width=True
                            )
                        else:
                            st.error("❌ 轉換失敗，伺服器過度繁忙或缺少套件。")