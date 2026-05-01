import streamlit as st
import pandas as pd

# ================= 1. 頁面基本設定 =================
st.set_page_config(
    page_title="正德調課小幫手",
    page_icon="🔄",
    layout="wide"
)

# ================= 2. 核心資料載入 =================
@st.cache_data
def load_data():
    return pd.read_csv("schedule.csv")

# ================= 3. 調課核心演算法 =================
def find_swap_options(df, my_name, target_class, my_day, my_period):
    """尋找所有符合互換條件的課"""
    class_schedule = df[df['Class'] == target_class]
    potential_matches = class_schedule[class_schedule['Teacher'] != my_name]
    
    recommendations = []
    for _, row in potential_matches.iterrows():
        other_teacher = row['Teacher']
        other_day = row['Day']
        other_period = row['Period']
        other_subject = row['Subject']
        
        is_other_busy = df[(df['Teacher'] == other_teacher) & 
                           (df['Day'] == my_day) & 
                           (df['Period'] == my_period)]
        is_me_busy = df[(df['Teacher'] == my_name) & 
                        (df['Day'] == other_day) & 
                        (df['Period'] == other_period)]
        
        if is_other_busy.empty and is_me_busy.empty:
            recommendations.append({
                "Teacher": other_teacher,
                "Subject": other_subject,
                "OtherDay": other_day,
                "OtherPeriod": other_period
            })
    return recommendations

def create_schedule_grid(df, teacher_name):
    """轉換成 5x7 課表網格"""
    t_df = df[df['Teacher'] == teacher_name].copy()
    if t_df.empty:
        return pd.DataFrame()
    
    t_df = t_df.drop_duplicates(subset=['Period', 'Day'])
    t_df['Cell'] = t_df['Class'].astype(str) + "班\n" + t_df['Subject']
    
    grid = t_df.pivot(index='Period', columns='Day', values='Cell')
    all_days = ['Mon', 'Tue', 'Wed', 'Thu', 'Fri']
    all_periods = list(range(1, 8)) 
    grid = grid.reindex(index=all_periods, columns=all_days).fillna("")
    
    grid.index = [f"第 {i} 節" for i in all_periods]
    grid.columns = ['星期一', '星期二', '星期三', '星期四', '星期五']
    return grid

# ================= 4. 系統狀態記憶 =================
if "last_user_name" not in st.session_state:
    st.session_state.last_user_name = None
if "source_class" not in st.session_state:
    st.session_state.source_class = None
    st.session_state.source_period = None
    st.session_state.source_day_en = None
    st.session_state.source_day_zh = None
if "target_teacher" not in st.session_state:
    st.session_state.target_teacher = None
    st.session_state.target_period = None
    st.session_state.target_day_en = None
    st.session_state.target_day_zh = None
if "last_clicked_cell" not in st.session_state:
    st.session_state.last_clicked_cell = None

# ================= 5. 樣式美化函數 (優化視覺) =================
def style_my_grid(val):
    val_str = str(val)
    if "🌟點選換" in val_str:
        # 優化一：將選項文字變淺灰色，背景變極淺灰，避免喧賓奪主
        return "color: #b0b0b0; background-color: #fafafa;" 
    elif "🔄[欲調走]" in val_str:
        # 標記自己想調走的課 (紅字)
        return "color: #d9534f; font-weight: bold; background-color: #fdf5f5;"
    return ""

def style_target_grid(val):
    val_str = str(val)
    if "🌟[" in val_str:
        # 自己去上的課 (藍字)
        return "color: #0066cc; font-weight: bold; background-color: #f0f8ff;"
    elif "🔄[" in val_str:
        # 對方幫忙代的課 (紅字)
        return "color: #d9534f; font-weight: bold; background-color: #fdf5f5;"
    return ""

# ================= 6. 主視覺介面 =================
st.title("🔄 正德調課小幫手")

try:
    df = load_data()
except FileNotFoundError:
    st.error("找不到 schedule.csv 檔案。")
    st.stop()

all_teachers = sorted(df['Teacher'].dropna().unique())
my_name = st.selectbox(
    "🙋‍♂️ 請輸入您的名字：", 
    all_teachers, 
    index=None, 
    placeholder="請選擇您的名字..."
)

st.markdown("---")

if my_name:
    if my_name != st.session_state.last_user_name:
        st.session_state.last_user_name = my_name
        st.session_state.source_class = None
        st.session_state.target_teacher = None
        st.session_state.last_clicked_cell = None

    day_en = ['Mon', 'Tue', 'Wed', 'Thu', 'Fri']
    day_zh = ['星期一', '星期二', '星期三', '星期四', '星期五']
    day_map_rev = dict(zip(day_zh, day_en))

    original_my_grid = create_schedule_grid(df, my_name)
    display_grid = original_my_grid.copy()

    # ====== 邏輯：繪製左側課表 ======
    if st.session_state.source_class:
        swaps = find_swap_options(df, my_name, st.session_state.source_class, st.session_state.source_day_en, st.session_state.source_period)
        
        source_r = st.session_state.source_period - 1
        source_c = day_zh.index(st.session_state.source_day_zh)
        orig_val = original_my_grid.iloc[source_r, source_c]
        display_grid.iloc[source_r, source_c] = f"🔄[欲調走]\n{orig_val}"
        
        for s in swaps:
            r = s['OtherPeriod'] - 1
            c = day_en.index(s['OtherDay'])
            display_grid.iloc[r, c] = f"🌟點選換:\n{s['Teacher']}"

    col_left, col_right = st.columns([1, 1], gap="large")

    with col_left:
        st.subheader(f"📅 【{my_name}老師】的課表")
        st.info("🎯 **操作步驟：**\n1️⃣ 點擊您想調走的班級。\n2️⃣ 點擊出現 **🌟** 的淺色格子選擇老師。")
        
        # 套用自訂義顏色樣式
        try:
            styled_display_grid = display_grid.style.map(style_my_grid)
        except AttributeError:
            styled_display_grid = display_grid.style.applymap(style_my_grid)

        event = st.dataframe(
            styled_display_grid,
            use_container_width=True,
            height=320,
            on_select="rerun",
            selection_mode="single-cell",
            key="my_schedule_grid"
        )

        selection = event.selection.cells
        if selection:
            cell = selection[0]
            try:
                if isinstance(cell, (tuple, list)):
                    r_val, c_val = cell[0], cell[1]
                elif isinstance(cell, dict):
                    r_val, c_val = cell.get('row'), cell.get('column')
                else:
                    r_val, c_val = getattr(cell, 'row', 0), getattr(cell, 'column', 0)

                if str(c_val) in day_zh:
                    t_day_zh = str(c_val)
                    c_idx = day_zh.index(t_day_zh)
                else:
                    c_idx = int(c_val)
                    t_day_zh = display_grid.columns[c_idx]
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
                        match_data = df[(df['Teacher'] == my_name) & (df['Day'] == t_day_en) & (df['Period'] == t_period)]
                        if not match_data.empty:
                            st.session_state.source_class = match_data.iloc[0]['Class']
                            st.session_state.source_period = t_period
                            st.session_state.source_day_en = t_day_en
                            st.session_state.source_day_zh = t_day_zh
                            st.session_state.target_teacher = None
                            st.session_state.last_clicked_cell = clicked_id
                            st.rerun()
                    else:
                        if st.session_state.source_class:
                            swaps = find_swap_options(df, my_name, st.session_state.source_class, st.session_state.source_day_en, st.session_state.source_period)
                            target_found = False
                            for s in swaps:
                                if s['OtherDay'] == t_day_en and s['OtherPeriod'] == t_period:
                                    st.session_state.target_teacher = s['Teacher']
                                    st.session_state.target_period = s['OtherPeriod']
                                    st.session_state.target_day_en = s['OtherDay']
                                    st.session_state.target_day_zh = t_day_zh
                                    target_found = True
                                    break
                            
                            if target_found:
                                st.session_state.last_clicked_cell = clicked_id
                                st.rerun()
            except Exception:
                pass
        else:
            st.session_state.last_clicked_cell = None

        if st.session_state.source_class:
            st.success(f"📍 準備將 **{st.session_state.source_day_zh} 第 {st.session_state.source_period} 節 ({st.session_state.source_class}班)** 調走。")

    # ====== 右側：目標老師的課表預覽 ======
    with col_right:
        if st.session_state.target_teacher:
            st.subheader(f"👀 【{st.session_state.target_teacher}老師】調課後狀態")
            st.info(f"💡 請確認對方是否會因為換課而過度連堂。")
            
            target_grid = create_schedule_grid(df, st.session_state.target_teacher)
            
            # 優化二：直接標上對方與自己的真實姓名
            tr_source = st.session_state.source_period - 1
            tc_source = day_en.index(st.session_state.source_day_en)
            target_grid.iloc[tr_source, tc_source] = f"🔄[{st.session_state.target_teacher}]\n{st.session_state.source_class}班"
            
            tr_target = st.session_state.target_period - 1
            tc_target = day_en.index(st.session_state.target_day_en)
            original_target_val = target_grid.iloc[tr_target, tc_target]
            target_grid.iloc[tr_target, tc_target] = f"🌟[{my_name}]\n{original_target_val}"
            
            # 套用與左側一致的顏色樣式
            try:
                styled_target_grid = target_grid.style.map(style_target_grid)
            except AttributeError:
                styled_target_grid = target_grid.style.applymap(style_target_grid)
                
            # 將右側也改用 st.dataframe 確保左右對齊與高度一致
            st.dataframe(
                styled_target_grid,
                use_container_width=True,
                height=320
            )
            
        elif st.session_state.source_class:
            st.info("👈 請在左側點擊一個帶有 **🌟** 標記的格子，來查看該老師的課表！")
        else:
            st.info("👈 準備好了嗎？請先在左側課表點選一堂您想調走的課。")
else:
    st.info("👋 歡迎使用調課小幫手！請先在上方選單選擇您的名字。")