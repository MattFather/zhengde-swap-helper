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
    # 讀取 schedule.csv 檔案
    return pd.read_csv("schedule.csv")

# ================= 3. 調課核心演算法 =================
def find_swap_options(df, my_name, target_class, my_day, my_period):
    """
    尋找該班級中所有符合互換條件的課。
    """
    class_schedule = df[df['Class'] == target_class]
    potential_matches = class_schedule[class_schedule['Teacher'] != my_name]
    
    recommendations = []
    for _, row in potential_matches.iterrows():
        other_teacher = row['Teacher']
        other_day = row['Day']
        other_period = row['Period']
        other_subject = row['Subject']
        
        # 雙方空堂檢查
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
    """將資料轉換成 5x7 的課表網格"""
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

# ================= 4. 主視覺介面 =================
st.title("🔄 正德調課小幫手")

try:
    df = load_data()
except FileNotFoundError:
    st.error("找不到 schedule.csv 檔案。")
    st.stop()

# 選擇名字 (預設選取黃漢鵬老師)
all_teachers = sorted(df['Teacher'].dropna().unique())
default_idx = all_teachers.index("黃漢鵬") if "黃漢鵬" in all_teachers else 0
my_name = st.selectbox("🙋‍♂️ 請輸入您的名字：", all_teachers, index=default_idx)

st.markdown("---")

col_left, col_right = st.columns([1, 1], gap="large")

day_en = ['Mon', 'Tue', 'Wed', 'Thu', 'Fri']
day_zh = ['星期一', '星期二', '星期三', '星期四', '星期五']
day_map_rev = dict(zip(day_zh, day_en))

# --- 左側：操作區 ---
with col_left:
    # 修正一：加上「老師」稱謂
    st.subheader(f"📅 【{my_name}老師】的課表")
    st.info("💡 操作說明：點擊下方課表中的「班級」格子即可開始配對。")
    
    my_grid = create_schedule_grid(df, my_name)
    
    event = st.dataframe(
        my_grid,
        use_container_width=True,
        height=320,
        on_select="rerun",
        selection_mode="single-cell"
    )

    selection = event.selection.cells
    if selection:
        cell = selection[0]
        try:
            # 座標格式解析
            if isinstance(cell, (tuple, list)):
                r_val, c_val = cell[0], cell[1]
            elif isinstance(cell, dict):
                r_val, c_val = cell.get('row'), cell.get('column')
            else:
                r_val, c_val = cell.row, cell.column
                
            t_day_zh = my_grid.columns[int(c_val)] if str(c_val).isdigit() else str(c_val)
            t_day_en = day_map_rev.get(t_day_zh, "Mon")
            t_period = int(str(r_val).replace("第 ", "").replace(" 節", "")) if "第" in str(r_val) else int(r_val) + 1
        except:
            st.error("選取座標解析異常。")
            st.stop()

        cell_content = my_grid.iloc[t_period-1, day_zh.index(t_day_zh)]
        
        if cell_content == "":
            st.warning(f"⚠️ {t_day_zh} 第 {t_period} 節是空堂。")
        else:
            match_data = df[(df['Teacher'] == my_name) & (df['Day'] == t_day_en) & (df['Period'] == t_period)]
            if not match_data.empty:
                t_class = match_data.iloc[0]['Class']
                st.success(f"📍 已鎖定：{t_day_zh} 第 {t_period} 節 ({t_class}班)")
                
                # 搜尋名單
                swaps = find_swap_options(df, my_name, t_class, t_day_en, t_period)
                
                if not swaps:
                    st.error("😢 此班級目前沒有可直接互換的對象。")
                else:
                    st.write("---")
                    # 依照老師歸類結果
                    teacher_groups = {}
                    for s in swaps:
                        if s['Teacher'] not in teacher_groups:
                            teacher_groups[s['Teacher']] = []
                        teacher_groups[s['Teacher']].append(s)

                    st.subheader("🤝 選擇調課對象")
                    selected_teacher = st.radio(
                        "🔍 請點選老師名字查看調課方案：",
                        options=list(teacher_groups.keys()),
                        horizontal=True,
                        index=None,
                        key="teacher_radio"
                    )
                    
                    # --- 右側連動對照 ---
                    if selected_teacher:
                        teacher_slots = teacher_groups[selected_teacher]
                        with col_right:
                            # 修正二：加上「老師」稱謂
                            st.subheader(f"👀 【{selected_teacher}老師】的完整課表")
                            st.info(f"💡 課表中標記 **🌟** 的格子，即為可對調的時段。")
                            
                            target_grid = create_schedule_grid(df, teacher_name=selected_teacher)
                            
                            # 標記符合條件的時段 (🌟)
                            for slot in teacher_slots:
                                r = slot['OtherPeriod'] - 1
                                c = day_en.index(slot['OtherDay'])
                                target_grid.iloc[r, c] += " 🌟"
                            
                            st.table(target_grid)
            else:
                st.error("讀取課程資訊失敗。")
    else:
        st.info("👈 請先點擊左側課表中的課程。")
        with col_right:
            st.info("📊 選取課程後，右側會同步顯示對方的課表。")