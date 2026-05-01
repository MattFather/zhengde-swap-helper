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

# ================= 4. 系統狀態記憶 (核心靈魂) =================
# 這裡負責記住您「點擊了哪一節課」，才能支援連續點擊
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
    st.session_state.target_subject = None

# ================= 5. 主視覺介面 =================
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
    # 換人時清空所有記憶
    if my_name != st.session_state.last_user_name:
        st.session_state.last_user_name = my_name
        st.session_state.source_class = None
        st.session_state.target_teacher = None

    day_en = ['Mon', 'Tue', 'Wed', 'Thu', 'Fri']
    day_zh = ['星期一', '星期二', '星期三', '星期四', '星期五']
    day_map_rev = dict(zip(day_zh, day_en))

    # 取得原始乾淨的課表
    original_my_grid = create_schedule_grid(df, my_name)
    display_grid = original_my_grid.copy()

    # ====== 邏輯一：攔截點擊事件 (解析您點了什麼) ======
    selection = st.session_state.get("my_schedule_grid", {}).get("selection", {}).get("cells", [])
    if selection:
        cell = selection[0]
        try:
            if isinstance(cell, (tuple, list)):
                r_val, c_val = cell[0], cell[1]
            elif isinstance(cell, dict):
                r_val, c_val = cell.get('row'), cell.get('column')
            else:
                r_val, c_val = getattr(cell, 'row', 0), getattr(cell, 'column', 0)
                
            c_val = int(c_val)
            t_day_zh = display_grid.columns[c_val]
            t_day_en = day_map_rev[t_day_zh]
            t_period = int(r_val) + 1
            
            # 對照「原始課表」來判斷您點的是什麼
            orig_content = original_my_grid.iloc[r_val, c_val]
            
            if orig_content != "":
                # 第一步：您點了「有課」的格子 -> 準備調走
                match_data = df[(df['Teacher'] == my_name) & (df['Day'] == t_day_en) & (df['Period'] == t_period)]
                if not match_data.empty:
                    st.session_state.source_class = match_data.iloc[0]['Class']
                    st.session_state.source_period = t_period
                    st.session_state.source_day_en = t_day_en
                    st.session_state.source_day_zh = t_day_zh
                    st.session_state.target_teacher = None # 清空目標，重新選擇
            else:
                # 第二步：您點了「空堂」-> 檢查這是不是可調的方案
                if st.session_state.source_class:
                    swaps = find_swap_options(df, my_name, st.session_state.source_class, st.session_state.source_day_en, st.session_state.source_period)
                    for s in swaps:
                        if s['OtherDay'] == t_day_en and s['OtherPeriod'] == t_period:
                            st.session_state.target_teacher = s['Teacher']
                            st.session_state.target_period = s['OtherPeriod']
                            st.session_state.target_day_en = s['OtherDay']
                            st.session_state.target_subject = s['Subject']
                            break
        except:
            pass

    # ====== 邏輯二：動態繪製您的左側課表 ======
    if st.session_state.source_class:
        swaps = find_swap_options(df, my_name, st.session_state.source_class, st.session_state.source_day_en, st.session_state.source_period)
        
        # 標記欲調走的課
        source_r = st.session_state.source_period - 1
        source_c = day_zh.index(st.session_state.source_day_zh)
        orig_val = original_my_grid.iloc[source_r, source_c]
        display_grid.iloc[source_r, source_c] = f"🔄[欲調走]\n{orig_val}"
        
        # 將「換課選項」直接塞進您的空堂裡！
        for s in swaps:
            r = s['OtherPeriod'] - 1
            c = day_en.index(s['OtherDay'])
            display_grid.iloc[r, c] = f"🌟點擊換:\n{s['Teacher']}"

    # ====== 邏輯三：畫面渲染 ======
    col_left, col_right = st.columns([1, 1], gap="large")

    with col_left:
        st.subheader(f"📅 【{my_name}老師】的課表")
        st.info("🎯 **互動步驟：**\n1️⃣ 點擊您想調走的班級。\n2️⃣ 點擊出現 **🌟** 的空堂，選擇調課對象。")
        
        # 顯示左邊課表
        st.dataframe(
            display_grid,
            use_container_width=True,
            height=320,
            on_select="rerun",
            selection_mode="single-cell",
            key="my_schedule_grid"  # 綁定給系統記憶
        )
        
        if st.session_state.source_class:
            st.success(f"📍 已鎖定：將 **{st.session_state.source_day_zh} 第 {st.session_state.source_period} 節 ({st.session_state.source_class}班)** 調出。")

    with col_right:
        if st.session_state.target_teacher:
            # 顯示目標老師的課表
            st.subheader(f"👀 【{st.session_state.target_teacher}老師】調課後狀態")
            st.info(f"💡 這是**換課完成後**對方的課表，您可藉此確認對方是否會太累。")
            
            target_grid = create_schedule_grid(df, st.session_state.target_teacher)
            
            # 1. 標記他幫我上的課 (對方原本空堂)
            tr_source = st.session_state.source_period - 1
            tc_source = day_en.index(st.session_state.source_day_en)
            target_grid.iloc[tr_source, tc_source] = f"🔄[幫您代]\n{st.session_state.source_class}班"
            
            # 2. 標記我幫他上的課 (對方原本有課，被調走)
            tr_target = st.session_state.target_period - 1
            tc_target = day_en.index(st.session_state.target_day_en)
            original_target_val = target_grid.iloc[tr_target, tc_target]
            target_grid.iloc[tr_target, tc_target] = f"🌟[您去上]\n{original_target_val}"
            
            st.table(target_grid)
        elif st.session_state.source_class:
            st.info("👈 請在左側點擊一個帶有 **🌟** 標記的格子，來查看該老師的課表！")
        else:
            st.info("👈 準備好了嗎？請先在左側課表點選一堂課。")
else:
    st.info("👋 歡迎使用調課小幫手！請先在上方選單選擇您的名字。")