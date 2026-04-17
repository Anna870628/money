import streamlit as st
import pandas as pd
from sqlalchemy import text

# 1. 頁面配置優化：使用廣視角並設定主題色
st.set_page_config(page_title="營收管理系統 v3", layout="wide")

# 套用自定義 CSS 提升閱讀性
st.markdown("""
    <style>
    .stMetric { background-color: #f8f9fa; padding: 15px; border-radius: 10px; border: 1px solid #e9ecef; }
    .project-card { border: 1px solid #ddd; padding: 20px; border-radius: 15px; background: white; margin-bottom: 20px; }
    </style>
""", unsafe_allow_html=True)

# 建立資料庫連線
try:
    conn = st.connection("postgresql", type="sql")
except Exception as e:
    st.error("❌ 連線失敗")
    st.stop()

# --- 強大版核心運算 ---
def get_clean_data():
    """從資料庫抓取資料並自動清理字串空白"""
    df = conn.query('SELECT * FROM financials', ttl=0)
    if not df.empty:
        # 自動清理所有文字欄位的空白，解決「讀不到」的問題
        str_cols = df.select_dtypes(['object']).columns
        for col in str_cols:
            df[col] = df[col].astype(str).str.strip()
    return df

def calculate_yearly_sum(df):
    months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
    temp_df = df.copy()
    for m in months:
        temp_df[m] = pd.to_numeric(temp_df[m], errors='coerce').fillna(0)
    return temp_df[months].sum(axis=1)

# --- 主介面 ---
tab1, tab2, tab3 = st.tabs(["🚀 專案進度看板", "💾 資料庫管理", "📊 分類彙整報表"])

# --- Tab 1: 專案進度看板 (解決顯示不全問題) ---
with tab1:
    df = get_clean_data()
    if not df.empty:
        df['年度總額'] = calculate_yearly_sum(df)
        
        # 側邊過濾器 (方便快速閱讀)
        all_cats = df['營收分類'].unique().tolist()
        selected_cat = st.multiselect("🔍 快速篩選營收分類", all_cats, default=all_cats)
        
        display_df = df[df['營收分類'].isin(selected_cat)]
        projects = display_df['專案說明'].unique()
        
        st.subheader(f"目前顯示專案數：{len(projects)}")
        
        # 視覺化卡片佈局
        cols = st.columns(2) # 改為兩列，讓內容更寬大好讀
        for i, project in enumerate(projects):
            p_rows = display_df[display_df['專案說明'] == project]
            
            # 使用 contains 模糊比對，避免空格或名稱微差
            goal_rev = p_rows[p_rows['紀錄類型'].str.contains('^收入$', na=False)]['年度總額'].sum()
            est_rev = p_rows[p_rows['紀錄類型'].str.contains('收入預估', na=False)]['年度總額'].sum()
            cat = p_rows['營收分類'].iloc[0]
            
            prog_rate = (est_rev / goal_rev) if goal_rev > 0 else 0
            
            with cols[i % 2]:
                with st.container(border=True):
                    c1, c2 = st.columns([2, 1])
                    with c1:
                        st.markdown(f"### {project}")
                        st.caption(f"營收分類：{cat}")
                    with c2:
                        # 顯示推進率大圖標
                        color = "green" if prog_rate >= 0.8 else "orange"
                        st.markdown(f"<h2 style='text-align:right; color:{color};'>{prog_rate:.1%}</h2>", unsafe_allow_html=True)
                    
                    st.divider()
                    mc1, mc2 = st.columns(2)
                    mc1.metric("目標營收 (收入)", f"${goal_rev:,.0f}")
                    mc2.metric("預估營收", f"${est_rev:,.0f}", f"{est_rev-goal_rev:,.0f}")
                    st.progress(min(prog_rate, 1.0) if prog_rate >= 0 else 0)
    else:
        st.warning("⚠️ 系統偵測不到資料。請至「資料庫管理」檢查是否有匯入成功。")

# --- Tab 2: 資料庫管理 (增加診斷功能) ---
with tab2:
    st.header("數據維護與診斷")
    
    # 診斷區：解決你看到的「明明有4個卻只呈現2個」
    with st.expander("🔍 資料健康診斷 (如果數據對不起來，點開看這裡)"):
        raw_df = get_clean_data()
        if not raw_df.empty:
            st.write("1. 目前 DB 總列數:", len(raw_df))
            st.write("2. 現有的營收分類清單:", raw_df['營營分類'].unique().tolist())
            st.write("3. 現有的紀錄類型清單:", raw_df['紀錄類型'].unique().tolist())
            st.write("4. 原始資料預覽:")
            st.dataframe(raw_df[['id', '專案說明', '營收分類', '紀錄類型']].head(10))
        else:
            st.error("DB 目前是空的")

    st.divider()
    
    # 批次刪除與清空按鈕
    c1, c2 = st.columns([1, 1])
    with c1:
        if st.button("🗑️ 勾選批次刪除", use_container_width=True):
            st.info("請在下方表格勾選後按下方的 Delete 鍵 (Streamlit 內建)")
    with c2:
        if st.button("🔥 清空全表資料", type="primary", use_container_width=True):
            with conn.session as s:
                s.execute(text("TRUNCATE TABLE financials;"))
                s.commit()
            st.rerun()

    # 數據編輯器
    st.data_editor(get_clean_data(), key="editor_v3", num_rows="dynamic", use_container_width=True)

# --- Tab 3: 營收分類彙整表 (強化計算邏輯) ---
with tab3:
    st.header("營收彙整分析")
    sum_df = get_clean_data()
    if not sum_df.empty:
        sum_df['年度總額'] = calculate_yearly_sum(sum_df)
        
        # 解決「分類消失」的問題：先建立完整的分類索引
        all_categories = sum_df['營收分類'].unique()
        
        # Pivot Table
        pivot = sum_df.pivot_table(
            index='營收分類', 
            columns='紀錄類型', 
            values='年度總額', 
            aggfunc='sum'
        ).reindex(all_categories).fillna(0) # 確保所有分類都在，沒資料的補 0
        
        # 欄位補齊，防止 KeyError
        for col in ['收入', '收入預估', '支出', '支出預估']:
            if col not in pivot.columns: pivot[col] = 0

        # 計算指標
        report = pd.DataFrame(index=pivot.index)
        report['目標收入'] = pivot['收入']
        report['預估收入'] = pivot['收入預估']
        report['目標毛利'] = pivot['收入'] - pivot['支出']
        report['預估毛利'] = pivot['收入預估'] - pivot['支出預估']
        report['預估毛利率'] = (report['預估毛利'] / report['預估收入']).fillna(0)
        report['差異 (缺口)'] = report['目標收入'] - report['預估收入']

        # 視覺化呈現
        st.subheader("📊 分類彙整彙整總表")
        st.dataframe(
            report.style.format({
                '目標收入': '{:,.0f}', '預估收入': '{:,.0f}',
                '目標毛利': '{:,.0f}', '預估毛利': '{:,.0f}',
                '預估毛利率': '{:.1%}', '差異 (缺口)': '{:,.0f}'
            }).background_gradient(cmap='Blues', subset=['預估毛利率'])
              .applymap(lambda x: 'color: red; font-weight: bold' if x > 0 else 'color: green', subset=['差異 (缺口)']),
            use_container_width=True,
            height=500
        )
