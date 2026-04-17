import streamlit as st
import pandas as pd
from sqlalchemy import text

# 1. 頁面配置
st.set_page_config(page_title="營收管理系統 v4", layout="wide")

# 套用樣式優化
st.markdown("""
    <style>
    .stMetric { background-color: #f8f9fa; padding: 10px; border-radius: 8px; border-left: 5px solid #007bff; }
    div[data-testid="stExpander"] { border: 1px solid #ff4b4b; border-radius: 5px; }
    </style>
""", unsafe_allow_html=True)

# 建立資料庫連線
try:
    conn = st.connection("postgresql", type="sql")
except Exception as e:
    st.error("❌ 無法連線至資料庫")
    st.stop()

# --- 核心運算函數 ---
def get_clean_data():
    df = conn.query('SELECT * FROM financials', ttl=0)
    if not df.empty:
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

# ==========================================
# ⬅️ 左側邊欄：匯入與刪除 (控制區)
# ==========================================
with st.sidebar:
    st.title("🛠️ 資料管理控制台")
    st.subheader("1. 匯入最新營收")
    uploaded_file = st.file_uploader("請上傳 Excel 檔案", type=["xlsx"], help="請確保欄位包含：專案說明、營收分類、紀錄類型、及 1-12 月欄位")
    
    if uploaded_file:
        try:
            new_df = pd.read_excel(uploaded_file)
            new_df.columns = new_df.columns.str.strip()
            st.success("✅ 檔案已讀取")
            if st.button("🚀 確定寫入資料庫", use_container_width=True):
                with conn.session as s:
                    for _, r in new_df.iterrows():
                        query = text("""
                            INSERT INTO financials ("專案說明", "Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec", "營收分類", "紀錄類型", "說明")
                            VALUES (:p, :m1, :m2, :m3, :m4, :m5, :m6, :m7, :m8, :m9, :m10, :m11, :m12, :cat, :type, :desc)
                        """)
                        s.execute(query, {
                            "p": r.get('專案說明'), "cat": r.get('營收分類'), "type": r.get('紀錄類型'), "desc": r.get('說明',''),
                            "m1":r.get('Jan',0), "m2":r.get('Feb',0), "m3":r.get('Mar',0), "m4":r.get('Apr',0), "m5":r.get('May',0), "m6":r.get('Jun',0),
                            "m7":r.get('Jul',0), "m8":r.get('Aug',0), "m9":r.get('Sep',0), "m10":r.get('Oct',0), "m11":r.get('Nov',0), "m12":r.get('Dec',0)
                        })
                    s.commit()
                st.toast("數據匯入成功！", icon="🎉")
                st.rerun()
        except Exception as e:
            st.error(f"讀取失敗: {e}")

    st.divider()
    
    st.subheader("2. 刪除與清理")
    raw_data_for_del = get_clean_data()
    if not raw_data_for_del.empty:
        # 批次刪除
        selected_ids = st.multiselect("勾選欲刪除 ID", raw_data_for_del['id'].unique().tolist())
        if st.button("🗑️ 執行批次刪除", type="primary", use_container_width=True) and selected_ids:
            with conn.session as s:
                s.execute(text("DELETE FROM financials WHERE id IN :ids"), {"ids": tuple(selected_ids)})
                s.commit()
            st.rerun()
        
        # 全表清空 (放在 expander 內防止誤觸)
        with st.expander("💣 危險：清空資料庫"):
            st.warning("此動作會刪除 DB 內所有專案資料")
            if st.button("確認全表清空", use_container_width=True):
                with conn.session as s:
                    s.execute(text("TRUNCATE TABLE financials;"))
                    s.commit()
                st.rerun()
    else:
        st.caption("目前 DB 為空，無可刪除資料")

# ==========================================
# 🏠 主畫面：看板與報表
# ==========================================
st.title("📈 營收戰情室")
df = get_clean_data()

if df.empty:
    st.info("👋 你好！請先在左側邊欄匯入 Excel 檔案以開始分析。")
else:
    # 預運算年度總額
    df['年度總額'] = calculate_yearly_sum(df)
    
    tab1, tab2, tab3 = st.tabs(["🚀 專案進度看板", "📊 分類彙整報表", "📝 原始數據檢視"])

    # --- Tab 1: 專案進度看板 ---
    with tab1:
        # 頂部過濾器
        all_cats = sorted(df['營收分類'].unique().tolist())
        selected_cat = st.multiselect("🔍 篩選分類", all_cats, default=all_cats)
        
        display_df = df[df['營收分類'].isin(selected_cat)]
        projects = display_df['專案說明'].unique()
        
        cols = st.columns(2)
        for i, project in enumerate(projects):
            p_rows = display_df[display_df['專案說明'] == project]
            
            # 使用 contains 進行模糊比對並清理空白
            target_rev = p_rows[p_rows['紀錄類型'].str.contains('^收入$', na=False)]['年度總額'].sum()
            est_rev = p_rows[p_rows['紀錄類型'].str.contains('收入預估', na=False)]['年度總額'].sum()
            cat_name = p_rows['營收分類'].iloc[0]
            
            rate = (est_rev / target_rev) if target_rev > 0 else 0
            
            with cols[i % 2]:
                with st.container(border=True):
                    # 標題與分類
                    header_c1, header_c2 = st.columns([3, 1])
                    header_c1.subheader(f"{project}")
                    header_c1.caption(f"分類：{cat_name}")
                    
                    # 推進率圓環模擬 (文字)
                    header_c2.markdown(f"### `{rate:.1%}`")
                    
                    # 數據指標
                    m1, m2, m3 = st.columns(3)
                    m1.metric("目標營收", f"${target_rev:,.0f}")
                    m2.metric("預估收入", f"${est_rev:,.0f}")
                    m3.metric("差異", f"${est_rev-target_rev:,.0f}", delta_color="normal")
                    
                    # 進度條
                    st.progress(min(rate, 1.0) if rate >= 0 else 0)

    # --- Tab 2: 分類彙整報表 ---
    with tab2:
        st.subheader("📊 各類別營收彙整總表")
        all_cats_in_db = df['營收分類'].unique()
        
        # 樞紐分析
        pivot = df.pivot_table(
            index='營收分類', 
            columns='紀錄類型', 
            values='年度總額', 
            aggfunc='sum'
        ).reindex(all_cats_in_db).fillna(0)
        
        # 強制補齊欄位，避免 KeyError
        cols_needed = ['收入', '收入預估', '支出', '支出預估']
        for c in cols_needed:
            if c not in pivot.columns: pivot[c] = 0

        # 計算報表
        report = pd.DataFrame(index=pivot.index)
        report['目標收入'] = pivot['收入']
        report['預估收入'] = pivot['收入預估']
        report['目標毛利'] = pivot['收入'] - pivot['支出']
        report['預估毛利'] = pivot['收入預估'] - pivot['支出預估']
        report['預估毛利率'] = (report['預估毛利'] / report['預估收入']).fillna(0)
        report['營收缺口'] = report['目標收入'] - report['預估收入']

        # 樣式設定
        st.dataframe(
            report.style.format({
                '目標收入': '{:,.0f}', '預估收入': '{:,.0f}',
                '目標毛利': '{:,.0f}', '預估毛利': '{:,.0f}',
                '預估毛利率': '{:.2%}', '營收缺口': '{:,.0f}'
            }).applymap(lambda x: 'color: #ff4b4b; font-weight: bold' if x > 0 else 'color: #28a745', subset=['營收缺口']),
            use_container_width=True,
            height=450
        )

    # --- Tab 3: 原始數據檢視 (方便快速檢查哪裡填錯) ---
    with tab3:
        st.subheader("📝 資料庫原始內容檢視")
        st.data_editor(df, use_container_width=True, num_rows="dynamic")
