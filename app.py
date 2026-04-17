import streamlit as st
import pandas as pd
from sqlalchemy import text

# 頁面配置
st.set_page_config(page_title="專案營收管理系統", layout="wide")

# 建立資料庫連線
# 注意：這會讀取 .streamlit/secrets.toml 中的 [connections.postgresql]
try:
    conn = st.connection("postgresql", type="sql")
except Exception as e:
    st.error("❌ 無法連線至資料庫，請檢查 Secrets 設定。")
    st.stop()

# --- 核心運算函數 ---
def get_data():
    return conn.query('SELECT * FROM financials', ttl=0)

def calculate_yearly_sum(df):
    months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
    return df[months].sum(axis=1)

# --- 主介面 Tabs ---
tab1, tab2, tab3 = st.tabs(["📊 各專案推進營收", "⚙️ 數據管理庫", "📈 營收分類彙整表"])

# --- Tab 1: 各專案推進營收 (卡片呈現) ---
with tab1:
    st.header("專案進度看板")
    df = get_data()
    if not df.empty:
        df['年度總額'] = calculate_yearly_sum(df)
        projects = df['專案說明'].unique()
        
        cols = st.columns(3)
        for i, project in enumerate(projects):
            with cols[i % 3]:
                p_data = df[df['專案說明'] == project]
                # 篩選目標 (收入) 與 預估 (收入預估)
                goal_rev = p_data[p_data['紀錄類型'] == '收入']['年度總額'].sum()
                est_rev = p_data[p_data['紀錄類型'] == '收入預估']['年度總額'].sum()
                category = p_data['營收分類'].iloc[0]
                
                prog_rate = (est_rev / goal_rev) if goal_rev != 0 else 0
                
                with st.container(border=True):
                    st.subheader(f"📁 {project}")
                    st.caption(f"分類：{category}")
                    st.metric("目標營收 (灰底)", f"${goal_rev:,.0f}")
                    st.metric("預估收入 (粉底)", f"${est_rev:,.0f}", f"{prog_rate:.1%}")
                    st.progress(min(prog_rate, 1.0))
    else:
        st.info("目前尚無資料，請至管理庫匯入 Excel。")

# --- Tab 2: 數據管理庫 (編輯與刪除) ---
with tab2:
    st.header("資料庫維護")
    
    # A. 匯入功能
    with st.expander("📥 匯入新專案 Excel"):
        uploaded_file = st.file_uploader("選擇檔案", type="xlsx")
        if uploaded_file:
            new_df = pd.read_excel(uploaded_file)
            st.dataframe(new_df.head(3))
            if st.button("確認寫入 DB"):
                with conn.session as s:
                    for _, r in new_df.iterrows():
                        query = text("""
                            INSERT INTO financials ("專案說明", "Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec", "營收分類", "紀錄類型", "說明")
                            VALUES (:p, :m1, :m2, :m3, :m4, :m5, :m6, :m7, :m8, :m9, :m10, :m11, :m12, :cat, :type, :desc)
                        """)
                        s.execute(query, {
                            "p": r['專案說明'], "cat": r['營收分類'], "type": r['紀錄類型'], "desc": r.get('說明',''),
                            "m1":r['Jan'], "m2":r['Feb'], "m3":r['Mar'], "m4":r['Apr'], "m5":r['May'], "m6":r['Jun'],
                            "m7":r['Jul'], "m8":r['Aug'], "m9":r['Sep'], "m10":r['Oct'], "m11":r['Nov'], "m12":r['Dec']
                        })
                    s.commit()
                st.success("資料匯入成功！")
                st.rerun()

    # B. 資料管理 (標示紅字邏輯)
    st.subheader("📝 即時數據編輯")
    manage_df = get_data()
    if not manage_df.empty:
        # 標示紅字邏輯：這裡使用 Streamlit 的 column_config 或是直接在 dataframe 做樣式處理
        def color_negative_red(val):
            color = 'red' if isinstance(val, (int, float)) and val < 0 else 'black'
            return f'color: {color}'

        # 批次刪除勾選
        edited_df = st.data_editor(
            manage_df, 
            key="db_editor", 
            num_rows="dynamic",
            use_container_width=True
        )
        
        selected_ids = st.multiselect("選擇要刪除的 ID", manage_df['id'].tolist())
        if st.button("🗑️ 執行批次刪除", type="primary") and selected_ids:
            with conn.session as s:
                s.execute(text("DELETE FROM financials WHERE id IN :ids"), {"ids": tuple(selected_ids)})
                s.commit()
            st.rerun()

# --- Tab 3: 營收分類彙整表 ---
with tab3:
    st.header("營收彙整分析")
    sum_df = get_data()
    if not sum_df.empty:
        sum_df['年度總額'] = calculate_yearly_sum(sum_df)
        
        # 透過 Pivot 整理數據
        pivot = sum_df.pivot_table(
            index='營收分類', 
            columns='紀錄類型', 
            values='年度總額', 
            aggfunc='sum'
        ).fillna(0)
        
        # 確保所有必要列都存在，避免報錯
        for col in ['收入', '收入預估', '支出', '支出預估']:
            if col not in pivot.columns: pivot[col] = 0

        # 計算指標
        report = pd.DataFrame(index=pivot.index)
        report['目標收入'] = pivot['收入']
        report['預估收入'] = pivot['收入預估']
        report['目標毛利'] = pivot['收入'] - pivot['支出']
        report['預估毛利'] = pivot['收入預估'] - pivot['支出預估']
        report['預估毛利率'] = (report['預估毛利'] / report['預估收入']).fillna(0)
        report['差異'] = report['目標收入'] - report['預估收入']

        # 樣式與格式化
        st.dataframe(
            report.style.format({
                '目標收入': '{:,.0f}', '預估收入': '{:,.0f}',
                '目標毛利': '{:,.0f}', '預估毛利': '{:,.0f}',
                '預估毛利率': '{:.1%}', '差異': '{:,.0f}'
            }).applymap(lambda x: 'color: red' if x > 0 else '', subset=['差異']),
            use_container_width=True
        )
    else:
        st.warning("暫無資料計算。")
