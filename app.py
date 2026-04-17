import streamlit as st
import pandas as pd
from sqlalchemy import text

# 頁面設定
st.set_page_config(page_title="營收管理系統", layout="wide")

# 建立資料庫連線
conn = st.connection("postgresql", type="sql")

# --- 共用函數 ---
def fetch_data():
    return conn.query("SELECT * FROM financials", ttl=0)

def calculate_row_total(df):
    months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
    return df[months].sum(axis=1)

# --- 主選單 ---
tabs = st.tabs(["1. 各專案推進營收", "2. 數據管理庫", "3. 營收分類彙整表"])

# --- Tab 1: 各專案推進營收 (卡片呈現) ---
with tabs[0]:
    st.header("📈 專案營收推進看板")
    df = fetch_data()
    if not df.empty:
        df['年度總額'] = calculate_row_total(df)
        projects = df['專案說明'].unique()
        
        cols = st.columns(3)
        for i, project in enumerate(projects):
            with cols[i % 3]:
                p_data = df[df['專案說明'] == project]
                target_rev = p_data[p_data['紀錄類型'] == '收入']['年度總額'].sum()
                est_rev = p_data[p_data['紀錄類型'] == '收入預估']['年度總額'].sum()
                prog_rate = (est_rev / target_rev * 100) if target_rev != 0 else 0
                cat = p_data['營收分類'].iloc[0]
                
                with st.container(border=True):
                    st.subheader(f"{project}")
                    st.caption(f"分類: {cat}")
                    st.metric("目標營收 (收入)", f"${target_rev:,.0f}")
                    st.metric("預估收入", f"${est_rev:,.0f}", f"{prog_rate:.1f}% 推進率")
                    st.progress(min(prog_rate/100, 1.0))
    else:
        st.info("目前無資料，請先至分頁 2 匯入 Excel")

# --- Tab 2: 數據管理庫 (匯入與刪除) ---
with tabs[1]:
    st.header("🗄️ 資料庫維護")
    
    # Excel 匯入區
    uploaded_file = st.file_uploader("匯入營收 Excel (需符合格式)", type=["xlsx"])
    if uploaded_file:
        new_df = pd.read_excel(uploaded_file)
        if st.button("確認寫入資料庫"):
            with conn.session as session:
                for _, row in new_df.iterrows():
                    # 這裡根據你提供的 SQL 欄位寫入
                    insert_sql = text("""
                        INSERT INTO financials ("專案說明", "Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec", "營收分類", "紀錄類型", "說明")
                        VALUES (:p, :m1, :m2, :m3, :m4, :m5, :m6, :m7, :m8, :m9, :m10, :m11, :m12, :cat, :type, :desc)
                    """)
                    session.execute(insert_sql, {
                        "p": row['專案說明'], "m1": row['Jan'], "m2": row['Feb'], "m3": row['Mar'],
                        "m4": row['Apr'], "m5": row['May'], "m6": row['Jun'], "m7": row['Jul'],
                        "m8": row['Aug'], "m9": row['Sep'], "m10": row['Oct'], "m11": row['Nov'],
                        "m12": row['Dec'], "cat": row['營收分類'], "type": row['紀錄類型'], "desc": row['說明']
                    })
                session.commit()
            st.success("資料已匯入！")
            st.rerun()

    # 資料編輯與批次刪除
    df_manage = fetch_data()
    if not df_manage.empty:
        # 使用 data_editor 進行手動更新
        edited_df = st.data_editor(df_manage, key="db_editor", num_rows="dynamic", use_container_width=True)
        
        # 標示紅字邏輯 (CSS 模擬)
        st.markdown("""<style>.red-text { color: red; font-weight: bold; }</style>""", unsafe_allow_width=True)
        
        selected_ids = st.multiselect("選擇要刪除的 ID", df_manage['id'].tolist())
        if st.button("🗑️ 執行批次刪除", type="primary"):
            if selected_ids:
                with conn.session as session:
                    session.execute(text("DELETE FROM financials WHERE id IN :ids"), {"ids": tuple(selected_ids)})
                    session.commit()
                st.rerun()

# --- Tab 3: 營收分類彙整表 ---
with tabs[2]:
    st.header("📊 分類彙整分析")
    df_all = fetch_data()
    if not df_all.empty:
        df_all['年度總額'] = calculate_row_total(df_all)
        
        # 建立彙整邏輯
        summary = df_all.groupby(['營收分類', '紀錄類型'])['年度總額'].sum().unstack(fill_value=0)
        
        # 確保所有欄位都存在，避免報錯
        required_cols = ['收入', '收入預估', '支出', '支出預估']
        for c in required_cols:
            if c not in summary.columns: summary[c] = 0
            
        summary['目標毛利'] = summary['收入'] - summary['支出']
        summary['預估毛利'] = summary['收入預估'] - summary['支出預估']
        summary['預估毛利率'] = (summary['預估毛利'] / summary['收入預估']).apply(lambda x: f"{x:.1%}" if x != 0 else "0%")
        summary['差異'] = summary['收入'] - summary['收入預估']
        
        st.table(summary[['收入', '收入預估', '目標毛利', '預估毛利', '預估毛利率', '差異']])
    else:
        st.warning("暫無資料可彙整")
