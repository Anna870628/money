import streamlit as st
import pandas as pd
import numpy as np

# --- 1. 初始化 PostgreSQL 連線 ---
# Streamlit 會自動從 Secrets 裡的 [connections.postgresql] 讀取配置
conn = st.connection("postgresql", type="sql")

st.set_page_config(layout="wide", page_title="車聯網營收整理系統")

# --- 2. 資料存取函數 ---
def fetch_data():
    # 使用 conn.query 直接執行 SQL
    return conn.query("SELECT * FROM revenue_records", ttl="10m")

def upload_to_db(df):
    # 使用 sqlalchemy 引擎寫回資料庫
    with conn.session as session:
        # 這裡建議先清空再寫入，或視您的需求而定
        session.execute("DELETE FROM revenue_records")
        df.to_sql("revenue_records", con=conn.engine, if_exists="append", index=False)
        session.commit()

# --- 3. 介面設定 ---
tab1, tab2, tab3 = st.tabs(["各專案推進營收", "資料庫管理", "營收分類彙整表"])

df = fetch_data()

# --- 分頁內容 (邏輯同前，僅連線方式改變) ---

with tab1:
    if not df.empty:
        st.subheader("📊 專案進度卡片")
        # 繪製卡片的邏輯...
        # 範例：
        projects = df['project_name'].unique()
        cols = st.columns(3)
        for idx, p in enumerate(projects):
            p_df = df[df['project_name'] == p]
            target = p_df[p_df['row_type'] == "收入"]['total'].sum()
            est = p_df[p_df['row_type'] == "收入預估"]['total'].sum()
            rate = (est/target*100) if target != 0 else 0
            with cols[idx % 3]:
                st.metric(label=p, value=f"{est:,.0f}", delta=f"推進率 {rate:.1f}%")

with tab2:
    if not df.empty:
        st.subheader("🛠️ 手動更正與批次刪除")
        edited_df = st.data_editor(df, num_rows="dynamic", key="data_editor")
        if st.button("💾 儲存所有變更至資料庫"):
            upload_to_db(edited_df)
            st.success("資料庫已同步更新！")
            st.rerun()

with tab3:
    if not df.empty:
        st.subheader("📈 營收分類彙整 (毛利計算)")
        # 分組計算邏輯... (使用 pandas 處理 df)
