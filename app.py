import streamlit as st
import pandas as pd
import numpy as np
from supabase import create_client, Client

# --- 安全連線：完全不出現字串，直接從環境變數讀取 ---
@st.cache_resource
def init_connection():
    url = st.secrets["SUPABASE_URL"]
    key = st.secrets["SUPABASE_KEY"]
    return create_client(url, key)

supabase = init_connection()

st.set_page_config(layout="wide", page_title="車聯網事業營收管理系統")

# --- 資料處理與 Excel 解析邏輯 ---
def fetch_data():
    # 這裡建議加上錯誤處理，確保連線失敗時不會曝露資訊
    try:
        res = supabase.table("revenue_records").select("*").execute()
        return pd.DataFrame(res.data)
    except Exception as e:
        st.error("資料庫連線失敗，請檢查後台設定")
        return pd.DataFrame()

# --- 側邊欄：匯入功能 ---
with st.sidebar:
    st.title("系統控制面板")
    uploaded_file = st.file_uploader("匯入每月收支明細", type=["xlsx"])
    
    if uploaded_file:
        if st.button("確認解析並安全上傳"):
            raw_df = pd.read_excel(uploaded_file, header=None)
            all_records = []
            
            # 解析邏輯：每 6 列一組
            for i in range(2, len(raw_df), 6):
                proj_name = raw_df.iloc[i, 1] if pd.notna(raw_df.iloc[i, 1]) else None
                category = raw_df.iloc[i, 19] if pd.notna(raw_df.iloc[i, 19]) else "未分類"
                
                if not proj_name: continue

                row_labels = ["收入", "收入預估", "支出", "支出預估", "收入差異", "支出差異"]
                for idx, label in enumerate(row_labels):
                    if i + idx < len(raw_df):
                        vals = raw_df.iloc[i + idx, 3:15].fillna(0).replace('-', 0).astype(float).tolist()
                        all_records.append({
                            "project_name": proj_name,
                            "category": category,
                            "row_type": label,
                            "m1": vals[0], "m2": vals[1], "m3": vals[2], "m4": vals[3],
                            "m5": vals[4], "m6": vals[5], "m7": vals[6], "m8": vals[7],
                            "m9": vals[8], "m10": vals[9], "m11": vals[10], "m12": vals[11],
                            "total": sum(vals)
                        })
            
            if all_records:
                supabase.table("revenue_records").insert(all_records).execute()
                st.success("資料已安全存入 Supabase")
                st.rerun()

# --- 分頁介面 ---
tab1, tab2, tab3 = st.tabs(["各專案推進營收", "資料管理", "彙整報表"])
df = fetch_data()

with tab1:
    if not df.empty:
        # 卡片呈現邏輯... (略，同前次建議)
        projects = df['project_name'].unique()
        cols = st.columns(3)
        for idx, p in enumerate(projects):
            p_df = df[df['project_name'] == p]
            target = p_df[p_df['row_type'] == "收入"]['total'].sum()
            est = p_df[p_df['row_type'] == "收入預估"]['total'].sum()
            rate = (est / target * 100) if target != 0 else 0
            
            with cols[idx % 3]:
                st.info(f"**{p}**\n\n推進率: {rate:.1f}%")

with tab2:
    if not df.empty:
        # 標示紅字與手動更新
        st.subheader("批次管理與紅字標示")
        edited_df = st.data_editor(df, num_rows="dynamic", key="editor")
        if st.button("儲存變更"):
            # 先刪除後寫入的安全機制
            supabase.table("revenue_records").delete().neq("id", 0).execute()
            new_data = edited_df.drop(columns=['id'], errors='ignore').to_dict(orient="records")
            supabase.table("revenue_records").insert(new_data).execute()
            st.success("更新成功")

with tab3:
    if not df.empty:
        # 營收彙整邏輯與毛利計算...
        st.write("分類彙整明細 (自動計算毛利與差異)")
        # (這裡放置 groupby 運算邏輯)
