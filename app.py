import streamlit as st
import pandas as pd
import numpy as np
from sqlalchemy import text
import openpyxl
import re

# ==========================================
# 0. 登入與基礎設定
# ==========================================
def check_password():
    def password_entered():
        if st.session_state["password"] == st.secrets["passwords"]["admin_password"]:
            st.session_state["password_correct"] = True
            del st.session_state["password"]
        else:
            st.session_state["password_correct"] = False
    if "password_correct" not in st.session_state:
        st.title("🔒 營收管理系統 - 登入")
        st.text_input("請輸入密碼", type="password", on_change=password_entered, key="password")
        return False
    return st.session_state.get("password_correct", False)

if check_password():
    st.set_page_config(page_title="車聯網營收戰情系統 v9", layout="wide")
    conn = st.connection("postgresql", type="sql")

    def load_data():
        try:
            df = conn.query("SELECT * FROM financials", ttl="0")
            months = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec']
            for m in months:
                df[m] = pd.to_numeric(df[m], errors='coerce').fillna(0)
            return df
        except:
            return pd.DataFrame()

    def save_to_supabase(df):
        with conn.session as session:
            session.execute(text("DROP TABLE IF EXISTS financials")) 
            df.to_sql('financials', conn.engine, if_exists='replace', index=False)
            session.commit()

    # --- 核心：Excel 顏色抓取邏輯 ---
    def process_imported_file(uploaded_file):
        wb = openpyxl.load_workbook(uploaded_file, data_only=True)
        sheet = wb[wb.sheetnames[0]]
        
        color_list = []
        for row in sheet.iter_rows(min_row=4):
            # 抓取紀錄類型那一格的底色 (通常是第三欄)
            sc = row[2].fill.start_color
            color_id = "無底色"
            if sc.type == 'rgb' and sc.rgb:
                color_id = f"#{str(sc.rgb)[-6:].upper()}"
            elif sc.type == 'theme':
                color_id = f"Theme_{sc.theme}"
            color_list.append(color_id)

        uploaded_file.seek(0)
        df = pd.read_excel(uploaded_file, skiprows=2)
        df.columns = [str(c).strip() for c in df.columns]
        
        # 強制填補分類與專案名稱
        df['專案說明'] = df['專案說明'].ffill()
        df['營收分類'] = df['營收分類'].ffill().fillna("其他")
        df['顏色標記'] = color_list[:len(df)]
        
        return df.dropna(subset=['紀錄類型'])

    # ==========================================
    # 2. 介面呈現
    # ==========================================
    st.title("📊 專案營收戰情系統 v9")
    data = load_data()

    tab_cards, tab_summary, tab_import = st.tabs(["🎴 專案戰情卡片", "📈 分類匯總報表", "📥 數據管理"])

    with tab_cards:
        if not data.empty:
            projects = data['專案說明'].unique()
            cols = st.columns(2)
            for idx, proj in enumerate(projects):
                with cols[idx % 2]:
                    p_df = data[data['專案說明'] == proj]
                    # v9 初步定義：粉紅底色 (#F2DCDB) 為預估
                    is_est = p_df['顏色標記'].str.contains('F2DCDB', na=False) | p_df['紀錄類型'].str.contains('預估', na=False)
                    is_inc = p_df['紀錄類型'].str.contains('收入', na=False)
                    
                    months = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec']
                    target_rev = p_df[is_inc & ~is_est][months].sum().sum()
                    est_rev = p_df[is_inc & is_est][months].sum().sum()
                    
                    with st.container(border=True):
                        st.markdown(f"### {proj}")
                        m1, m2 = st.columns(2)
                        m1.metric("目標營收", f"{target_rev:,.0f}")
                        m2.metric("預估營收", f"{est_rev:,.0f}", f"{est_rev-target_rev:,.0f}")

    with tab_summary:
        if not data.empty:
            # v9 的簡易彙整邏輯
            summary = data.groupby('營收分類')[['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec']].sum()
            st.dataframe(summary)

    with tab_import:
        f = st.file_uploader("匯入新檔案", type=["xlsx"])
        if f and st.button("上傳並取代資料庫"):
            df_new = process_imported_file(f)
            save_to_supabase(df_new)
            st.success("v9 引擎已完成資料匯入")
