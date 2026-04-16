import streamlit as st
import pandas as pd
from sqlalchemy import text
import openpyxl

# ==========================================
# 0. 密碼驗證 (CMX_BPT)
# ==========================================
def check_password():
    def password_entered():
        if st.session_state["password"] == st.secrets["passwords"]["admin_password"]:
            st.session_state["password_correct"] = True
            del st.session_state["password"]
        else:
            st.session_state["password_correct"] = False
    if "password_correct" not in st.session_state:
        st.title("🔒 專案管理系統 - 登入")
        st.text_input("請輸入存取密碼", type="password", on_change=password_entered, key="password")
        return False
    return st.session_state.get("password_correct", False)

if check_password():
    st.set_page_config(page_title="車聯網營收系統 v7", layout="wide")
    conn = st.connection("postgresql", type="sql")

    # ==========================================
    # 1. 核心資料處理：包含顏色抓取
    # ==========================================
    def load_data():
        try:
            return conn.query("SELECT * FROM financials", ttl="0")
        except:
            return pd.DataFrame(columns=['專案說明', '紀錄類型'] + ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'] + ['營收分類', '顏色標記', '說明'])

    def save_to_supabase(df):
        with conn.session as session:
            session.execute(text("DELETE FROM financials")) 
            df.to_sql('financials', conn.engine, if_exists='append', index=False)
            session.commit()

    def process_excel_with_colors(uploaded_file):
        """讀取 Excel 並抓取儲存格底色"""
        # 1. 使用 openpyxl 讀取顏色
        wb = openpyxl.load_workbook(uploaded_file, data_only=True)
        sheet = wb.active
        
        color_data = []
        # 假設資料從第 4 行開始 (skiprows=2 的意思)
        for row in sheet.iter_rows(min_row=4):
            # 抓取「專案說明」那一格的顏色 (假設在第 2 欄)
            cell_color = row[1].fill.start_color.index
            # 轉換為 Hex 或是簡單描述
            color_data.append(str(cell_color) if cell_color != '00000000' else "無底色")

        # 2. 回到 pandas 讀取數值
        uploaded_file.seek(0)
        df = pd.read_excel(uploaded_file, skiprows=2)
        
        df['顏色標記'] = color_data[:len(df)]
        df.rename(columns={df.columns[2]: '紀錄類型'}, inplace=True)
        df['專案說明'] = df['專案說明'].ffill()
        
        months = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec']
        for m in months:
            df[m] = pd.to_numeric(df[m], errors='coerce').fillna(0)
            
        target_cols = ['專案說明', '紀錄類型'] + months + ['營收分類', '顏色標標記', '說明']
        return df[df['專案說明'].notna()][target_cols]

    # ==========================================
    # 2. 介面呈現
    # ==========================================
    tab_edit, tab_summary = st.tabs(["📝 數據編輯與匯入", "📈 顏色加總與分析"])

    with tab_edit:
        with st.sidebar:
            st.header("📂 數據匯入")
            file = st.file_uploader("匯入帶有底色的 Excel", type=["xlsx"])
            if file and st.button("🚀 開始解析顏色與數據"):
                processed_df = process_excel_with_colors(file)
                save_to_supabase(processed_df)
                st.success("匯入完成！已自動抓取底色標記。")
                st.rerun()
            
            st.divider()
            if st.button("⚠️ 清空資料庫 (重來)"):
                with conn.session as session:
                    session.execute(text("DELETE FROM financials"))
                    session.commit()
                st.warning("資料庫已清空")
                st.rerun()

        data = load_data()
        st.subheader("📝 營收明細編輯 (支援刪除列)")
        st.caption("提示：點擊最左側選取列，按下鍵盤 Delete 鍵可刪除。")
        
        edited_df = st.data_editor(
            data,
            num_rows="dynamic",
            use_container_width=True,
            height=500,
            column_config={
                "營收分類": st.column_config.SelectboxColumn("營收分類", options=["24DCM開發/維運", "TOYOTA聯網服務", "LEXUS聯網服務", "其他"]),
                "紀錄類型": st.column_config.SelectboxColumn("紀錄類型", options=["收入", "支出", "收入預估", "支出預估"]),
                "顏色標記": st.column_config.TextColumn("Excel底色代碼", disabled=True)
            }
        )
        if st.button("💾 儲存變更", type="primary"):
            save_to_supabase(edited_df)
            st.success("雲端存儲成功！")

    with tab_summary:
        st.header("🎨 基於 Excel 底色的自動加總分析")
        if not edited_df.empty:
            months = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec']
            df_sum = edited_df.copy()
            for m in months:
                df_sum[m] = pd.to_numeric(df_sum[m], errors='coerce').fillna(0)
            df_sum['年度總計'] = df_sum[months].sum(axis=1)

            # 針對顏色進行加總
            color_summary = df_sum.groupby(['顏色標記', '紀錄類型'])['年度總計'].sum().unstack().fillna(0)
            
            st.subheader("📊 不同顏色區塊的營收規模")
            st.dataframe(color_summary.style.format("{:,.0f}"), use_container_width=True)
            
            # 差異 Pick-up
            st.divider()
            st.subheader("🔍 異常檢視 (目標 vs 預估)")
            # 此處保留之前的差異計算邏輯...
