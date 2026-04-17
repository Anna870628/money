import streamlit as st
import pandas as pd
import numpy as np
from sqlalchemy import text
import openpyxl
import re

# ==========================================
# 0. 安全驗證
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
        st.text_input("請輸入存取密碼", type="password", on_change=password_entered, key="password")
        return False
    return st.session_state.get("password_correct", False)

if check_password():
    st.set_page_config(page_title="車聯網營收系統 v32", layout="wide")
    conn = st.connection("postgresql", type="sql")

    # ==========================================
    # 🎯 定義你的專屬規則 (組合 A, B, C)
    # ==========================================
    def apply_user_rules(record_type, color_marker):
        """
        根據使用者指定的四種組合進行歸類
        """
        # 判定是否為粉紅色系
        is_pink = any(k in color_marker for k in ["#F2DCDB", "#FCE4D6", "THEME_PINK", "#FFC7CE", "#FAD0C9", "#F8CBAD"])
        # 判定是否為白色/灰色系 (或無底色)
        is_white_grey = color_marker == "無底色" or any(k in color_marker for k in ["#D9D9D9", "#D8D8D8", "#FFFFFF", "THEME_GREY", "#E7E6E6", "#F2F2F2"])

        # 組合 A & B: 收入邏輯
        if "收入" in record_type or "營收" in record_type or "實績" in record_type:
            if is_pink:
                return "🔮 預估收入"
            if is_white_grey:
                return "🎯 原目標收入"
            return "🎯 原目標收入" # 預設

        # 組合 C & D: 支出邏輯
        if "支出" in record_type or "成本" in record_type:
            if is_pink:
                return "💸 預估支出"
            if is_white_grey:
                return "📉 原目標支出"
            return "📉 原目標支出" # 預設
            
        return "❌ 忽略不計"

    # ==========================================
    # 1. 核心讀取與清洗
    # ==========================================
    def load_data():
        try:
            df = conn.query("SELECT * FROM financials", ttl="0")
            months = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec']
            for m in months:
                if m in df.columns:
                    df[m] = pd.to_numeric(df[m], errors='coerce').fillna(0)
            return df
        except:
            return pd.DataFrame()

    def save_to_supabase(df):
        with conn.session as session:
            session.execute(text("DELETE FROM financials"))
            session.commit()
        df.to_sql('financials', conn.engine, if_exists='append', index=False)

    def clean_currency(val):
        if pd.isna(val): return 0.0
        val_str = str(val).strip()
        if not val_str or val_str.lower() in ['nan', 'none', 'null']: return 0.0
        val_str = re.sub(r'[^\d\.\-\(\)]', '', val_str)
        if val_str.startswith('(') and val_str.endswith(')'):
            val_str = '-' + val_str[1:-1]
        try: return float(val_str)
        except: return 0.0

    def process_imported_file(uploaded_file):
        wb = openpyxl.load_workbook(uploaded_file, data_only=True)
        sheet = wb[wb.sheetnames[0]]
        
        color_list = []
        for row in sheet.iter_rows(min_row=4):
            final_marker = "無底色"
            # 掃描前 15 欄位獲取顏色
            for cell in row[1:16]: 
                fill = cell.fill
                if fill and hasattr(fill, 'start_color') and fill.start_color:
                    sc = fill.start_color
                    if sc.type == 'rgb' and sc.rgb and str(sc.rgb) not in ['00000000', '000000']:
                        final_marker = f"#{str(sc.rgb)[-6:].upper()}"
                        break
                    elif sc.type == 'theme' and sc.theme is not None:
                        if sc.theme in [5, 7, 9]: final_marker = f"THEME_PINK_{sc.theme}"
                        else: final_marker = f"THEME_GREY_{sc.theme}"
                        break
            color_list.append(final_marker)

        uploaded_file.seek(0)
        df = pd.read_excel(uploaded_file, skiprows=2)
        df.columns = [str(c).strip() for c in df.columns]
        df['顏色標記'] = color_list[:len(df)]
        df.rename(columns={df.columns[2]: '紀錄類型'}, inplace=True)
        df['專案說明'] = df['專案說明'].replace(r'^\s*$', pd.NA, regex=True).ffill()
        
        cat_col = next((c for c in df.columns if '營收分類' in str(c)), None)
        df['營收分類'] = df[cat_col].astype(str).str.replace('\n', ' ').str.strip() if cat_col else "其他"
        df['營收分類'] = df['營收分類'].replace(['nan', 'None', '', 'NaN'], np.nan).ffill().fillna("其他")
        
        months = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec']
        for m in months:
            df[m] = df[m].apply(clean_currency) if m in df.columns else 0.0

        target_cols = ['專案說明', '紀錄類型'] + months + ['營收分類', '顏色標記']
        return df.dropna(subset=['專案說明', '紀錄類型'])[target_cols]

    # ==========================================
    # 2. 側邊欄
    # ==========================================
    with st.sidebar:
        st.header("📂 匯入數據")
        f = st.file_uploader("選擇 Excel", type=["xlsx"])
        if f and st.button("🚀 匯入並執行規則對齊"):
            save_to_supabase(process_imported_file(f))
            st.success("匯入完成！已自動套用 A/B/C 組合規則。")
            st.rerun()

    # ==========================================
    # 3. 介面呈現
    # ==========================================
    data = load_data()
    if not data.empty:
        df = data.copy()
        months = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec']
        df['年度總計'] = df[months].sum(axis=1)

        # 核心：套用你的規則
        df['財務屬性'] = df.apply(lambda row: apply_user_rules(str(row['紀錄類型']), str(row['顏色標記'])), axis=1)

        tabs = st.tabs(["📈 1. 營收分類彙整總表", "🎴 2. 專案績效卡片", "🎨 3. 對應結果診斷"])

        with tabs[0]:
            st.subheader("📋 營收分類彙整總表")
            # 根據分類與屬性加總
            summary = df.groupby(['營收分類', '財務屬性'])['年度總計'].sum().unstack().fillna(0)
            
            # 確保欄位存在
            for c in ['🎯 原目標收入', '🔮 預估收入', '📉 原目標支出', '💸 預估支出']:
                if c not in summary.columns: summary[c] = 0
            
            # --- 套用你的呈現邏輯 ---
            summary['原毛利'] = summary['🎯 原目標收入'] - summary['📉 原目標支出']
            summary['預計毛利'] = summary['🔮 預估收入'] - summary['💸 預估支出']
            summary['差異'] = summary['🎯 原目標收入'] - summary['🔮 預估收入'] # 原目標收入 - 預估收入
            summary['毛利率'] = (summary['預計毛利'] / summary['🔮 預估收入']).replace([np.inf, -np.inf], 0).fillna(0)
            
            final_table = summary[['🎯 原目標收入', '🔮 預估收入', '原毛利', '預計毛利', '差異', '毛利率']]
            st.dataframe(final_table.reset_index().style.format({
                '🎯 原目標收入': '{:,.0f}', '🔮 預估收入': '{:,.0f}', 
                '原毛利': '{:,.0f}', '預計毛利': '{:,.0f}', 
                '差異': '{:,.0f}', '毛利率': '{:.2%}'
            }), use_container_width=True, hide_index=True)

        with tabs[1]:
            st.subheader("💡 專案績效卡片")
            projects = df['專案說明'].unique()
            cols = st.columns(2)
            for idx, proj in enumerate(projects):
                with cols[idx % 2]:
                    p_df = df[df['專案說明'] == proj]
                    target_rev = p_df[p_df['財務屬性'] == '🎯 原目標收入']['年度總計'].sum()
                    est_rev = p_df[p_df['財務屬性'] == '🔮 預估收入']['年度總計'].sum()
                    target_exp = p_df[p_df['財務屬性'] == '📉 原目標支出']['年度總計'].sum()
                    est_exp = p_df[p_df['財務屬性'] == '💸 預估支出']['年度總計'].sum()
                    
                    # 計算卡片指標
                    est_profit = est_rev - est_exp
                    achievement_rate = (est_rev / target_rev) if target_rev != 0 else 0
                    
                    with st.container(border=True):
                        st.markdown(f"#### {proj}")
                        m1, m2, m3 = st.columns(3)
                        m1.metric("原目標收入", f"${target_rev:,.0f}")
                        m2.metric("預估收入", f"${est_rev:,.0f}", f"{est_rev-target_rev:,.0f}")
                        m3.metric("目標達成率", f"{achievement_rate:.1%}")
                        
                        m4, m5 = st.columns(2)
                        m4.metric("預計毛利", f"${est_profit:,.0f}")
                        m5.metric("毛利率", f"{(est_profit/est_rev*100 if est_rev != 0 else 0):.1f}%")
                        
                        st.progress(min(max(achievement_rate, 0.0), 1.0))

        with tabs[2]:
            st.markdown("### 🔍 數據歸類自動檢查")
            st.write("如果數字不對，請檢查下方的「財務屬性」歸類是否正確。")
            st.dataframe(df[['專案說明', '紀錄類型', '顏色標記', '財務屬性', '年度總計']], use_container_width=True)
