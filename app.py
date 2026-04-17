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
    st.set_page_config(page_title="車聯網營收系統 v38", layout="wide")
    conn = st.connection("postgresql", type="sql")

    # ==========================================
    # 🎯 核心商業邏輯 (組合 A, B, C, D)
    # ==========================================
    def apply_business_rules(record_type, color_marker):
        # 判定是否為粉紅色系
        is_pink = any(k in color_marker for k in ["#F2DCDB", "#FCE4D6", "THEME_PINK", "#FFC7CE", "#FAD0C9", "#F8CBAD"])
        
        # 收入邏輯
        if any(k in record_type for k in ["收入", "營收", "實績"]):
            return "🔮 預估收入" if is_pink else "🎯 原目標收入"
        # 支出邏輯
        if any(k in record_type for k in ["支出", "成本"]):
            return "💸 預估支出" if is_pink else "📉 原目標支出"
            
        return "❌ 忽略不計"

    # ==========================================
    # 1. 數據清洗引擎
    # ==========================================
    def clean_currency(val):
        if pd.isna(val) or val == "": return 0.0
        s = str(val).replace(',', '').strip()
        if s.startswith('(') and s.endswith(')'): s = '-' + s[1:-1]
        try: return float(re.sub(r'[^\d\.\-]', '', s))
        except: return 0.0

    def process_imported_file(uploaded_file):
        wb = openpyxl.load_workbook(uploaded_file, data_only=True)
        sheet = wb[wb.sheetnames[0]]
        
        # --- 顏色掃描 ---
        color_list = []
        for row in sheet.iter_rows(min_row=4):
            marker = "無底色"
            # 掃描 B, C 兩欄判定性質
            for cell in row[1:3]:
                if cell.fill and hasattr(cell.fill, 'start_color') and cell.fill.start_color:
                    sc = cell.fill.start_color
                    if sc.type == 'rgb' and sc.rgb and str(sc.rgb) not in ['00000000', '000000']:
                        marker = f"#{str(sc.rgb)[-6:].upper()}"
                        break
                    elif sc.type == 'theme' and sc.theme is not None:
                        marker = f"THEME_PINK" if sc.theme in [4,5,7,9] else "THEME_GREY"
                        break
            color_list.append(marker)

        uploaded_file.seek(0)
        # 讀取完整 DataFrame (不先過濾，確保 ffill 成功)
        df = pd.read_excel(uploaded_file, skiprows=2, header=0)
        df.columns = [str(c).strip() for c in df.columns]
        
        # --- 暴力座標定位 (不依賴欄位名稱) ---
        df['專案說明'] = df.iloc[:, 1].replace(r'^\s*$', np.nan, regex=True).ffill()
        df['紀錄類型'] = df.iloc[:, 2].astype(str).str.strip()
        
        # 處理營收分類 (Index 19)
        if len(df.columns) > 19:
            df['營收分類'] = df.iloc[:, 19].astype(str).str.replace('\n', ' ').str.strip()
            df['營收分類'] = df['營收分類'].replace(['nan', 'None', '', 'NaN'], np.nan).ffill().fillna("其他")
        else:
            df['營收分類'] = "其他"

        df['顏色標記'] = color_list[:len(df)]
        
        # 處理 1-12 月 (Index 3-14)
        months = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec']
        for i, m in enumerate(months):
            df[m] = df.iloc[:, 3+i].apply(clean_currency)

        # --- 🚀 救援引擎 (針對 24DCM 等只有小計有值的專案) ---
        # 收入小計在 Index 15, 支出小計在 Index 16
        for idx, row in df.iterrows():
            if sum(row[months]) == 0:
                is_income = any(k in str(row['紀錄類型']) for k in ["收入", "營收", "實績"])
                if is_income:
                    df.at[idx, 'Jan'] = clean_currency(row.iloc[15])
                else:
                    df.at[idx, 'Jan'] = clean_currency(row.iloc[16])
                            
        target_cols = ['專案說明', '紀錄類型'] + months + ['營收分類', '顏色標記']
        # 最後才過濾掉無效行，確保分類已填補
        return df.dropna(subset=['紀錄類型'])[target_cols]

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

    # ==========================================
    # 2. 頁面介面
    # ==========================================
    with st.sidebar:
        st.header("📂 數據導入")
        f = st.file_uploader("選擇您的 Excel", type=["xlsx"])
        if f and st.button("🚀 執行終極數據校正"):
            try:
                save_to_supabase(process_imported_file(f))
                st.success("匯入成功！分類標籤與小計數字已校正。")
                st.rerun()
            except Exception as e:
                st.error(f"解析發生錯誤: {e}")

    data = load_data()
    if not data.empty:
        df = data.copy()
        months = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec']
        df['年度總計'] = df[months].sum(axis=1)
        
        # 套用 A/B/C/D 規則
        df['財務屬性'] = df.apply(lambda row: apply_business_rules(str(row['紀錄類型']), str(row['顏色標記'])), axis=1)

        tabs = st.tabs(["📈 營收彙整總表", "🎴 專案卡片摘要", "🧪 數據診斷器"])

        with tabs[0]:
            st.subheader("📋 營收分類彙整總表")
            summary = df.groupby(['營收分類', '財務屬性'])['年度總計'].sum().unstack().fillna(0)
            for c in ['🎯 原目標收入', '🔮 預估收入', '📉 原目標支出', '💸 預估支出']:
                if c not in summary.columns: summary[c] = 0
            
            # --- 套用使用者要求的計算公式 ---
            summary['原毛利'] = summary['🎯 原目標收入'] - summary['📉 原目標支出']
            summary['預計毛利'] = summary['🔮 預估收入'] - summary['💸 預估支出']
            summary['差異'] = summary['🎯 原目標收入'] - summary['🔮 預估收入'] 
            summary['毛利率'] = (summary['預計毛利'] / summary['🔮 預估收入']).replace([np.inf, -np.inf], 0).fillna(0)
            
            st.dataframe(summary[['🎯 原目標收入', '🔮 預估收入', '原毛利', '預計毛利', '差異', '毛利率']].style.format({
                '🎯 原目標收入': '{:,.0f}', '🔮 預估收入': '{:,.0f}', 
                '原毛利': '{:,.0f}', '預計毛利': '{:,.0f}', 
                '差異': '{:,.0f}', '毛利率': '{:.2%}'
            }), use_container_width=True)

        with tabs[1]:
            st.subheader("💡 專案績效卡片")
            projects = df['專案說明'].unique()
            cols = st.columns(2)
            for idx, proj in enumerate(projects):
                with cols[idx % 2]:
                    p_df = df[df['專案說明'] == proj]
                    t_rev = p_df[p_df['財務屬性'] == '🎯 原目標收入']['年度總計'].sum()
                    e_rev = p_df[p_df['財務屬性'] == '🔮 預估收入']['年度總計'].sum()
                    e_exp = p_df[p_df['財務屬性'] == '💸 預估支出']['年度總計'].sum()
                    ach = (e_rev / t_rev) if t_rev != 0 else 0
                    
                    with st.container(border=True):
                        st.markdown(f"#### {proj}")
                        m1, m2, m3 = st.columns(3)
                        m1.metric("原目標收入", f"${t_rev:,.0f}")
                        m2.metric("預估收入", f"${e_rev:,.0f}", f"{e_rev-t_rev:,.0f}")
                        m3.metric("達成率", f"{ach:.1%}")
                        st.progress(min(max(ach, 0.0), 1.0))

        with tabs[2]:
            st.write("這是系統當前抓到的分類與標記，若數字不對請檢查此表（特別是『營收分類』與『財務屬性』）：")
            st.dataframe(df[['專案說明', '紀錄類型', '營收分類', '顏色標記', '財務屬性', '年度總計']], use_container_width=True)
