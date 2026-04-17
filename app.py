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
    st.set_page_config(page_title="車聯網營收系統 v37", layout="wide")
    conn = st.connection("postgresql", type="sql")

    # ==========================================
    # 🎯 使用者定義的商業規則 (A/B/C/D)
    # ==========================================
    def apply_strict_rules(record_type, color_marker):
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
    # 1. 數據清洗核心
    # ==========================================
    def clean_currency(val):
        if pd.isna(val) or val == "": return 0.0
        s = str(val).replace(',', '').strip()
        # 處理會計格式 (100) -> -100
        if s.startswith('(') and s.endswith(')'):
            s = '-' + s[1:-1]
        try:
            return float(re.sub(r'[^\d\.\-]', '', s))
        except:
            return 0.0

    def process_imported_file(uploaded_file):
        wb = openpyxl.load_workbook(uploaded_file, data_only=True)
        sheet = wb[wb.sheetnames[0]]
        
        # --- 顏色掃描 ---
        color_list = []
        # 從第 4 行開始讀取數據
        for row in sheet.iter_rows(min_row=4):
            marker = "無底色"
            # 檢查專案說明(B)與類型(C)兩欄的底色
            for cell in row[1:3]:
                if cell.fill and hasattr(cell.fill, 'start_color') and cell.fill.start_color:
                    sc = cell.fill.start_color
                    if sc.type == 'rgb' and sc.rgb and str(sc.rgb) not in ['00000000', '000000']:
                        marker = f"#{str(sc.rgb)[-6:].upper()}"
                        break
                    elif sc.type == 'theme' and sc.theme is not None:
                        # 4, 5, 7, 9 通常是粉色系
                        marker = f"THEME_PINK_{sc.theme}" if sc.theme in [4,5,7,9] else f"THEME_GREY_{sc.theme}"
                        break
            color_list.append(marker)

        uploaded_file.seek(0)
        # 跳過前 2 行，從第 3 行標題列開始讀
        df = pd.read_excel(uploaded_file, skiprows=2)
        
        # 修正欄位名稱 (你的檔案第 3 個欄位通常沒名字)
        df.columns = [str(c).strip() for c in df.columns]
        df.rename(columns={df.columns[1]: '專案說明', df.columns[2]: '紀錄類型'}, inplace=True)
        
        # 填補專案說明 (ffill)
        df['專案說明'] = df['專案說明'].replace(r'^\s*$', np.nan, regex=True).ffill()
        
        # 營收分類處理 (包含 24DCM 的填補)
        cat_col = next((c for c in df.columns if '營收分類' in c), None)
        if cat_col:
            df['營收分類'] = df[cat_col].astype(str).str.replace('\n', ' ').str.strip()
            df['營收分類'] = df['營收分類'].replace(['nan', 'None', '', 'NaN'], np.nan).ffill().fillna("其他")
        else:
            df['營收分類'] = "其他"

        df['顏色標記'] = color_list[:len(df)]
        
        # 處理 1-12 月數據
        months = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec']
        for m in months:
            m_col = next((c for c in df.columns if m in c), None)
            df[m] = df[m_col].apply(clean_currency) if m_col else 0.0

        # --- 🚀 24DCM 孤兒數字救援引擎 ---
        # 如果 1-12 月總和為 0，去抓後面「收入小計」或「支出小計」
        fallback_income = next((c for c in df.columns if '收入小計' in c), None)
        fallback_expense = next((c for c in df.columns if '支出小計' in c), None)

        for idx, row in df.iterrows():
            if sum(row[months]) == 0:
                is_inc = any(k in str(row['紀錄類型']) for k in ["收入", "營收", "實績"])
                if is_inc and fallback_income:
                    df.at[idx, 'Jan'] = clean_currency(row[fallback_income])
                elif not is_inc and fallback_expense:
                    df.at[idx, 'Jan'] = clean_currency(row[fallback_expense])
                            
        target_cols = ['專案說明', '紀錄類型'] + months + ['營收分類', '顏色標記']
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
        if f and st.button("🚀 執行數據分析"):
            with st.spinner("正在解析您的商業邏輯與金額..."):
                try:
                    new_df = process_imported_file(f)
                    save_to_supabase(new_df)
                    st.success("匯入成功！數據已根據底色與小計校正。")
                    st.rerun()
                except Exception as e:
                    st.error(f"解析發生錯誤: {e}")

    data = load_data()
    if not data.empty:
        df = data.copy()
        months = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec']
        df['年度總計'] = df[months].sum(axis=1)
        
        # 套用規則
        df['財務屬性'] = df.apply(lambda row: apply_strict_rules(str(row['紀錄類型']), str(row['顏色標記'])), axis=1)

        tabs = st.tabs(["📈 營收彙整報表", "🎴 專案卡片摘要", "📝 診斷診斷器"])

        with tabs[0]:
            st.subheader("📋 營收分類彙整總表")
            # 建立透視加總
            summary = df.groupby(['營收分類', '財務屬性'])['年度總計'].sum().unstack().fillna(0)
            
            # 補齊可能缺失的欄位
            for c in ['🎯 原目標收入', '🔮 預估收入', '📉 原目標支出', '💸 預估支出']:
                if c not in summary.columns: summary[c] = 0
            
            # --- 絕對商業公式 ---
            summary['原毛利'] = summary['🎯 原目標收入'] - summary['📉 原目標支出']
            summary['預計毛利'] = summary['🔮 預估收入'] - summary['💸 預估支出']
            summary['差異'] = summary['🎯 原目標收入'] - summary['🔮 預估收入'] # 目標-預估
            summary['毛利率'] = (summary['預計毛利'] / summary['🔮 預估收入']).replace([np.inf, -np.inf], 0).fillna(0)
            
            final_view = summary[['🎯 原目標收入', '🔮 預估收入', '原毛利', '預計毛利', '差異', '毛利率']].reset_index()
            st.dataframe(final_view.style.format({
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
                    t_rev = p_df[p_df['財務屬性'] == '🎯 原目標收入']['年度總計'].sum()
                    e_rev = p_df[p_df['財務屬性'] == '🔮 預估收入']['年度總計'].sum()
                    e_exp = p_df[p_df['財務屬性'] == '💸 預估支出']['年度總計'].sum()
                    
                    ach = (e_rev / t_rev) if t_rev != 0 else 0
                    
                    with st.container(border=True):
                        st.markdown(f"#### {proj}")
                        c1, c2, c3 = st.columns(3)
                        c1.metric("原目標收入", f"${t_rev:,.0f}")
                        c2.metric("預估收入", f"${e_rev:,.0f}", f"{e_rev-t_rev:,.0f}")
                        c3.metric("達成率", f"{ach:.1%}")
                        st.progress(min(max(ach, 0.0), 1.0))

        with tabs[2]:
            st.write("這是系統目前抓到的分類與底色歸類，若數字不對請檢查此表：")
            st.dataframe(df[['專案說明', '紀錄類型', '顏色標記', '財務屬性', '營收分類', '年度總計']], use_container_width=True)
