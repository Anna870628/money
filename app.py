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
    st.set_page_config(page_title="車聯網營收系統 v36", layout="wide")
    conn = st.connection("postgresql", type="sql")

    # ==========================================
    # 🎯 絕對商業規則 (組合 A, B, C, D)
    # ==========================================
    def apply_business_rules(record_type, color_marker):
        is_pink = any(k in color_marker for k in ["#F2DCDB", "#FCE4D6", "THEME_PINK", "#FFC7CE", "#FAD0C9", "#F8CBAD"])
        is_white_grey = color_marker == "無底色" or any(k in color_marker for k in ["#D9D9D9", "#D8D8D8", "#FFFFFF", "THEME_GREY"])

        # 收入類
        if any(k in record_type for k in ["收入", "營收", "實績"]):
            return "🔮 預估收入" if is_pink else "🎯 原目標收入"
        # 支出類
        if any(k in record_type for k in ["支出", "成本"]):
            return "💸 預估支出" if is_pink else "📉 原目標支出"
        return "❌ 忽略不計"

    # ==========================================
    # 1. 核心清洗引擎 (徹底修復 KeyError)
    # ==========================================
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
        
        # 顏色掃描
        color_list = []
        for row in sheet.iter_rows(min_row=4):
            final_marker = "無底色"
            for cell in row[1:3]: 
                fill = cell.fill
                if fill and hasattr(fill, 'start_color') and fill.start_color:
                    sc = fill.start_color
                    if sc.type == 'rgb' and sc.rgb and str(sc.rgb) not in ['00000000', '000000']:
                        final_marker = f"#{str(sc.rgb)[-6:].upper()}"
                        break
                    elif sc.type == 'theme' and sc.theme is not None:
                        if sc.theme in [5, 7, 9, 4]: final_marker = f"THEME_PINK_{sc.theme}"
                        else: final_marker = f"THEME_GREY_{sc.theme}"
                        break
            color_list.append(final_marker)

        uploaded_file.seek(0)
        df = pd.read_excel(uploaded_file, skiprows=2)
        
        # --- 暴力修復欄位名稱 ---
        df.columns = [str(c).strip() for c in df.columns]
        
        # 1. 定位「專案說明」
        proj_col = next((c for c in df.columns if '專案說明' in c), df.columns[1])
        df.rename(columns={proj_col: '專案說明'}, inplace=True)
        
        # 2. 定位「紀錄類型」 (通常是第三欄，且標題可能是空的)
        type_col = df.columns[2]
        df.rename(columns={type_col: '紀錄類型'}, inplace=True)
        
        # 3. 定位「營收分類」
        cat_col = next((c for c in df.columns if '營收分類' in c), None)
        if cat_col:
            df.rename(columns={cat_col: '營收分類'}, inplace=True)
            df['營收分類'] = df['營收分類'].astype(str).str.replace('\n', ' ').str.strip()
            df['營收分類'] = df['營收分類'].replace(['nan', 'None', '', 'NaN'], np.nan).ffill().fillna("其他")
        else:
            df['營收分類'] = "其他"

        df['顏色標記'] = color_list[:len(df)]
        df['專案說明'] = df['專案說明'].replace(r'^\s*$', np.nan, regex=True).ffill()
        df['紀錄類型'] = df['紀錄類型'].astype(str).str.strip()
        
        # 數字處理
        months = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec']
        for m in months:
            m_col = next((c for c in df.columns if m in c), None)
            if m_col:
                df[m] = df[m_col].apply(clean_currency)
            else:
                df[m] = 0.0

        # 孤兒數字救援
        fallback_cols = [c for c in df.columns if any(k in c for k in ['小計', '總計', '合計', '實績'])]
        for idx, row in df.iterrows():
            if sum(row[months]) == 0:
                for f_col in fallback_cols:
                    val = clean_currency(row[f_col])
                    if val != 0:
                        df.at[idx, 'Jan'] = val
                        break
                            
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
    # 2. UI 渲染
    # ==========================================
    with st.sidebar:
        st.header("📂 匯入數據")
        f = st.file_uploader("選擇 Excel", type=["xlsx"])
        if f and st.button("🚀 匯入並自動計算報表"):
            try:
                save_to_supabase(process_imported_file(f))
                st.success("匯入成功！")
                st.rerun()
            except Exception as e:
                st.error(f"解析失敗: {e}")

    data = load_data()
    if not data.empty:
        df = data.copy()
        months = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec']
        df['年度總計'] = df[months].sum(axis=1)
        df['財務屬性'] = df.apply(lambda row: apply_business_rules(str(row['紀錄類型']), str(row['顏色標記'])), axis=1)

        tabs = st.tabs(["📈 營收分類彙整總表", "🎴 專案績效卡片", "📝 診斷數據"])

        with tabs[0]:
            st.subheader("📋 營收分類彙整總表")
            summary = df.groupby(['營收分類', '財務屬性'])['年度總計'].sum().unstack().fillna(0)
            for c in ['🎯 原目標收入', '🔮 預估收入', '📉 原目標支出', '💸 預估支出']:
                if c not in summary.columns: summary[c] = 0
            
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
                        m1.metric("原目標", f"${t_rev:,.0f}")
                        m2.metric("預估收入", f"${e_rev:,.0f}", f"{e_rev-t_rev:,.0f}")
                        m3.metric("達成率", f"{ach:.1%}")
                        st.progress(min(max(ach, 0.0), 1.0))
