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
    st.set_page_config(page_title="車聯網營收戰情系統 v42", layout="wide")
    conn = st.connection("postgresql", type="sql")

    # ==========================================
    # 🎯 核心自動對齊規則
    # ==========================================
    def apply_business_rules(record_type, color_marker):
        is_pink = any(k in str(color_marker) for k in ["#F2DCDB", "#FCE4D6", "THEME_PINK", "#FFC7CE", "#FAD0C9", "#F8CBAD"])
        if any(k in str(record_type) for k in ["收入", "營收", "實績"]):
            return "🔮 預估收入" if is_pink else "🎯 原目標收入"
        if any(k in str(record_type) for k in ["支出", "成本"]):
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
            session.execute(text("DROP TABLE IF EXISTS financials"))
            df.to_sql('financials', conn.engine, if_exists='replace', index=False)
            session.commit()

    def process_imported_file(uploaded_file):
        wb = openpyxl.load_workbook(uploaded_file, data_only=True)
        sheet = wb[wb.sheetnames[0]]
        color_list = []
        for row in sheet.iter_rows(min_row=4):
            marker = "無底色"
            for cell in row[1:3]: # 檢查 B, C 欄底色
                if cell.fill and hasattr(cell.fill, 'start_color') and cell.fill.start_color:
                    sc = cell.fill.start_color
                    if sc.type == 'rgb' and sc.rgb and str(sc.rgb) not in ['00000000', '000000']:
                        marker = f"#{str(sc.rgb)[-6:].upper()}"
                        break
                    elif sc.type == 'theme' and sc.theme is not None:
                        marker = "THEME_PINK" if sc.theme in [4,5,7,9] else "THEME_GREY"
                        break
            color_list.append(marker)

        uploaded_file.seek(0)
        raw = pd.read_excel(uploaded_file, skiprows=2)
        raw.columns = [str(c).strip() for c in raw.columns]
        
        def find_idx(keywords, default):
            for i, col in enumerate(raw.columns):
                if any(k in col for k in keywords): return i
            return default

        proj_idx = find_idx(["專案"], 1)
        type_idx = find_idx(["類型", "Unnamed: 2"], 2)
        cat_idx  = find_idx(["分類"], -1)
        jan_idx  = find_idx(["Jan"], 3)

        df = pd.DataFrame()
        df['專案說明'] = raw.iloc[:, proj_idx].replace(r'^\s*$', np.nan, regex=True).ffill()
        df['紀錄類型'] = raw.iloc[:, type_idx].astype(str).str.strip()
        
        if cat_idx != -1 and cat_idx < len(raw.columns):
            df['營收分類'] = raw.iloc[:, cat_idx].astype(str).str.replace('\n', ' ').str.strip()
            df['營收分類'] = df['營收分類'].replace(['nan', 'None', '', 'NaN'], np.nan).ffill().fillna("其他")
        else:
            df['營收分類'] = "其他"

        df['顏色標記'] = color_list[:len(df)]
        months = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec']
        for i, m in enumerate(months):
            c_idx = jan_idx + i
            df[m] = raw.iloc[:, c_idx].apply(clean_currency) if c_idx < len(raw.columns) else 0.0

        f_inc_idx = find_idx(["收入小計"], -1)
        f_exp_idx = find_idx(["支出小計"], -1)
        for idx, row in df.iterrows():
            if sum(row[months]) == 0:
                is_inc = any(k in str(row['紀錄類型']) for k in ["收入", "營收", "實績"])
                if is_inc and f_inc_idx != -1: df.at[idx, 'Jan'] = clean_currency(raw.iloc[idx, f_inc_idx])
                elif not is_inc and f_exp_idx != -1: df.at[idx, 'Jan'] = clean_currency(raw.iloc[idx, f_exp_idx])
                            
        return df.dropna(subset=['紀錄類型'])[['專案說明', '紀錄類型'] + months + ['營收分類', '顏色標記']]

    # ==========================================
    # 2. 介面呈現 (固定顯示分頁)
    # ==========================================
    st.title("📊 專案營收戰情系統")
    tab_cards, tab_summary, tab_import = st.tabs(["🎴 專案戰情卡片", "📈 分類匯總報表", "📥 數據匯入與編輯"])

    data = load_data()

    # --- 分頁 1 & 2: 只有有資料時才顯示內容 ---
    with tab_cards:
        if data.empty:
            st.info("💡 雲端資料庫目前是空的，請先至『📥 數據匯入與編輯』分頁上傳 Excel。")
        else:
            df = data.copy()
            months = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec']
            df['年度總計'] = df[months].sum(axis=1)
            df['財務屬性'] = df.apply(lambda row: apply_business_rules(row['紀錄類型'], row['顏色標記']), axis=1)
            
            all_cats = ["全部分類"] + list(df['營收分類'].unique())
            sel_cat = st.selectbox("篩選營收分類", all_cats)
            display_data = df if sel_cat == "全部分類" else df[df['營收分類'] == sel_cat]
            
            cols = st.columns(2)
            for idx, proj in enumerate(display_data['專案說明'].unique()):
                with cols[idx % 2]:
                    p_df = display_data[display_data['專案說明'] == proj]
                    t_rev = p_df[p_df['財務屬性'] == '🎯 原目標收入']['年度總計'].sum()
                    e_rev = p_df[p_df['財務屬性'] == '🔮 預估收入']['年度總計'].sum()
                    e_exp = p_df[p_df['財務屬性'] == '💸 預估支出']['年度總計'].sum()
                    margin = ((e_rev - e_exp) / e_rev) if e_rev != 0 else 0
                    reach = (e_rev / t_rev) if t_rev != 0 else 0
                    
                    with st.container(border=True):
                        st.markdown(f"### {proj}")
                        m1, m2, m3 = st.columns(3)
                        m1.metric("目標營收", f"{t_rev:,.0f}")
                        m2.metric("預估營收", f"{e_rev:,.0f}", f"{e_rev-t_rev:,.0f}")
                        m3.metric("預估毛利率", f"{margin:.1%}")
                        st.write(f"**目標達成率: {reach:.1%}**")
                        st.progress(min(max(reach, 0.0), 1.0))

    with tab_summary:
        if data.empty:
            st.info("💡 雲端資料庫目前是空的，請先至『📥 數據匯入與編輯』分頁上傳 Excel。")
        else:
            df = data.copy()
            months = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec']
            df['年度總計'] = df[months].sum(axis=1)
            df['財務屬性'] = df.apply(lambda row: apply_business_rules(row['紀錄類型'], row['顏色標記']), axis=1)
            
            summary = df.groupby(['營收分類', '財務屬性'])['年度總計'].sum().unstack().fillna(0)
            for c in ['🎯 原目標收入', '🔮 預估收入', '📉 原目標支出', '💸 預估支出']:
                if c not in summary.columns: summary[c] = 0
            
            summary['原毛利'] = summary['🎯 原目標收入'] - summary['📉 原目標支出']
            summary['預計毛利'] = summary['🔮 預估收入'] - summary['💸 預估支出']
            summary['差異'] = summary['🎯 原目標收入'] - summary['🔮 預估收入']
            summary['毛利率'] = (summary['預計毛利'] / summary['🔮 預估收入']).replace([np.inf, -np.inf], 0).fillna(0)
            st.dataframe(summary[['🎯 原目標收入', '🔮 預估收入', '原毛利', '預計毛利', '差異', '毛利率']].style.format({'毛利率': '{:.2%}', '預估收入': '{:,.0f}'}), use_container_width=True)

    # --- 分頁 3: 永遠顯示匯入功能 ---
    with tab_import:
        st.subheader("📥 匯入新資料")
        st.markdown("---")
        c1, c2 = st.columns([1, 2])
        with c1:
            f = st.file_uploader("選擇 Excel", type=["xlsx"], key="file_upload_widget")
            if f and st.button("🚀 執行自動對齊匯入"):
                try:
                    processed_df = process_imported_file(f)
                    save_to_supabase(processed_df)
                    st.success("匯入成功！數據已存入雲端資料庫。")
                    st.rerun()
                except Exception as e:
                    st.error(f"解析發生錯誤: {e}")
        with c2:
            if not data.empty:
                st.write("📝 **雲端數據即時微調**")
                edited = st.data_editor(data, num_rows="dynamic", use_container_width=True)
                if st.button("💾 儲存微調變更"):
                    save_to_supabase(edited)
                    st.rerun()
