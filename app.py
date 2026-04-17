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
    st.set_page_config(page_title="車聯網營收系統 v35", layout="wide")
    conn = st.connection("postgresql", type="sql")

    # ==========================================
    # 🎯 絕對商業規則 (組合 A, B, C, D)
    # ==========================================
    def apply_business_rules(record_type, color_marker):
        # 判定是否為粉紅色系
        is_pink = any(k in color_marker for k in ["#F2DCDB", "#FCE4D6", "THEME_PINK", "#FFC7CE", "#FAD0C9", "#F8CBAD"])
        # 判定是否為白色/灰色系 (包含無底色)
        is_white_grey = color_marker == "無底色" or any(k in color_marker for k in ["#D9D9D9", "#D8D8D8", "#FFFFFF", "THEME_GREY", "#E7E6E6", "#F2F2F2"])

        # 收入邏輯 (組合 A & B)
        if any(k in record_type for k in ["收入", "營收", "實績"]):
            if is_pink: return "🔮 預估收入"
            return "🎯 原目標收入"

        # 支出邏輯 (組合 C & D)
        if any(k in record_type for k in ["支出", "成本"]):
            if is_pink: return "💸 預估支出"
            return "📉 原目標支出"
            
        return "❌ 忽略不計"

    # ==========================================
    # 1. 核心數據處理引擎
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
        
        # 顏色掃描：確保不漏掉任何底色
        color_list = []
        for row in sheet.iter_rows(min_row=4):
            final_marker = "無底色"
            # 優先掃描 B(1) 跟 C(2) 兩欄的顏色，這通常是決定性質的關鍵
            for cell in row[1:3]: 
                fill = cell.fill
                if fill and hasattr(fill, 'start_color') and fill.start_color:
                    sc = fill.start_color
                    if sc.type == 'rgb' and sc.rgb and str(sc.rgb) not in ['00000000', '000000']:
                        final_marker = f"#{str(sc.rgb)[-6:].upper()}"
                        break
                    elif sc.type == 'theme' and sc.theme is not None:
                        # 4, 5, 7, 9 是常見的粉/橘系主題色
                        if sc.theme in [5, 7, 9, 4]: final_marker = f"THEME_PINK_{sc.theme}"
                        else: final_marker = f"THEME_GREY_{sc.theme}"
                        break
            color_list.append(final_marker)

        uploaded_file.seek(0)
        df = pd.read_excel(uploaded_file, skiprows=2)
        df.columns = [str(c).strip() for c in df.columns]
        
        # 附加顏色與紀錄類型
        df['顏色標記'] = color_list[:len(df)]
        df.rename(columns={df.columns[2]: '紀錄類型'}, inplace=True)
        
        # --- 核心修復：分類與名稱填補 ---
        # 先把空字串轉成 NaN 才能正確向下填補 (ffill)
        df['專案說明'] = df['專案說明'].replace(r'^\s*$', np.nan, regex=True).ffill()
        
        cat_col = next((c for c in df.columns if '營收分類' in str(c)), None)
        if cat_col:
            df['營收分類'] = df[cat_col].astype(str).str.replace('\n', ' ').str.strip()
            df['營收分類'] = df['營收分類'].replace(['nan', 'None', '', 'NaN'], np.nan).ffill().fillna("其他")
        else:
            df['營收分類'] = "其他"
        
        # 數字清洗
        months = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec']
        for m in months:
            if m in df.columns:
                df[m] = df[m].apply(clean_currency)
            else:
                df[m] = 0.0

        # 孤兒數字救援：如果 1-12 月是空的，就去抓小計
        fallback_cols = [c for c in df.columns if any(k in str(c) for k in ['小計', '總計', '合計', '實績'])]
        for idx, row in df.iterrows():
            if sum(row[months]) == 0:
                for f_col in fallback_cols:
                    if f_col in row:
                        val = clean_currency(row[f_col])
                        if val != 0:
                            df.at[idx, 'Jan'] = val
                            break
                            
        target_cols = ['專案說明', '紀錄類型'] + months + ['營收分類', '顏色標記']
        return df.dropna(subset=['紀錄類型'])[target_cols]

    # ==========================================
    # 2. UI 渲染區
    # ==========================================
    with st.sidebar:
        st.header("📂 匯入數據")
        f = st.file_uploader("選擇 Excel", type=["xlsx"])
        if f and st.button("🚀 匯入並自動計算報表"):
            save_to_supabase(process_imported_file(f))
            st.success("匯入成功！系統已完成所有分類與公式對齊。")
            st.rerun()

    data = load_data()
    if not data.empty:
        df = data.copy()
        months = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec']
        df['年度總計'] = df[months].sum(axis=1)

        # 套用商業規則
        df['財務屬性'] = df.apply(lambda row: apply_business_rules(str(row['紀錄類型']), str(row['顏色標記'])), axis=1)

        tabs = st.tabs(["📈 營收分類彙整總表", "🎴 專案績效卡片", "📝 原始診斷數據"])

        with tabs[0]:
            st.subheader("📋 營收分類彙整總表")
            summary = df.groupby(['營收分類', '財務屬性'])['年度總計'].sum().unstack().fillna(0)
            
            # 確保欄位齊全
            for c in ['🎯 原目標收入', '🔮 預估收入', '📉 原目標支出', '💸 預估支出']:
                if c not in summary.columns: summary[c] = 0
            
            # --- 絕對商業公式 ---
            summary['原毛利'] = summary['🎯 原目標收入'] - summary['📉 原目標支出']
            summary['預計毛利'] = summary['🔮 預估收入'] - summary['💸 預估支出']
            summary['差異'] = summary['🎯 原目標收入'] - summary['🔮 預估收入'] 
            summary['毛利率'] = (summary['預計毛利'] / summary['🔮 預估收入']).replace([np.inf, -np.inf], 0).fillna(0)
            
            display_df = summary[['🎯 原目標收入', '🔮 預估收入', '原毛利', '預計毛利', '差異', '毛利率']].reset_index()
            st.dataframe(display_df.style.format({
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
                    # 預先計算，防止錯誤
                    t_rev = p_df[p_df['財務屬性'] == '🎯 原目標收入']['年度總計'].sum()
                    e_rev = p_df[p_df['財務屬性'] == '🔮 預估收入']['年度總計'].sum()
                    t_exp = p_df[p_df['財務屬性'] == '📉 原目標支出']['年度總計'].sum()
                    e_exp = p_df[p_df['財務屬性'] == '💸 預估支出']['年度總計'].sum()
                    
                    est_profit = e_rev - e_exp
                    achievement = (e_rev / t_rev) if t_rev != 0 else 0
                    
                    with st.container(border=True):
                        st.markdown(f"#### {proj}")
                        m1, m2, m3 = st.columns(3)
                        m1.metric("原目標收入", f"${t_rev:,.0f}")
                        m2.metric("預估收入", f"${e_rev:,.0f}", f"{e_rev-t_rev:,.0f}")
                        m3.metric("目標達成率", f"{achievement:.1%}")
                        
                        m4, m5 = st.columns(2)
                        m4.metric("預計毛利", f"${est_profit:,.0f}")
                        m5.metric("預計毛利率", f"{(est_profit/e_rev*100 if e_rev != 0 else 0):.1f}%")
                        st.progress(min(max(achievement, 0.0), 1.0))

        with tabs[2]:
            st.markdown("### 🔍 歸類檢查診斷器")
            st.write("如果數值不對，請檢查『財務屬性』與『顏色標記』是否正確。")
            st.dataframe(df[['專案說明', '紀錄類型', '顏色標記', '財務屬性', '營收分類', '年度總計']], use_container_width=True)
