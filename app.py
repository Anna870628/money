import streamlit as st
import pandas as pd
import numpy as np
from sqlalchemy import text
import openpyxl
from openpyxl.styles.colors import COLOR_INDEX
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
    st.set_page_config(page_title="車聯網營收系統 v30", layout="wide")
    conn = st.connection("postgresql", type="sql")

    # ==========================================
    # 1. 核心邏輯區
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
        try:
            return float(val_str)
        except:
            return 0.0

    def get_standard_hex(color_obj):
        """將 openpyxl 的顏色物件轉換為網頁可用的 HEX 色碼"""
        if color_obj.type == 'rgb' and color_obj.rgb:
            rgb_val = str(color_obj.rgb)
            if rgb_val not in ['00000000', '000000']:
                return f"#{rgb_val[-6:].upper()}"
        elif color_obj.type == 'indexed' and color_obj.indexed is not None:
            try:
                # 查找 Excel 標準 64 色索引
                idx_rgb = COLOR_INDEX[color_obj.indexed]
                return f"#{idx_rgb[-6:].upper()}"
            except:
                pass
        elif color_obj.type == 'theme' and color_obj.theme is not None:
            # 標準 Office 主題色盤預估 (根據常見編號)
            theme_defaults = {
                0: "#FFFFFF", 1: "#000000", 2: "#E7E6E6", 3: "#44546A",
                4: "#5B9BD5", 5: "#ED7D31", 6: "#A5A5A5", 7: "#FFC000",
                8: "#4472C4", 9: "#70AD47"
            }
            # 如果是粉紅色系 (通常是 Accent 2 或特別色)
            if color_obj.theme in [5, 7, 9]: return f"THEME_PINK_{color_obj.theme}"
            return f"THEME_{color_obj.theme}"
        return "無底色"

    def process_imported_file(uploaded_file):
        wb = openpyxl.load_workbook(uploaded_file, data_only=True)
        sheet_names = wb.sheetnames
        target_s = sheet_names[0]
        for s in sheet_names:
            if any(k in s for k in ['營收', '預估', '收支']):
                target_s = s
                break
        sheet = wb[target_s]
        
        color_list = []
        for row in sheet.iter_rows(min_row=4):
            final_marker = "無底色"
            # 掃描前 15 欄，尋找最明顯的底色
            for cell in row[1:16]: 
                fill = cell.fill
                if fill and hasattr(fill, 'start_color') and fill.start_color:
                    marker = get_standard_hex(fill.start_color)
                    if marker != "無底色":
                        final_marker = marker
                        break
            color_list.append(final_marker)

        uploaded_file.seek(0)
        df = pd.read_excel(uploaded_file, sheet_name=target_s, skiprows=2)
        df.columns = [str(c).strip() for c in df.columns]
        
        df['顏色標記'] = color_list[:len(df)]
        df.rename(columns={df.columns[2]: '紀錄類型'}, inplace=True)
        df['專案說明'] = df['專案說明'].replace(r'^\s*$', pd.NA, regex=True).ffill()
        df['紀錄類型'] = df['紀錄類型'].astype(str).str.strip()
        
        cat_col_name = next((c for c in df.columns if '營收分類' in str(c)), None)
        if cat_col_name:
            df['營收分類'] = df[cat_col_name].astype(str).str.replace('\n', ' ').str.strip()
            df['營收分類'] = df['營收分類'].replace(['nan', 'None', '', 'NaN'], np.nan).ffill().fillna("其他")
        else:
            df['營收分類'] = "其他"
        
        months = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec']
        for m in months:
            if m in df.columns:
                df[m] = df[m].apply(clean_currency)
            else:
                df[m] = 0.0

        fallback_cols = [c for c in df.columns if any(k in str(c) for k in ['小計', '總計', '合計', '實績'])]
        for idx, row in df.iterrows():
            if sum(row[months]) == 0:
                for f_col in fallback_cols:
                    val = clean_currency(row[f_col])
                    if val != 0:
                        df.at[idx, 'Jan'] = val
                        break
                            
        target_cols = ['專案說明', '紀錄類型'] + months + ['營收分類', '顏色標記']
        return df.dropna(subset=['專案說明', '紀錄類型'])[target_cols]

    # ==========================================
    # 2. 側邊欄與資料載入
    # ==========================================
    with st.sidebar:
        st.header("📂 匯入數據")
        f = st.file_uploader("選擇 Excel", type=["xlsx"])
        if f and st.button("🚀 啟動全方位色彩解析"):
            save_to_supabase(process_imported_file(f))
            st.success("解析成功！色塊已重新打撈。")
            st.rerun()

    data = load_data()

    if data.empty:
        st.warning("⚠️ 資料庫目前無數據。")
    else:
        df = data.copy()
        months = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec']
        df['年度總計'] = df[months].sum(axis=1)

        tabs = st.tabs(["🎨 1. 視覺化對應規則", "📈 2. 營收彙整報表", "🎴 3. 專案績效卡片"])

        with tabs[0]:
            st.markdown("### 🧩 財務歸類設定")
            unique_combos = df[['紀錄類型', '顏色標記']].drop_duplicates().values.tolist()
            mapping_dict = {}
            cols = st.columns(2)
            
            for idx, (record_type, color_marker) in enumerate(unique_combos):
                # 視覺化色塊渲染
                hex_render = "#FFFFFF"
                if color_marker.startswith("#"):
                    hex_render = color_marker
                elif "PINK" in color_marker:
                    hex_render = "#F2DCDB" # 系統識別出的粉紅色系
                elif "THEME" in color_marker:
                    hex_render = "#D9D9D9" # 系統識別出的灰色系

                with cols[idx % 2]:
                    with st.container(border=True):
                        st.markdown(f'''
                            <div style="display:flex; align-items:center; margin-bottom:10px;">
                                <div style="width:32px; height:32px; background-color:{hex_render}; border:2px solid #333; border-radius:4px; margin-right:12px;"></div>
                                <b>文字：{record_type}</b>
                            </div>
                        ''', unsafe_allow_html=True)
                        
                        mapping_dict[(record_type, color_marker)] = st.selectbox(
                            "財務定義：", 
                            ['❌ 忽略不計', '🎯 原目標收入', '🔮 預估收入', '📉 實際支出', '💸 預估支出'],
                            key=f"map_{idx}",
                            index=2 if '預估' in record_type or hex_render == "#F2DCDB" else 1 if '收入' in record_type else 3
                        )

            df['財務屬性'] = df.apply(lambda row: mapping_dict.get((str(row['紀錄類型']), str(row['顏色標記'])), '❌ 忽略不計'), axis=1)

        with tabs[1]:
            st.subheader("📋 營收分類彙整總表")
            unique_cats = df['營收分類'].unique().tolist()
            summary = pd.DataFrame(index=unique_cats)
            summary['原目標收入'] = df[df['財務屬性'] == '🎯 原目標收入'].groupby('營收分類')['年度總計'].sum()
            summary['預估收入'] = df[df['財務屬性'] == '🔮 預估收入'].groupby('營營分類' if False else '營收分類')['年度總計'].sum()
            summary['實際支出'] = df[df['財務屬性'] == '📉 實際支出'].groupby('營收分類')['年度總計'].sum()
            summary['預估支出'] = df[df['財務屬性'] == '💸 預估支出'].groupby('營收分類')['年度總計'].sum()
            summary = summary.fillna(0)
            summary['預估毛利'] = summary['預估收入'] - summary['預估支出']
            summary['差異'] = summary['預估收入'] - summary['原目標收入']
            summary['毛利率'] = (summary['預估毛利'] / summary['預估收入']).replace([np.inf, -np.inf], 0).fillna(0)
            st.dataframe(summary.reset_index().style.format({'毛利率': '{:.2%}', '預估收入': '{:,.0f}'}), use_container_width=True, hide_index=True)

        with tabs[2]:
            st.subheader("💡 專案績效卡片")
            projects = df['專案說明'].unique()
            cols = st.columns(2)
            for idx, proj in enumerate(projects):
                with cols[idx % 2]:
                    p_df = df[df['專案說明'] == proj]
                    t_rev = p_df[p_df['財務屬性'] == '🎯 原目標收入']['年度總計'].sum()
                    e_rev = p_df[p_df['財務屬性'] == '🔮 預估收入']['年度總計'].sum()
                    with st.container(border=True):
                        st.markdown(f"#### {proj}")
                        m1, m2 = st.columns(2)
                        m1.metric("目標", f"${t_rev:,.0f}")
                        m2.metric("預估", f"${e_rev:,.0f}", f"{e_rev-t_rev:,.0f}")
