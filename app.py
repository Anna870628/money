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
    st.set_page_config(page_title="車聯網營收系統 v29", layout="wide")
    conn = st.connection("postgresql", type="sql")

    # ==========================================
    # 1. 核心邏輯區 (補回遺漏的功能)
    # ==========================================
    def load_data():
        """從資料庫讀取資料"""
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
        """將清洗後的資料存入資料庫"""
        with conn.session as session:
            session.execute(text("DELETE FROM financials"))
            session.commit()
        df.to_sql('financials', conn.engine, if_exists='append', index=False)

    def clean_currency(val):
        """強制洗淨所有 $、,、() 等會計符號"""
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

    def process_imported_file(uploaded_file):
        """X 光全列掃描與格式清洗"""
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
            final_hex = "無底色"
            # 掃描前 15 欄，找尋非白色背景
            for cell in row[1:16]: 
                fill = cell.fill
                if fill and hasattr(fill, 'start_color') and fill.start_color:
                    sc = fill.start_color
                    # A. 標準 RGB 色碼
                    if sc.type == 'rgb' and sc.rgb and str(sc.rgb) not in ['00000000', '000000']:
                        final_hex = f"#{str(sc.rgb)[-6:].upper()}"
                        break
                    # B. 主題色 (Theme)
                    elif sc.type == 'theme' and sc.theme is not None:
                        # 將主題色編號記錄下來，UI 會根據此編號顯示不同色塊
                        final_hex = f"THEME_{sc.theme}"
                        break
            color_list.append(final_hex)

        uploaded_file.seek(0)
        df = pd.read_excel(uploaded_file, sheet_name=target_s, skiprows=2)
        df.columns = [str(c).strip() for c in df.columns]
        
        df['顏色標記'] = color_list[:len(df)]
        df.rename(columns={df.columns[2]: '紀錄類型'}, inplace=True)
        df['專案說明'] = df['專案說明'].replace(r'^\s*$', pd.NA, regex=True).ffill()
        df['紀錄類型'] = df['紀錄類型'].astype(str).str.strip()
        
        # 分類清洗 (解決換行或空值問題)
        cat_col_name = next((c for c in df.columns if '營收分類' in str(c)), None)
        if cat_col_name:
            df['營收分類'] = df[cat_col_name].astype(str).str.replace('\n', ' ').str.strip()
            df['營收分類'] = df['營營分類' if False else '營收分類'].replace(['nan', 'None', '', 'NaN'], np.nan)
            df['營收分類'] = df['營收分類'].ffill().fillna("其他")
        else:
            df['營收分類'] = "其他"
        
        # 1-12 月數字強制淨化
        months = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec']
        for m in months:
            if m in df.columns:
                df[m] = df[m].apply(clean_currency)
            else:
                df[m] = 0.0

        # 孤兒數字救援 (防止 1-12 月沒填卻有總計的情況)
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
    # 2. 側邊欄：匯入
    # ==========================================
    with st.sidebar:
        st.header("📂 匯入數據")
        f = st.file_uploader("選擇 Excel", type=["xlsx"])
        if f and st.button("🚀 啟動全方位掃描並上傳"):
            with st.spinner("讀取顏色與數字淨化中..."):
                try:
                    new_df = process_imported_file(f)
                    save_to_supabase(new_df)
                    st.success("匯入成功！請前往對應分頁。")
                    st.rerun()
                except Exception as e:
                    st.error(f"解析失敗: {e}")
        
        st.divider()
        if st.button("⚠️ 清空資料庫"):
            with conn.session as s:
                s.execute(text("DELETE FROM financials"))
                s.commit()
            st.rerun()

    # ==========================================
    # 3. 介面呈現
    # ==========================================
    data = load_data()

    if data.empty:
        st.warning("⚠️ 資料庫為空，請由側邊欄匯入。")
    else:
        df = data.copy()
        months = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec']
        df['年度總計'] = df[months].sum(axis=1)

        tabs = st.tabs(["🎨 1. 視覺化資料對應", "📈 2. 營收分類總表", "🎴 3. 專案卡片摘要"])

        # --- TAB 1: 視覺化對應 ---
        with tabs[0]:
            st.markdown("### 🧩 將 Excel 格式對應到正確分類")
            st.info("請看下方的色塊與文字，直接指定它們在財務報表中的角色。")
            
            unique_combos = df[['紀錄類型', '顏色標記']].drop_duplicates().values.tolist()
            mapping_dict = {}
            cols = st.columns(2)
            
            for idx, (record_type, color_marker) in enumerate(unique_combos):
                # UI 色塊渲染
                hex_color = "#FFFFFF" # 預設白色
                if color_marker.startswith("#"):
                    hex_color = color_marker
                elif "THEME_4" in color_marker or "THEME_5" in color_marker or "THEME_7" in color_marker:
                    hex_color = "#F2DCDB" # 強制給予粉紅色視覺
                elif "THEME_" in color_marker:
                    hex_color = "#D9D9D9" # 預設灰色視覺

                with cols[idx % 2]:
                    with st.container(border=True):
                        st.markdown(f'''
                            <div style="display:flex; align-items:center; margin-bottom:10px;">
                                <div style="width:30px; height:30px; background-color:{hex_color}; border:2px solid #555; border-radius:4px; margin-right:12px;"></div>
                                <b>文字：{record_type}</b>
                            </div>
                        ''', unsafe_allow_html=True)
                        
                        mapping_dict[(record_type, color_marker)] = st.selectbox(
                            "歸類為：", 
                            ['❌ 忽略不計', '🎯 原目標收入', '🔮 預估收入', '📉 實際支出', '💸 預估支出'],
                            key=f"map_{idx}",
                            index=2 if '預估' in record_type or hex_color == "#F2DCDB" else 1 if '收入' in record_type else 3
                        )

            df['財務屬性'] = df.apply(lambda row: mapping_dict.get((str(row['紀錄類型']), str(row['顏色標記'])), '❌ 忽略不計'), axis=1)

        # --- TAB 2: 總表 ---
        with tabs[1]:
            st.subheader("📋 營收分類彙整總表")
            unique_cats = df['營收分類'].unique().tolist()
            summary = pd.DataFrame(index=unique_cats)
            
            summary['原目標收入'] = df[df['財務屬性'] == '🎯 原目標收入'].groupby('營收分類')['年度總計'].sum()
            summary['預估收入'] = df[df['財務屬性'] == '🔮 預估收入'].groupby('營收分類')['年度總計'].sum()
            summary['實際支出'] = df[df['財務屬性'] == '📉 實際支出'].groupby('營營分類' if False else '營收分類')['年度總計'].sum()
            summary['預估支出'] = df[df['財務屬性'] == '💸 預估支出'].groupby('營收分類')['年度總計'].sum()
            
            summary = summary.fillna(0)
            summary['原毛利'] = summary['原目標收入'] - summary['實際支出']
            summary['預估毛利'] = summary['預估收入'] - summary['預估支出']
            summary['差異'] = summary['預估收入'] - summary['原目標收入']
            summary['毛利率'] = (summary['預估毛利'] / summary['預估收入']).replace([np.inf, -np.inf], 0).fillna(0)
            
            st.dataframe(summary.reset_index().rename(columns={'index':'營收分類'}).style.format({
                '原目標收入': '{:,.0f}', '預估收入': '{:,.0f}', '原毛利': '{:,.0f}', 
                '預估毛利': '{:,.0f}', '差異': '{:,.0f}', '毛利率': '{:.2%}'
            }), use_container_width=True, hide_index=True)

        # --- TAB 3: 卡片 ---
        with tabs[2]:
            st.subheader("💡 專案績效卡片")
            projects = df['專案說明'].unique()
            cols = st.columns(2)
            for idx, proj in enumerate(projects):
                with cols[idx % 2]:
                    p_df = df[df['專案說明'] == proj]
                    t_rev = p_df[p_df['財務屬性'] == '🎯 原目標收入']['年度總計'].sum()
                    e_rev = p_df[p_df['財務屬性'] == '🔮 預估收入']['年度總計'].sum()
                    e_exp = p_df[p_df['財務屬性'] == '💸 預估支出']['年度總計'].sum()
                    
                    with st.container(border=True):
                        st.markdown(f"#### {proj}")
                        m1, m2, m3 = st.columns(3)
                        m1.metric("目標", f"${t_rev:,.0f}")
                        m2.metric("預估", f"${e_rev:,.0f}", f"{e_rev-t_rev:,.0f}")
                        m3.metric("毛利", f"${e_rev-e_exp:,.0f}")
