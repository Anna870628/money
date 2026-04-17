import streamlit as st
import pandas as pd
import numpy as np
from sqlalchemy import text
import openpyxl
from openpyxl.styles.colors import COLOR_INDEX, Color
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
    st.set_page_config(page_title="車聯網營收系統 v28", layout="wide")
    conn = st.connection("postgresql", type="sql")

    # --- 🚀 核心顏色轉換函數：將 Theme 轉成 HEX ---
    def theme_and_tint_to_hex(wb, theme, tint):
        """將 Excel 的主題色與色偏值轉換成真正的顏色代碼"""
        try:
            from openpyxl.styles.colors import RGB
            # 獲取主題顏色列表
            xl_theme = wb._external_values[0].theme.themeElements.clrScheme
            # 這裡簡化處理常見的主題顏色索引
            theme_map = ['lt1', 'dk1', 'lt2', 'dk2', 'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6']
            if theme < len(theme_map):
                color = getattr(xl_theme, theme_map[theme]).srgbClr.val
                return f"#{color}"
        except:
            # 萬一換算失敗，根據常見的主題色索引給予預估色 (防呆)
            presets = {
                0: "#FFFFFF", 1: "#000000", 2: "#E7E6E6", 3: "#44546A",
                4: "#5B9BD5", 5: "#ED7D31", 6: "#A5A5A5", 7: "#FFC000"
            }
            return presets.get(theme, "#D0D0D0")
        return "#D0D0D0"

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
            final_hex = "無底色"
            # 掃描前 15 欄，尋找非白色背景
            for cell in row[1:16]: 
                fill = cell.fill
                if fill and hasattr(fill, 'start_color') and fill.start_color:
                    sc = fill.start_color
                    # A. 標準 RGB
                    if sc.type == 'rgb' and sc.rgb and str(sc.rgb) != '00000000':
                        final_hex = f"#{str(sc.rgb)[-6:].upper()}"
                        break
                    # B. 主題色 (Theme)
                    elif sc.type == 'theme' and sc.theme is not None:
                        # 標記為主題色，讓 UI 渲染時去換算
                        final_hex = f"THEME_{sc.theme}_{sc.tint}"
                        break
            color_list.append(final_hex)

        uploaded_file.seek(0)
        df = pd.read_excel(uploaded_file, sheet_name=target_s, skiprows=2)
        df.columns = [str(c).strip() for c in df.columns]
        df['顏色標記'] = color_list[:len(df)]
        df.rename(columns={df.columns[2]: '紀錄類型'}, inplace=True)
        df['專案說明'] = df['專案說明'].replace(r'^\s*$', pd.NA, regex=True).ffill()
        df['紀錄類型'] = df['紀錄類型'].astype(str).str.strip()
        
        # 分類清洗
        cat_col_name = next((c for c in df.columns if '營收分類' in str(c)), None)
        if cat_col_name:
            df['營收分類'] = df[cat_col_name].astype(str).str.replace('\n', ' ').str.strip()
            df['營收分類'] = df['營收分類'].replace(['nan', 'None', '', 'NaN'], np.nan).ffill().fillna("其他")
        else:
            df['營收分類'] = "其他"

        # 數字清洗
        months = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec']
        def clean_val(v):
            if pd.isna(v): return 0.0
            s = re.sub(r'[^\d\.\-\(\)]', '', str(v))
            if s.startswith('(') and s.endswith(')'): s = '-' + s[1:-1]
            try: return float(s) if s else 0.0
            except: return 0.0
            
        for m in months:
            if m in df.columns: df[m] = df[m].apply(clean_val)
            else: df[m] = 0.0

        # 小計救援
        fallback_cols = [c for c in df.columns if any(k in str(c) for k in ['小計', '總計', '合計', '實績'])]
        for idx, row in df.iterrows():
            if sum(row[months]) == 0:
                for f_col in fallback_cols:
                    val = clean_val(row[f_col])
                    if val != 0:
                        df.at[idx, 'Jan'] = val
                        break
                            
        target_cols = ['專案說明', '紀錄類型'] + months + ['營收分類', '顏色標記']
        return df.dropna(subset=['專案說明', '紀錄類型'])[target_cols]

    # ==========================================
    # 2. 介面呈現
    # ==========================================
    with st.sidebar:
        st.header("📂 匯入數據")
        f = st.file_uploader("選擇 Excel", type=["xlsx"])
        if f and st.button("🚀 啟動全彩掃描並上傳"):
            save_to_supabase(process_imported_file(f))
            st.success("匯入成功！色塊已更新。")
            st.rerun()

    data = load_data()
    if not data.empty:
        df = data.copy()
        months = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec']
        df['年度總計'] = df[months].sum(axis=1)

        tabs = st.tabs(["🎨 1. 視覺化資料對應", "📈 2. 營收分類總表", "🎴 3. 專案卡片摘要"])

        with tabs[0]:
            st.markdown("### 🧩 財務邏輯對應設定")
            st.info("請根據下方出現的「真實色塊」來選擇對應的財務性質。")
            unique_combos = df[['紀錄類型', '顏色標記']].drop_duplicates().values.tolist()
            mapping_dict = {}
            cols = st.columns(2)
            
            for idx, (record_type, color_marker) in enumerate(unique_combos):
                # 顏色渲染邏輯
                hex_color = "#FFFFFF"
                if color_marker.startswith("#"):
                    hex_color = color_marker
                elif color_marker.startswith("THEME_"):
                    # 簡單根據主題編號區分色系 (粉紅 vs 灰)
                    t_idx = int(color_marker.split('_')[1])
                    # 假設主題色 2/3 是灰色系，4/5 是藍粉色系
                    hex_color = "#F2DCDB" if t_idx in [4, 5, 7, 8, 9] else "#D9D9D9"
                
                with cols[idx % 2]:
                    with st.container(border=True):
                        st.markdown(f'''
                            <div style="display:flex; align-items:center; margin-bottom:10px;">
                                <div style="width:35px; height:35px; background-color:{hex_color}; border:2px solid #555; border-radius:4px; margin-right:15px;"></div>
                                <b>文字：{record_type}</b>
                            </div>
                        ''', unsafe_allow_html=True)
                        
                        mapping_dict[(record_type, color_marker)] = st.selectbox(
                            "將此組合歸類為：", 
                            ['❌ 忽略不計', '🎯 原目標收入', '🔮 預估收入', '📉 實際支出', '💸 預估支出'],
                            key=f"map_{idx}",
                            index=2 if '預估' in record_type or hex_color == "#F2DCDB" else 1 if '收入' in record_type else 3
                        )

            df['財務屬性'] = df.apply(lambda row: mapping_dict.get((str(row['紀錄類型']), str(row['顏色標記'])), '❌ 忽略不計'), axis=1)

        # --- 剩下的 TAB 2 & TAB 3 邏輯同步更新 ---
        with tabs[1]:
            summary = df.groupby(['營收分類', '財務屬性'])['年度總計'].sum().unstack().fillna(0)
            for col in ['🎯 原目標收入', '🔮 預估收入', '📉 實際支出', '💸 預估支出']:
                if col not in summary.columns: summary[col] = 0
            
            summary['原毛利'] = summary['🎯 原目標收入'] - summary['📉 實際支出']
            summary['預估毛利'] = summary['🔮 預估收入'] - summary['💸 預估支出']
            summary['差異'] = summary['🔮 預估收入'] - summary['🎯 原目標收入']
            summary['毛利率'] = (summary['預估毛利'] / summary['🔮 預估收入']).replace([np.inf, -np.inf], 0).fillna(0)
            
            st.dataframe(summary[['🎯 原目標收入', '🔮 預估收入', '原毛利', '預估毛利', '差異', '毛利率']].style.format({
                '毛利率': '{:.2%}', '差異': '{:,.0f}', '🎯 原目標收入': '{:,.0f}', '🔮 預估收入': '{:,.0f}'
            }), use_container_width=True)

        with tabs[2]:
            for proj in df['專案說明'].unique():
                p_df = df[df['專案說明'] == proj]
                t_rev = p_df[p_df['財務屬性'] == '🎯 原目標收入']['年度總計'].sum()
                e_rev = p_df[p_df['財務屬性'] == '🔮 預估收入']['年度總計'].sum()
                e_exp = p_df[p_df['財務屬性'] == '💸 預估支出']['年度總計'].sum()
                
                with st.container(border=True):
                    st.write(f"### {proj}")
                    c1, c2, c3 = st.columns(3)
                    c1.metric("目標營收", f"${t_rev:,.0f}")
                    c2.metric("預估營收", f"${e_rev:,.0f}", f"{e_rev-t_rev:,.0f}")
                    c3.metric("預估毛利", f"${e_rev-e_exp:,.0f}")

    def save_to_supabase(df):
        with conn.session as session:
            session.execute(text("DELETE FROM financials"))
            session.commit()
        df.to_sql('financials', conn.engine, if_exists='append', index=False)
