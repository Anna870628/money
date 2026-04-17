import streamlit as st
import pandas as pd
from sqlalchemy import text
import openpyxl

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
    st.set_page_config(page_title="車聯網營收戰情系統 v17", layout="wide")
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
            return pd.DataFrame(columns=['專案說明', '紀錄類型', '營收分類', '顏色標記', '說明'] + ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'])

    def save_to_supabase(df):
        with conn.session as session:
            session.execute(text("DELETE FROM financials"))
            session.commit()
        df.to_sql('financials', conn.engine, if_exists='append', index=False, chunksize=50, method='multi')

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
            color_id = "無底色"
            for cell in [row[1], row[2], row[19]]: 
                fill = cell.fill
                if fill and hasattr(fill, 'start_color') and fill.start_color:
                    sc = fill.start_color
                    if getattr(sc, 'type', None) == 'rgb' and getattr(sc, 'rgb', None):
                        val = str(sc.rgb)
                        if val != '00000000' and val != '000000': 
                            color_id = f"色碼_{val[-6:].upper()}"
                            break
                    elif getattr(sc, 'type', None) == 'theme':
                        color_id = f"主題色_T{getattr(sc, 'theme', '未知')}_色偏{getattr(sc, 'tint', 0.0)}"
                        break
                    elif getattr(sc, 'type', None) == 'indexed':
                        color_id = f"索引色_{getattr(sc, 'indexed', '未知')}"
                        break
            color_list.append(color_id)

        uploaded_file.seek(0)
        df = pd.read_excel(uploaded_file, sheet_name=target_s, skiprows=2)
        df.columns = [str(c).strip() for c in df.columns]
        
        df['顏色標記'] = color_list[:len(df)]
        df.rename(columns={df.columns[2]: '紀錄類型'}, inplace=True)
        df['專案說明'] = df['專案說明'].replace(r'^\s*$', pd.NA, regex=True).ffill()
        
        # --- 確保分類絕對不丟失 ---
        cat_col_name = ""
        for c in df.columns:
            if '營收分類' in str(c):
                cat_col_name = c
                break
        
        if cat_col_name:
            df[cat_col_name] = df[cat_col_name].astype(str).str.replace('\n', ' ').str.strip()
            df[cat_col_name] = df[cat_col_name].replace(['nan', 'None', '', '<NA>', 'NaN'], pd.NA)
            df['營收分類'] = df[cat_col_name].ffill().fillna("其他")
        else:
            df['營收分類'] = "其他"
        
        # --- 處理數字 ---
        months = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec']
        for m in months:
            if m in df.columns:
                if df[m].dtype == object:
                    df[m] = df[m].astype(str).str.replace(',', '')
                df[m] = pd.to_numeric(df[m], errors='coerce').fillna(0)
            else:
                df[m] = 0.0

        # --- 🚀 孤兒數字救援引擎 (加回來了！) ---
        fallback_cols = [c for c in df.columns if any(k in str(c) for k in ['小計', '總計', '合計', '實績'])]
        df['temp_sum'] = df[months].sum(axis=1)
        for idx, row in df.iterrows():
            if row['temp_sum'] == 0:
                for f_col in fallback_cols:
                    if f_col in row:
                        try:
                            val_str = str(row[f_col]).replace(',', '').strip()
                            if val_str and val_str.lower() not in ['nan', 'none', '']:
                                val = float(val_str)
                                if val != 0:
                                    df.at[idx, 'Jan'] = val
                                    break
                        except:
                            pass
        df = df.drop(columns=['temp_sum'])
        
        if '說明' not in df.columns: df['說明'] = ""
        target_cols = ['專案說明', '紀錄類型'] + months + ['營收分類', '顏色標記', '說明']
        return df.dropna(subset=['專案說明', '紀錄類型'])[target_cols]

    # ==========================================
    # 2. 介面呈現
    # ==========================================
    tabs = st.tabs(["🎴 專案卡片摘要", "📝 原始數據管理", "📈 營收分類總表"])
    data = load_data()

    with tabs[1]:
        with st.sidebar:
            st.header("📂 匯入數據")
            f = st.file_uploader("選擇 Excel", type=["xlsx"])
            if f and st.button("🚀 開始解析並上傳"):
                with st.spinner("啟動救援引擎與數據清洗..."):
                    try:
                        new_df = process_imported_file(f)
                        save_to_supabase(new_df)
                        st.success("匯入成功！所有專案已歸位。")
                        st.rerun()
                    except Exception as e:
                        st.error(f"解析失敗: {e}")
            
            st.divider()
            if st.button("⚠️ 清空資料庫"):
                with conn.session as s:
                    s.execute(text("DELETE FROM financials"))
                    s.commit()
                st.rerun()

        edited = st.data_editor(data, num_rows="dynamic", use_container_width=True, height=500)
        if st.button("💾 儲存變更", type="primary"):
            save_to_supabase(edited)
            st.success("雲端已更新！")

    with tabs[2]:
        if not data.empty:
            df_sum = data.copy()
            months = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec']
            df_sum['年度總計'] = df_sum[months].sum(axis=1)
            
            st.subheader("📋 營收分類戰情總表")
            
            # 抓取所有唯一的分類，並強制作為表格的基底，避免任何分類消失
            unique_cats = df_sum['營收分類'].unique().tolist()
            st.caption(f"系統目前偵測到的分類標籤：{', '.join(unique_cats)}")

            unique_colors = df_sum['顏色標記'].dropna().unique().tolist()
            if "無底色" not in unique_colors: unique_colors.insert(0, "無底色")
                
            with st.container(border=True):
                st.markdown("##### 🛠️ 報表邏輯設定")
                c1, c2 = st.columns(2)
                with c1:
                    target_color = st.selectbox("🎯『原目標』底色？", unique_colors, index=0)
                    inc_key = st.text_input("💰 收入關鍵字", value="收入")
                with c2:
                    est_color = st.selectbox("🔮『預估』底色？", unique_colors, index=0)
                    est_key = st.text_input("📑 預估排除關鍵字 (用於排除目標收入)", value="預估")

            st.divider()

            is_income_pure = df_sum['紀錄類型'].str.contains(inc_key, na=False) & ~df_sum['紀錄類型'].str.contains(est_key, na=False)
            is_income_est = df_sum['紀錄類型'].str.contains(inc_key, na=False) & df_sum['紀錄類型'].str.contains(est_key, na=False)
            is_exp_pure = df_sum['紀錄類型'].str.contains('支出', na=False) & ~df_sum['紀錄類型'].str.contains(est_key, na=False)
            is_exp_est = df_sum['紀錄類型'].str.contains('支出', na=False) & df_sum['紀錄類型'].str.contains(est_key, na=False)

            # --- 🚀 強制打樁：無論篩選結果為何，保證所有分類皆會顯示 ---
            summary = pd.DataFrame(index=unique_cats)
            
            summary['原目標收入'] = df_sum[(df_sum['顏色標記'] == target_color) & is_income_pure].groupby('營收分類')['年度總計'].sum()
            summary['預估收入'] = df_sum[(df_sum['顏色標記'] == est_color) & is_income_est].groupby('營收分類')['年度總計'].sum()
            summary['實際支出'] = df_sum[(df_sum['顏色標記'] == target_color) & is_exp_pure].groupby('營收分類')['年度總計'].sum()
            summary['預估支出'] = df_sum[(df_sum['顏色標記'] == est_color) & is_exp_est].groupby('營收分類')['年度總計'].sum()
            
            summary = summary.fillna(0)

            # 公式計算
            summary['原毛利'] = summary['原目標收入'] - summary['實際支出']
            summary['預估毛利'] = summary['預估收入'] - summary['預估支出']
            summary['差異'] = summary['預估收入'] - summary['原目標收入']
            summary['毛利率'] = (summary['預估毛利'] / summary['預估收入']).replace([float('inf'), -float('inf')], 0).fillna(0)

            display_cols = ['原目標收入', '預估收入', '原毛利', '預估毛利', '差異', '毛利率']
            st.dataframe(
                summary[display_cols].style.format({
                    '原目標收入': '{:,.0f}', '預估收入': '{:,.0f}', 
                    '原毛利': '{:,.0f}', '預估毛利': '{:,.0f}', 
                    '差異': '{:,.0f}', '毛利率': '{:.2%}'
                }).map(lambda x: 'color: red' if isinstance(x, (int, float)) and x < 0 else '', subset=['差異', '預估毛利']),
                use_container_width=True
            )
