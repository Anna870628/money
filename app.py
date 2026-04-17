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
    st.set_page_config(page_title="車聯網營收戰情系統 v15", layout="wide")
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
        """讀取引擎：包含顏色解析與『孤兒數字自動救援』機制"""
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
            for cell in [row[1], row[2]]: 
                fill = cell.fill
                if fill and hasattr(fill, 'start_color') and fill.start_color:
                    sc = fill.start_color
                    
                    if getattr(sc, 'type', None) == 'rgb' and getattr(sc, 'rgb', None):
                        val = str(sc.rgb)
                        if val != '00000000' and val != '000000': 
                            color_id = f"色碼_{val[-6:].upper()}"
                            break
                    elif getattr(sc, 'type', None) == 'theme':
                        theme = getattr(sc, 'theme', '未知')
                        tint = getattr(sc, 'tint', 0.0)
                        color_id = f"主題色_T{theme}_色偏{tint}"
                        break
                    elif getattr(sc, 'type', None) == 'indexed':
                        idx = getattr(sc, 'indexed', '未知')
                        color_id = f"索引色_{idx}"
                        break
                    else:
                        try:
                            raw = str(getattr(sc, 'rgb', '無RGB資料'))
                            if raw != '00000000' and raw != '無RGB資料':
                                safe_raw = "".join(char for char in raw if char.isalnum())
                                color_id = f"特殊格式_{safe_raw[-6:]}"
                                break
                        except:
                            color_id = "無法解析的特殊顏色"
                            break
                            
            color_list.append(color_id)

        uploaded_file.seek(0)
        df = pd.read_excel(uploaded_file, sheet_name=target_s, skiprows=2)
        df.columns = [str(c).strip() for c in df.columns]
        
        df['顏色標記'] = color_list[:len(df)]
        df.rename(columns={df.columns[2]: '紀錄類型'}, inplace=True)
        df['專案說明'] = df['專案說明'].replace(r'^\s*$', pd.NA, regex=True).ffill()
        
        cat_col = [c for c in df.columns if '營收分類' in str(c)]
        if cat_col:
            df.rename(columns={cat_col[0]: '營收分類'}, inplace=True)
            df['營收分類'] = df['營收分類'].ffill().fillna("其他")
        
        # 數字清理
        months = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec']
        for m in months:
            if m in df.columns:
                if df[m].dtype == object:
                    df[m] = df[m].astype(str).str.replace(',', '')
                df[m] = pd.to_numeric(df[m], errors='coerce').fillna(0)
            else:
                df[m] = 0.0

        # 孤兒數字救援引擎
        fallback_cols = [c for c in df.columns if any(k in str(c) for k in ['小計', '總計', '合計', '實績'])]
        df['temp_sum'] = df[months].sum(axis=1)
        for idx, row in df.iterrows():
            if row['temp_sum'] == 0:
                for f_col in fallback_cols:
                    try:
                        val_str = str(row[f_col]).replace(',', '').strip()
                        val = float(val_str) if val_str and val_str.lower() != 'nan' else 0.0
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
                with st.spinner("讀取中..."):
                    try:
                        new_df = process_imported_file(f)
                        save_to_supabase(new_df)
                        st.success("匯入成功！請至『營收分類總表』設定條件。")
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
            st.success("已更新至雲端！")

    with tabs[2]:
        if not data.empty:
            df_sum = data.copy()
            months = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec']
            df_sum['年度總計'] = df_sum[months].sum(axis=1)
            
            st.subheader("📋 營收分類戰情總表 (全自訂引擎)")
            
            unique_colors = df_sum['顏色標記'].dropna().unique().tolist()
            if "無底色" not in unique_colors:
                unique_colors.insert(0, "無底色")
                
            with st.container(border=True):
                st.markdown("##### 🛠️ 報表條件設定")
                st.caption("請根據您上傳的真實 Excel 格式，設定對應的顏色與關鍵字。")
                c1, c2 = st.columns(2)
                with c1:
                    target_color_setting = st.selectbox("🎯『原目標』的底色？", unique_colors, index=0)
                    income_keyword = st.text_input("💰 包含什麼字算『收入』？", value="收入")
                with c2:
                    est_color_setting = st.selectbox("🔮『預估』的底色？", unique_colors, index=0)
                    expense_keyword = st.text_input("💸 包含什麼字算『支出』？", value="支出")

            st.divider()

            # --- 🚀 自訂邏輯運算 ---
            # 根據使用者輸入的關鍵字來找列
            is_income_row = df_sum['紀錄類型'].astype(str).str.contains(income_keyword, na=False)
            is_expense_row = df_sum['紀錄類型'].astype(str).str.contains(expense_keyword, na=False)

            # 根據使用者選擇的顏色來分目標或預估
            is_target_color = df_sum['顏色標記'] == target_color_setting
            is_est_color = df_sum['顏色標記'] == est_color_setting

            # 組合計算
            target_rev = df_sum[is_target_color & is_income_row].groupby('營收分類')['年度總計'].sum()
            est_rev = df_sum[is_est_color & is_income_row].groupby('營收分類')['年度總計'].sum()
            actual_exp = df_sum[is_target_color & is_expense_row].groupby('營收分類')['年度總計'].sum()
            est_exp = df_sum[is_est_color & is_expense_row].groupby('營收分類')['年度總計'].sum()

            summary = pd.DataFrame({
                '原目標收入': target_rev,
                '預估收入': est_rev,
                '實際支出': actual_exp,
                '預估支出': est_exp
            }).fillna(0)

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
            
            with st.expander("🔍 數據診斷器 (如果還是0，打開看這裡的文字)"):
                st.write(df_sum.groupby(['顏色標記', '紀錄類型']).size().reset_index(name='筆數'))
