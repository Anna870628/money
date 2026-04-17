import streamlit as st
import pandas as pd
from sqlalchemy import text
import openpyxl

# ==========================================
# 0. 安全驗證 (密碼: CMX_BPT)
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
    st.set_page_config(page_title="車聯網營收戰情系統 v9.3", layout="wide")
    conn = st.connection("postgresql", type="sql")

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
        df.to_sql('financials', conn.engine, if_exists='append', index=False, chunksize=100, method='multi')

    def process_imported_file(uploaded_file):
        """讀取 Excel 並精準抓取儲存格背景底色 (Fill Color)"""
        wb = openpyxl.load_workbook(uploaded_file, data_only=True)
        sheet_names = wb.sheetnames
        target_s = sheet_names[0]
        for s in sheet_names:
            if any(k in s for k in ['營收', '預估', '收支']):
                target_s = s
                break
        sheet = wb[target_s]
        
        color_list = []
        # 從第 4 行開始 (對應 skiprows=2)
        for row in sheet.iter_rows(min_row=4):
            # 優先抓取「紀錄類型」(row[2]) 的底色
            target_cell = row[2] 
            fill = target_cell.fill
            color_hex = "無底色"
            
            if fill and fill.start_color:
                rgb = str(fill.start_color.rgb)
                # openpyxl 的 RGB 可能帶 Alpha (如 FFD9D9D9)，我們取最後 6 碼
                if len(rgb) >= 6:
                    color_hex = rgb[-6:].upper()
            
            # 若紀錄類型沒顏色，再看專案說明 (row[1])
            if color_hex == "000000" or color_hex == "無底色":
                fill_alt = row[1].fill
                if fill_alt and fill_alt.start_color:
                    rgb_alt = str(fill_alt.start_color.rgb)
                    if len(rgb_alt) >= 6:
                        color_hex = rgb_alt[-6:].upper()
            
            color_list.append(color_hex)

        uploaded_file.seek(0)
        df = pd.read_excel(uploaded_file, sheet_name=target_s, skiprows=2)
        df.columns = [str(c).strip() for c in df.columns]
        
        df['顏色標記'] = color_list[:len(df)]
        df.rename(columns={df.columns[2]: '紀錄類型'}, inplace=True)
        df['專案說明'] = df['專案說明'].replace(r'^\s*$', pd.NA, regex=True).ffill()
        
        # 營收分類與數值處理
        cat_col = [c for c in df.columns if '營收分類' in str(c)]
        if cat_col:
            df.rename(columns={cat_col[0]: '營收分類'}, inplace=True)
            df['營收分類'] = df['營收分類'].ffill().fillna("其他")
        
        months = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec']
        for m in months:
            if m in df.columns:
                if df[m].dtype == object:
                    df[m] = df[m].astype(str).str.replace(',', '')
                df[m] = pd.to_numeric(df[m], errors='coerce').fillna(0)
            else:
                df[m] = 0.0
        
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
            if f and st.button("🚀 開始解析底色並上傳"):
                with st.spinner("正在讀取 Excel 背景色..."):
                    try:
                        new_df = process_imported_file(f)
                        save_to_supabase(new_df)
                        st.success("匯入成功！已鎖定底色邏輯。")
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
            
            # --- 財務邏輯運算 ---
            # 定義：原目標收入 (底色 D9D9D9 且 包含 '收入' 且不含 '預估')
            is_target_color = df_sum['顏色標記'].str.contains('D9D9D9', case=False, na=False)
            is_est_color = df_sum['顏色標記'].str.contains('F2DCDB', case=False, na=False)
            
            is_income_text = df_sum['紀錄類型'].str.contains('收入', na=False)
            is_est_text = df_sum['紀錄類型'].str.contains('預估', na=False)
            is_expense_text = df_sum['紀錄類型'].str.contains('支出', na=False)

            # 1. 原目標收入
            target_rev = df_sum[is_target_color & is_income_text & ~is_est_text].groupby('營收分類')['年度總計'].sum()
            # 2. 預估收入
            est_rev = df_sum[is_est_color & is_income_text & is_est_text].groupby('營收分類')['年度總計'].sum()
            # 支出 (用於計算毛利)
            actual_exp = df_sum[is_expense_text & ~is_est_text].groupby('營收分類')['年度總計'].sum()
            est_exp = df_sum[is_expense_text & is_est_text].groupby('營收分類')['年度總計'].sum()

            summary = pd.DataFrame({
                '原目標收入': target_rev,
                '預估收入': est_rev,
                '實際支出': actual_exp,
                '預估支出': est_exp
            }).fillna(0)

            # 3. 原毛利 = 原目標收入 - 實際支出
            summary['原毛利'] = summary['原目標收入'] - summary['實際支出']
            # 4. 預估毛利 = 預估收入 - 預估支出
            summary['預估毛利'] = summary['預估收入'] - summary['預估支出']
            # 5. 差異 = 預估收入 - 原目標收入
            summary['差異'] = summary['預估收入'] - summary['原目標收入']
            # 6. 毛利率 = 預估毛利 / 預估收入
            summary['毛利率'] = (summary['預估毛利'] / summary['預估收入']).replace([float('inf'), -float('inf')], 0).fillna(0)

            # 只顯示要求的欄位
            display_cols = ['原目標收入', '預估收入', '原毛利', '預估毛利', '差異', '毛利率']
            
            st.subheader("📋 營收分類戰情總表")
            st.dataframe(
                summary[display_cols].style.format({
                    '原目標收入': '{:,.0f}', '預估收入': '{:,.0f}', 
                    '原毛利': '{:,.0f}', '預估毛利': '{:,.0f}', 
                    '差異': '{:,.0f}', '毛利率': '{:.2%}'
                }).map(lambda x: 'color: red' if isinstance(x, (int, float)) and x < 0 else '', subset=['差異', '預估毛利']),
                use_container_width=True
            )

            # 診斷工具 (幫助確認色碼是否正確)
            with st.expander("🔍 診斷：系統偵測到的底色代碼統計"):
                st.write(df_sum.groupby(['顏色標記', '紀錄類型']).size())
