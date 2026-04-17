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
    st.set_page_config(page_title="車聯網營收系統 v9.2", layout="wide")
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
            color_rgb = "無底色"
            # 【優化1】同時檢查「專案說明(row[1])」和「紀錄類型(row[2])」有沒有塗顏色
            for cell in [row[1], row[2]]: 
                fill = cell.fill
                if fill and fill.start_color and fill.start_color.rgb:
                    rgb_val = str(fill.start_color.rgb)
                    val = rgb_val[-6:].upper()
                    if val != '000000': # 排除預設黑/透明
                        color_rgb = val
                        break
            color_list.append(color_rgb)

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
        
        months = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec']
        for m in months:
            if m in df.columns:
                # 【優化2】移除千分位逗號，確保數字能被正確轉換
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
    st.title("📊 車聯網事業本部 - 專案營收戰情室")
    tabs = st.tabs(["🎴 專案卡片摘要", "📝 原始數據管理", "📈 營收分類總表"])
    data = load_data()

    with tabs[1]:
        with st.sidebar:
            st.header("📂 匯入數據")
            f = st.file_uploader("選擇 Excel", type=["xlsx"])
            if f and st.button("🚀 開始解析並上傳"):
                try:
                    new_df = process_imported_file(f)
                    save_to_supabase(new_df)
                    st.success("匯入成功！")
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
            
            # 【優化3】透視鏡診斷工具
            with st.expander("🛠️ 抓不到資料？點此展開診斷 (檢查 Excel 實際讀到的色碼與文字)"):
                st.markdown("如果下方表格都是 0，請對照以下系統實際抓到的 **色碼** 與 **文字**，看是否與你設定的 `D9D9D9` 及 `收入` 有落差。")
                debug_df = df_sum.groupby(['顏色標記', '紀錄類型']).size().reset_index(name='資料筆數')
                st.dataframe(debug_df, use_container_width=True)

            st.subheader("📋 營收分類總表")
            
            # --- 核心邏輯計算 (加入模糊比對) ---
            
            # 將紀錄類型轉為字串，方便進行模糊搜尋 (包含 '收入' 且不含 '預估')
            is_income = df_sum['紀錄類型'].astype(str).str.contains('收入', na=False) & \
                        (~df_sum['紀錄類型'].astype(str).str.contains('預估', na=False))
            
            # 包含 '收入' 且包含 '預估'
            is_est_income = df_sum['紀錄類型'].astype(str).str.contains('收入', na=False) & \
                            df_sum['紀錄類型'].astype(str).str.contains('預估', na=False)

            # 1. 原目標收入：底色包含 D9D9D9 且 類型為收入
            target_rev = df_sum[
                (df_sum['顏色標記'].str.contains('D9D9D9', case=False, na=False)) & is_income
            ].groupby('營收分類')['年度總計'].sum()

            # 2. 預估收入：底色包含 F2DCDB 且 類型為預估收入
            est_rev = df_sum[
                (df_sum['顏色標記'].str.contains('F2DCDB', case=False, na=False)) & is_est_income
            ].groupby('營收分類')['年度總計'].sum()

            # 支出數據 (模糊比對 '支出')
            is_exp = df_sum['紀錄類型'].astype(str).str.contains('支出', na=False) & \
                     (~df_sum['紀錄類型'].astype(str).str.contains('預估', na=False))
            is_est_exp = df_sum['紀錄類型'].astype(str).str.contains('支出', na=False) & \
                         df_sum['紀錄類型'].astype(str).str.contains('預估', na=False)

            actual_exp = df_sum[is_exp].groupby('營收分類')['年度總計'].sum()
            est_exp = df_sum[is_est_exp].groupby('營收分類')['年度總計'].sum()

            # 合併成最終表
            final_summary = pd.DataFrame({
                '原目標收入': target_rev,
                '預估收入': est_rev,
                '實際支出': actual_exp,   
                '預估支出': est_exp      
            }).fillna(0)

            # 計算毛利與差異
            final_summary['原毛利'] = final_summary['原目標收入'] - final_summary['實際支出']
            final_summary['預估毛利'] = final_summary['預估收入'] - final_summary['預估支出']
            final_summary['差異'] = final_summary['預估收入'] - final_summary['原目標收入']
            final_summary['毛利率'] = (final_summary['預估毛利'] / final_summary['預估收入']).replace([float('inf'), -float('inf')], 0).fillna(0)

            # 只保留你要求的欄位並重排序
            display_cols = ['原目標收入', '預估收入', '原毛利', '預估毛利', '差異', '毛利率']
            final_display = final_summary[display_cols]

            # 樣式設定
            st.dataframe(
                final_display.style.format({
                    '原目標收入': '{:,.0f}', '預估收入': '{:,.0f}', 
                    '原毛利': '{:,.0f}', '預估毛利': '{:,.0f}', 
                    '差異': '{:,.0f}', '毛利率': '{:.2%}'
                }).map(lambda x: 'color: red' if isinstance(x, (int, float)) and x < 0 else '', subset=['差異', '預估毛利']),
                use_container_width=True
            )
