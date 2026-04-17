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
    st.set_page_config(page_title="車聯網營收系統 v9", layout="wide")
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
        """讀取 Excel、辨識顏色、清洗數據"""
        wb = openpyxl.load_workbook(uploaded_file, data_only=True)
        sheet_names = wb.sheetnames
        target_s = sheet_names[0]
        for s in sheet_names:
            if any(k in s for k in ['營收', '預估', '收支']):
                target_s = s
                break
        sheet = wb[target_s]
        
        color_list = []
        # 資料從第 4 行開始 (skiprows=2)
        for row in sheet.iter_rows(min_row=4):
            # 抓取專案說明格(Column B)的背景色 RGB
            fill = row[1].fill
            # openpyxl 的 RGB 通常帶有 Alpha 通道 (如 FFD9D9D9)，我們只取後 6 碼
            color_rgb = "無底色"
            if fill and fill.start_color and fill.start_color.rgb:
                rgb_val = str(fill.start_color.rgb)
                color_rgb = rgb_val[-6:] if len(rgb_val) >= 6 else rgb_val
            color_list.append(color_rgb.upper())

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

        edited = st.data_editor(
            data, num_rows="dynamic", use_container_width=True, height=500,
            column_config={
                "營收分類": st.column_config.SelectboxColumn(options=["24DCM開發/維運", "TOYOTA聯網服務", "LEXUS聯網服務", "其他"]),
                "紀錄類型": st.column_config.SelectboxColumn(options=["收入", "收入預估", "支出", "支出預估", "收入差異", "支出差異"])
            }
        )
        if st.button("💾 儲存變更", type="primary"):
            save_to_supabase(edited)
            st.success("已更新至雲端！")

    with tabs[2]:
        if not data.empty:
            df_sum = data.copy()
            months = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec']
            df_sum['年度總計'] = df_sum[months].sum(axis=1)
            
            st.subheader("📋 營收分類總表")
            
            # --- 核心邏輯計算 ---
            
            # 1. 原目標收入：底色 D9D9D9 且 紀錄類型(或專案說明) 為 收入
            target_rev = df_sum[
                (df_sum['顏色標記'] == 'D9D9D9') & 
                ((df_sum['紀錄類型'] == '收入') | (df_sum['專案說明'] == '收入'))
            ].groupby('營營分類')['年度總計'].sum()

            # 2. 預估收入：底色 F2DCDB 且 紀錄類型(或專案說明) 為 收入預估
            est_rev = df_sum[
                (df_sum['顏色標記'] == 'F2DCDB') & 
                ((df_sum['紀錄類型'] == '收入預估') | (df_sum['專案說明'] == '收入預估'))
            ].groupby('營收分類')['年度總計'].sum()

            # 為了計算毛利，我們需要支出數據 (通常支出不一定有特定底色要求，按類型抓取)
            actual_exp = df_sum[df_sum['紀錄類型'] == '支出'].groupby('營收分類')['年度總計'].sum()
            est_exp = df_sum[df_sum['紀錄類型'] == '支出預估'].groupby('營收分類')['年度總計'].sum()

            # 合併成最終表
            final_summary = pd.DataFrame({
                '原目標收入': target_rev,
                '預估收入': est_rev,
                '實際支出': actual_exp,   # 隱藏計算用
                '預估支出': est_exp      # 隱藏計算用
            }).fillna(0)

            # 3. 原毛利 = 原目標收入 - 實際支出
            final_summary['原毛利'] = final_summary['原目標收入'] - final_summary['實際支出']
            
            # 4. 預估毛利 = 預估收入 - 預估支出
            final_summary['預估毛利'] = final_summary['預估收入'] - final_summary['預估支出']
            
            # 5. 差異 = 預估收入 - 原目標收入
            final_summary['差異'] = final_summary['預估收入'] - final_summary['原目標收入']
            
            # 6. 毛利率 = 預估毛利 / 預估收入
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
