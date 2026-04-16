import streamlit as st
import pandas as pd
from sqlalchemy import text

# ==========================================
# 0. 系統環境設定與密碼檢查
# ==========================================
def check_password():
    def password_entered():
        if st.session_state["password"] == st.secrets["passwords"]["admin_password"]:
            st.session_state["password_correct"] = True
            del st.session_state["password"]
        else:
            st.session_state["password_correct"] = False

    if "password_correct" not in st.session_state:
        st.title("🔒 營收管理系統 - 身份驗證")
        st.text_input("請輸入存取密碼", type="password", on_change=password_entered, key="password")
        return False
    elif not st.session_state["password_correct"]:
        st.title("🔒 營收管理系統 - 身份驗證")
        st.text_input("請輸入存取密碼", type="password", on_change=password_entered, key="password")
        st.error("😕 密碼錯誤")
        return False
    return True

if check_password():
    st.set_page_config(page_title="車聯網營收管理系統 v4 (除錯版)", layout="wide")
    
    # 建立雲端 PostgreSQL 連線
    conn = st.connection("postgresql", type="sql")

    # ==========================================
    # 1. 資料庫與計算邏輯
    # ==========================================
    def load_data():
        try:
            df = conn.query("SELECT * FROM financials", ttl="0")
            return df
        except Exception:
            return pd.DataFrame(columns=['專案說明', '紀錄類型'] + ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'] + ['營收分類', '說明'])

    def save_to_supabase(df):
        with conn.session as session:
            session.execute(text("DELETE FROM financials")) 
            df.to_sql('financials', conn.engine, if_exists='append', index=False)
            session.commit()

    def process_imported_file(uploaded_file):
        if uploaded_file.name.endswith(('.xlsx', '.xls')):
            df = pd.read_excel(uploaded_file, skiprows=2)
        else:
            df = pd.read_csv(uploaded_file, skiprows=2)
        
        df.rename(columns={df.columns[2]: '紀錄類型'}, inplace=True)
        df['專案說明'] = df['專案說明'].ffill()
        
        category_col = [c for c in df.columns if '營收分類' in str(c)]
        if category_col:
            df.rename(columns={category_col[0]: '營收分類'}, inplace=True)
            df['營收分類'] = df['營收分類'].ffill()
        
        months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']
        for m in months:
            if m in df.columns:
                df[m] = pd.to_numeric(df[m], errors='coerce').fillna(0)
            else:
                df[m] = 0.0
                
        target_cols = ['專案說明', '紀錄類型'] + months + ['營收分類', '說明']
        return df[df['專案說明'].notna()][target_cols]

    # ==========================================
    # 2. 介面呈現與分頁
    # ==========================================
    st.title("📊 車聯網事業本部 - 專業營收管理系統")
    
    tab_edit, tab_summary = st.tabs(["📝 數據編輯與匯入", "📈 營收分析摘要"])

    with tab_edit:
        st.info("🟢 已連線至 Supabase | 提示：編輯完畢請點擊下方儲存按鈕")
        
        # ==========================================
        # ⚠️ 這裡換上了「除錯版」的側邊欄
        # ==========================================
        with st.sidebar:
            if st.button("🚪 安全登出"):
                st.session_state["password_correct"] = False
                st.rerun()
            
            st.divider()
            st.header("📂 數據匯入 (診斷模式)")
            uploaded_file = st.file_uploader("選擇原始報表 (CSV 或 Excel)", type=["csv", "xlsx", "xls"])
            
            if uploaded_file is not None:
                st.markdown("### 🔍 匯入前診斷")
                try:
                    # 測試讀取
                    test_df = process_imported_file(uploaded_file)
                    st.info(f"系統成功抓到 **{len(test_df)}** 筆有效資料")
                    
                    # 顯示預覽給你看
                    with st.expander("點我看清洗後的資料預覽"):
                        st.dataframe(test_df)

                    if st.button("🚀 確認沒問題，同步至雲端"):
                        with st.spinner("正在處理資料..."):
                            save_to_supabase(test_df)
                            st.success("匯入成功！")
                            st.rerun()
                            
                except Exception as e:
                    st.error(f"⚠️ 解析檔案時發生錯誤：{e}")

        # --- 主編輯器區塊 ---
        data = load_data()
        edited_df = st.data_editor(
            data,
            num_rows="dynamic",
            use_container_width=True,
            height=500,
            column_config={
                "紀錄類型": st.column_config.SelectboxColumn(
                    "紀錄類型",
                    options=["收入", "支出", "收入預估", "支出預估", "收入差異", "支出差異"],
                    required=True
                ),
                "營收分類": st.column_config.SelectboxColumn(
                    "營收分類",
                    options=["24DCM開發/維運", "TOYOTA聯網服務", "LEXUS聯網服務", "其他"],
                    required=True
                ),
                **{m: st.column_config.NumberColumn(format="%.0f") for m in ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec']}
            }
        )

        col1, col2, _ = st.columns([1, 1, 4])
        with col1:
            if st.button("💾 儲存所有變更至雲端", type="primary"):
                save_to_supabase(edited_df)
                st.success("資料已永久儲存！")
                
        with col2:
            csv_data = edited_df.to_csv(index=False).encode('utf-8-sig')
            st.download_button("📥 匯出當前報表 (CSV)", csv_data, "營收管理系統匯出.csv", "text/csv")

    # --- 分析摘要區塊 ---
    with tab_summary:
        st.header("📋 營收分類匯總分析")
        
        if not edited_df.empty:
            df = edited_df.copy()
            months = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec']
            # 將欄位轉換為數值，避免加總錯誤
            for m in months:
                df[m] = pd.to_numeric(df[m], errors='coerce').fillna(0)
            
            df['年度總計'] = df[months].sum(axis=1)

            target_rev = df[df['紀錄類型'] == '收入'].groupby('營收分類')['年度總計'].sum()
            est_rev = df[df['紀錄類型'] == '收入預估'].groupby('營收分類')['年度總計'].sum()
            actual_exp = df[df['紀錄類型'] == '支出'].groupby('營收分類')['年度總計'].sum()
            est_exp = df[df['紀錄類型'] == '支出預估'].groupby('營收分類')['年度總計'].sum()

            summary = pd.DataFrame({
                '目標營收': target_rev,
                '預估營收': est_rev,
                '實際支出': actual_exp,
                '預估支出': est_exp
            }).fillna(0)

            summary['營收差異 (目標-預估)'] = summary['目標營收'] - summary['預估營收']
            summary['原毛利'] = summary['目標營收'] - summary['實際支出']
            summary['預估毛利'] = summary['預估營收'] - summary['預估支出']
            summary['預估毛利率'] = (summary['預估毛利'] / summary['預估營收']).replace([float('inf'), -float('inf')], 0).fillna(0)

            def color_diff(val):
                color = 'red' if val < 0 else 'green' if val > 0 else 'black'
                return f'color: {color}'

            st.subheader("📊 分類統計總覽")
            st.dataframe(
                summary.style.format({
                    '目標營收': '{:,.0f}', '預估營收': '{:,.0f}', 
                    '營收差異 (目標-預估)': '{:,.0f}', '原毛利': '{:,.0f}', 
                    '預估毛利': '{:,.0f}', '預估毛利率': '{:.2%}'
                }).applymap(color_diff, subset=['營收差異 (目標-預估)', '預估毛利']),
                use_container_width=True
            )

            st.divider()
            st.subheader("🔍 營收差異項目明細 (目標 ≠ 預估)")
            
            diff_items = []
            for project in df['專案說明'].unique():
                proj_data = df[df['專案說明'] == project]
                t_val = proj_data[proj_data['紀錄類型'] == '收入']['年度總計'].sum()
                e_val = proj_data[proj_data['紀錄類型'] == '收入預估']['年度總計'].sum()
                
                if t_val != e_val:
                    diff_items.append({
                        '專案名稱': project,
                        '營收分類': proj_data['營收分類'].iloc[0] if not proj_data['營收分類'].empty else "",
                        '目標營收': t_val,
                        '預估營收': e_val,
                        '差異值': t_val - e_val,
                        '說明': proj_data['說明'].iloc[0] if not proj_data['說明'].empty else ""
                    })
            
            if diff_items:
                st.table(pd.DataFrame(diff_items))
            else:
                st.write("✅ 目前所有專案目標與預估皆一致。")
