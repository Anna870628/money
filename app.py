import streamlit as st
import pandas as pd

# ==========================================
# 0. 系統環境設定與密碼檢查
# ==========================================
def check_password():
    """檢查使用者是否輸入正確的存取密碼"""
    def password_entered():
        # 從 st.secrets 中讀取密碼，確保 GitHub 上看不到密碼
        if st.session_state["password"] == st.secrets["passwords"]["admin_password"]:
            st.session_state["password_correct"] = True
            del st.session_state["password"]  # 安全考量，不保留密碼在狀態中
        else:
            st.session_state["password_correct"] = False

    if "password_correct" not in st.session_state:
        st.title("🔒 營收管理系統 - 身份驗證")
        st.text_input("請輸入存取密碼 (Password)", type="password", on_change=password_entered, key="password")
        return False
    elif not st.session_state["password_correct"]:
        st.title("🔒 營收管理系統 - 身份驗證")
        st.text_input("請輸入存取密碼 (Password)", type="password", on_change=password_entered, key="password")
        st.error("😕 密碼錯誤，請重新輸入。")
        return False
    return True

# 只有密碼正確才會執行後續代碼
if check_password():
    st.set_page_config(page_title="車聯網專案營收管理系統", layout="wide")
    
    # ==========================================
    # 1. 資料庫連線與資料處理邏輯
    # ==========================================
    # 建立雲端 PostgreSQL 連線 (自動讀取 secrets 中的 [connections.postgresql])
    conn = st.connection("postgresql", type="sql")

    def load_data():
        """從 Supabase 讀取資料"""
        try:
            return conn.query("SELECT * FROM financials", ttl="0")
        except Exception:
            # 如果資料表是空的或不存在，回傳一個空殼 DataFrame
            return pd.DataFrame(columns=['專案說明', '紀錄類型'] + [m for m in ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']] + ['營營分類', '說明'])

    def save_to_supabase(df):
        """將變更後的資料同步回雲端"""
        with conn.session as session:
            session.execute("DELETE FROM financials") # 覆蓋模式：先清空
            df.to_sql('financials', conn.engine, if_exists='append', index=False)
            session.commit()

    def process_imported_file(uploaded_file):
        """處理匯入的檔案 (支援 CSV 與 Excel)"""
        # 判斷副檔名
        if uploaded_file.name.endswith(('.xlsx', '.xls')):
            df = pd.read_excel(uploaded_file, skiprows=2)
        else:
            df = pd.read_csv(uploaded_file, skiprows=2)
        
        # 欄位清理與重命名
        df.rename(columns={df.columns[2]: '紀錄類型'}, inplace=True)
        # 向下填補 (處理 Excel 合併儲存格造成的空白)
        df['專案說明'] = df['專案說明'].ffill()
        # 處理營收分類可能存在的欄位名稱差異
        category_col = [c for c in df.columns if '營收分類' in str(c)]
        if category_col:
            df.rename(columns={category_col[0]: '營收分類'}, inplace=True)
            df['營收分類'] = df['營收分類'].ffill()
        
        # 處理月份數值
        months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']
        for m in months:
            if m in df.columns:
                df[m] = pd.to_numeric(df[m], errors='coerce').fillna(0)
            else:
                df[m] = 0.0
                
        target_cols = ['專案說明', '紀錄類型'] + months + ['營收分類', '說明']
        # 排除掉全空或是只有序號的行
        return df[df['專案說明'].notna()][target_cols]

    # ==========================================
    # 2. 介面呈現
    # ==========================================
    st.title("📊 車聯網事業本部 - 專案每月收支管理系統")
    st.info("🟢 已成功連線至 Supabase 雲端資料庫")

    # 側邊欄：功能選單
    with st.sidebar:
        if st.button("🚪 安全登出"):
            st.session_state["password_correct"] = False
            st.rerun()
        
        st.divider()
        st.header("📂 數據匯入 (Table A/B)")
        uploaded_file = st.file_uploader("選擇原始報表 (CSV 或 Excel)", type=["csv", "xlsx", "xls"])
        if uploaded_file is not None:
            if st.button("🚀 確認匯入並同步雲端"):
                with st.spinner("正在處理資料..."):
                    try:
                        cleaned_df = process_imported_file(uploaded_file)
                        save_to_supabase(cleaned_df)
                        st.success("匯入成功！")
                        st.rerun()
                    except Exception as e:
                        st.error(f"匯入失敗，請確認檔案格式是否與範例一致。\n錯誤訊息: {e}")

    # 主畫面：資料編輯器
    data = load_data()
    st.subheader("📝 每月營收明細編輯")
    edited_df = st.data_editor(
        data,
        num_rows="dynamic",
        use_container_width=True,
        height=550,
        column_config={
            "Jan": st.column_config.NumberColumn(format="%.0f"),
            "Feb": st.column_config.NumberColumn(format="%.0f"),
            # ... 可依需求增加各月份格式
        }
    )

    # 操作按鈕區
    col1, col2, _ = st.columns([1, 1, 4])
    with col1:
        if st.button("💾 儲存所有變更", type="primary"):
            with st.spinner("正在同步至雲端..."):
                save_to_supabase(edited_df)
                st.success("雲端資料庫已更新！")
    
    with col2:
        # 匯出功能
        csv_data = edited_df.to_csv(index=False).encode('utf-8-sig')
        st.download_button(
            label="📥 匯出當前報表 (CSV)",
            data=csv_data,
            file_name="營收管理系統匯出.csv",
            mime="text/csv"
        )
