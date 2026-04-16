import streamlit as st
import pandas as pd
from sqlalchemy import text  # 修正：必須導入 text 函數

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
        st.text_input("請輸入存取密碼 (Password)", type="password", on_change=password_entered, key="password")
        return False
    elif not st.session_state["password_correct"]:
        st.title("🔒 營收管理系統 - 身份驗證")
        st.text_input("請輸入存取密碼 (Password)", type="password", on_change=password_entered, key="password")
        st.error("😕 密碼錯誤，請重新輸入。")
        return False
    return True

if check_password():
    st.set_page_config(page_title="車聯網專案營收管理系統", layout="wide")
    
    # ==========================================
    # 1. 資料庫連線與資料處理邏輯
    # ==========================================
    conn = st.connection("postgresql", type="sql")

    def load_data():
        try:
            return conn.query("SELECT * FROM financials", ttl="0")
        except Exception:
            return pd.DataFrame(columns=['專案說明', '紀錄類型'] + [m for m in ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']] + ['營收分類', '說明'])

    def save_to_supabase(df):
        """修正：使用 text() 包裝 SQL 語句"""
        with conn.session as session:
            # 這裡必須使用 text("...")
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
        
        # 尋找營收分類欄位
        category_col = [c for c in df.columns if '營收分類' in str(c)]
        if category_col:
            df.rename(columns={category_col[0]: '營營分類'}, inplace=True)
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
    # 2. 介面呈現
    # ==========================================
    st.title("📊 車聯網事業本部 - 專案每月收支管理系統")
    st.info("🟢 已成功連線至 Supabase 雲端資料庫")

    with st.sidebar:
        if st.button("🚪 安全登出"):
            st.session_state["password_correct"] = False
            st.rerun()
        
        st.divider()
        st.header("📂 數據匯入")
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
                        st.error(f"匯入失敗。錯誤訊息: {e}")

    data = load_data()
    st.subheader("📝 每月營收明細編輯")
    
    # 顯示編輯器，並確保金額格式正確
    edited_df = st.data_editor(
        data,
        num_rows="dynamic",
        use_container_width=True,
        height=550
    )

    col1, col2, _ = st.columns([1, 1, 4])
    with col1:
        if st.button("💾 儲存所有變更", type="primary"):
            with st.spinner("正在同步至雲端..."):
                try:
                    save_to_supabase(edited_df)
                    st.success("雲端資料庫已更新！")
                except Exception as e:
                    st.error(f"儲存失敗：{e}")
    
    with col2:
        csv_data = edited_df.to_csv(index=False).encode('utf-8-sig')
        st.download_button(
            label="📥 匯出當前報表 (CSV)",
            data=csv_data,
            file_name="營收管理系統匯出.csv",
            mime="text/csv"
        )
