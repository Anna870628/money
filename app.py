import streamlit as st
import pandas as pd
# 記得在 requirements.txt 中加入 sqlalchemy 和 psycopg2-binary

# ==========================================
# 1. 安全登入驗證 (雲端金鑰版)
# ==========================================
def check_password():
    def password_entered():
        # 從 Streamlit Cloud 的 Secrets 抓取我們設定的 admin_password
        if st.session_state["password"] == st.secrets["passwords"]["admin_password"]:
            st.session_state["password_correct"] = True
            del st.session_state["password"]
        else:
            st.session_state["password_correct"] = False

    if "password_correct" not in st.session_state:
        st.title("🔒 內部營收管理系統 - 登入")
        st.text_input("請輸入存取密碼", type="password", on_change=password_entered, key="password")
        return False
    return st.session_state["password_correct"]

# ==========================================
# 2. 雲端資料庫與資料處理邏輯
# ==========================================
# 只有密碼正確才會執行以下系統內容
if check_password():
    st.set_page_config(page_title="專案營收管理系統", layout="wide")
    
    # 建立 PostgreSQL 連線 (會自動讀取 Secrets 裡的資料庫 URI)
    conn = st.connection("postgresql", type="sql")

    def load_data():
        # ttl="0" 代表不快取，確保每次讀取都是 Supabase 上的最新資料
        return conn.query("SELECT * FROM financials", ttl="0")

    def save_data(df):
        with conn.session as session:
            # 清空舊資料並覆蓋寫入新資料
            session.execute("DELETE FROM financials") 
            df.to_sql('financials', conn.engine, if_exists='append', index=False)
            session.commit()

    def process_imported_csv(uploaded_file):
        df = pd.read_csv(uploaded_file, skiprows=2)
        df.rename(columns={df.columns[2]: '紀錄類型'}, inplace=True)
        df['專案說明'] = df['專案說明'].ffill()
        df['營收分類'] = df['營收分類'].ffill()
        months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']
        df[months] = df[months].fillna(0)
        target_cols = ['專案說明', '紀錄類型'] + months + ['營收分類', '說明']
        return df[target_cols]

    # ==========================================
    # 3. 網頁主介面
    # ==========================================
    st.title("🚀 雲端專案營收管理系統")
    st.info("資料庫狀態：🟢 已連線至 Supabase")

    # 側邊欄：匯入功能與登出
    with st.sidebar:
        if st.button("🚪 安全登出"):
            st.session_state["password_correct"] = False
            st.rerun()
            
        st.divider()
        st.header("📂 資料管理")
        uploaded_file = st.file_uploader("匯入現有 CSV 檔案", type="csv")
        if uploaded_file is not None:
            if st.button("確認匯入並同步至雲端"):
                cleaned_df = process_imported_csv(uploaded_file)
                save_data(cleaned_df)
                st.success("匯入成功！請重新整理頁面。")
                st.rerun()

    # 主編輯區
    try:
        current_data = load_data()
        edited_df = st.data_editor(
            current_data, 
            num_rows="dynamic", 
            use_container_width=True, 
            height=500
        )

        col1, col2, _ = st.columns([1, 1, 4])
        with col1:
            if st.button("💾 同步至雲端資料庫", type="primary"):
                save_data(edited_df)
                st.success("資料已永久儲存至 Supabase！")

        with col2:
            csv_output = edited_df.to_csv(index=False).encode('utf-8-sig')
            st.download_button(label="📥 匯出當前報表", data=csv_output, file_name="雲端營收匯出.csv", mime="text/csv")
            
    except Exception as e:
        st.error(f"⚠️ 資料庫讀取失敗！請確認 Supabase 已經正確建立 `financials` 資料表。詳細錯誤：{e}")
