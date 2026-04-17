import streamlit as st
import pandas as pd
from sqlalchemy import text

# 1. 頁面配置
st.set_page_config(page_title="Carmax 營收戰情室", layout="wide")

# 自定義 CSS 優化閱讀性 (PM 視角)
st.markdown("""
    <style>
    .stMetric { background-color: #ffffff; padding: 15px; border-radius: 10px; border: 1px solid #eee; box-shadow: 0 2px 4px rgba(0,0,0,0.05); }
    [data-testid="stSidebar"] { background-color: #f1f3f6; }
    .project-header { color: #1f4e79; border-left: 5px solid #1f4e79; padding-left: 10px; margin-bottom: 20px; }
    </style>
""", unsafe_allow_html=True)

# 建立資料庫連線
try:
    conn = st.connection("postgresql", type="sql")
except Exception as e:
    st.error("❌ 連線失敗，請檢查 secrets.toml")
    st.stop()

# --- 核心運算函數 ---
def get_clean_data():
    """抓取資料並清理字串空白"""
    df = conn.query('SELECT * FROM financials ORDER BY id ASC', ttl=0)
    if not df.empty:
        for col in df.select_dtypes(['object']).columns:
            df[col] = df[col].astype(str).str.strip()
    return df

def calculate_yearly_sum(df):
    """加總 1-12 月欄位"""
    months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
    temp_df = df.copy()
    for m in months:
        temp_df[m] = pd.to_numeric(temp_df[m], errors='coerce').fillna(0)
    return temp_df[months].sum(axis=1)

# ==========================================
# ⬅️ 左側邊欄：控制中心 (匯入與刪除)
# ==========================================
with st.sidebar:
    st.title("🎛️ 控制後台")
    
    # --- 1. 匯入功能 (處理合併儲存格) ---
    st.subheader("📥 匯入營收 Excel")
    uploaded_file = st.file_uploader("選擇檔案", type=["xlsx"])
    
    if uploaded_file:
        try:
            # 讀取 Excel
            new_df = pd.read_excel(uploaded_file)
            new_df.columns = new_df.columns.str.strip()
            
            # --- 重點：處理合併儲存格 ---
            # 將專案說明、營收分類向下填充，解決合併儲存格產生的 NaN
            if '專案說明' in new_df.columns:
                new_df['專案說明'] = new_df['專案說明'].ffill()
            if '營收分類' in new_df.columns:
                new_df['營收分類'] = new_df['營收分類'].ffill()
            
            st.success("✅ 檔案讀取成功 (已自動修復合併儲存格)")
            st.write("預覽修復後資料：", new_df.head(6))
            
            if st.button("🚀 確認寫入資料庫", use_container_width=True):
                with conn.session as s:
                    for _, r in new_df.iterrows():
                        if pd.isna(r.get('專案說明')): continue # 還是空的就跳過
                        
                        query = text("""
                            INSERT INTO financials ("專案說明", "Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec", "營收分類", "紀錄類型", "說明")
                            VALUES (:p, :m1, :m2, :m3, :m4, :m5, :m6, :m7, :m8, :m9, :m10, :m11, :m12, :cat, :type, :desc)
                        """)
                        s.execute(query, {
                            "p": r.get('專案說明'), 
                            "cat": r.get('營收分類', '未分類'), 
                            "type": r.get('紀錄類型', '未知'), 
                            "desc": r.get('說明',''),
                            "m1":r.get('Jan',0), "m2":r.get('Feb',0), "m3":r.get('Mar',0), "m4":r.get('Apr',0), 
                            "m5":r.get('May',0), "m6":r.get('Jun',0), "m7":r.get('Jul',0), "m8":r.get('Aug',0), 
                            "m
