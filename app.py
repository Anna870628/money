import streamlit as st
import pandas as pd
from sqlalchemy import text

# 1. 頁面配置
st.set_page_config(page_title="車聯網營收管理系統", layout="wide")

# 強制修正 Sidebar 視覺問題，確保文字與背景對比清晰
st.markdown("""
    <style>
    /* 側邊欄背景與文字顏色 */
    [data-testid="stSidebar"] {
        background-color: #262730 !important;
        color: white !important;
    }
    [data-testid="stSidebar"] .stMarkdown p, 
    [data-testid="stSidebar"] label, 
    [data-testid="stSidebar"] .stSubheader {
        color: white !important;
    }
    /* 修改按鈕在側邊欄的樣式 */
    [data-testid="stSidebar"] .stButton button {
        background-color: #4CAF50;
        color: white;
        border-radius: 5px;
    }
    /* 主畫面數值卡片樣式 */
    .stMetric {
        background-color: #ffffff;
        padding: 15px;
        border-radius: 10px;
        box-shadow: 0 2px 5px rgba(0,0,0,0.1);
    }
    .project-header {
        color: #1f4e79;
        border-left: 6px solid #ff4b4b;
        padding-left: 12px;
        margin-bottom: 20px;
        font-weight: bold;
    }
    </style>
""", unsafe_allow_html=True)

# 建立資料庫連線
try:
    conn = st.connection("postgresql", type="sql")
except Exception as e:
    st.error("❌ 資料庫連線失敗")
    st.stop()

# --- 資料清洗工具 ---
def clean_num(val):
    """將 Excel 中的 '-' 或 NaN 轉換為數值 0.0"""
    v = pd.to_numeric(val, errors='coerce')
    return float(v) if not pd.isna(v) else 0.0

def get_clean_db_data():
    """抓取資料庫資料並清理字串空白"""
    df = conn.query('SELECT * FROM financials ORDER BY id ASC', ttl=0)
    if not df.empty:
        for col in df.select_dtypes(['object']).columns:
            df[col] = df[col].astype(str).str.strip()
    return df

# ==========================================
# ⬅️ 左側邊欄：控制台
# ==========================================
with st.sidebar:
    st.title("📂 數據中心")
    
    st.subheader("1. 匯入 Excel 報表")
    uploaded_file = st.file_uploader("選擇 2026 專案收支 Excel", type=["xlsx"])
    
    if uploaded_file:
        try:
            # 讀取 Excel 並找尋標題列
            temp_df = pd.read_excel(uploaded_file, nrows=15)
            header_row = 0
            for i, row in temp_df.iterrows():
                if "專案說明" in str(row.values):
                    header_row = i + 1
                    break
            
            new_df = pd.read_excel(uploaded_file, header=header_row)
            new_df.columns = [str(c).strip() for c in new_df.columns]
            
            # 處理「紀錄類型」欄位 (通常在專案說明右邊那一欄)
            if "專案說明" in new_df.columns:
                proj_idx = new_df.columns.get_loc("專案說明")
                # 取得 Unnamed 欄位並更名為紀錄類型
                if proj_idx + 1 < len(new_df.columns):
                    type_col_name = new_df.columns[proj_idx + 1]
                    new_df = new_df.rename(columns={type_col_name: "紀錄類型"})
            
            # 處理合併儲存格
            new_df['專案說明'] = new_df['專案說明'].ffill()
            if "營收分類" in new_df.columns:
                new_df['營收分類'] = new_df['營收分類'].ffill()
            
            st.success("✅ 檔案讀取成功")
            
            if st.button("🚀 確認更新至資料庫", use_container_width=True):
                with conn.session as s:
                    for _, r in new_df.iterrows():
                        # 過濾掉無意義的列
                        p_name = str(r.get('專案說明', ''))
                        if pd.isna(r.get('專案說明')) or "序號" in p_name or "版本" in p_name: 
                            continue
                        
                        sql = text("""
                            INSERT INTO financials (
                                "專案說明", "Jan", "Feb", "Mar", "Apr", "May", "Jun", 
                                "Jul", "Aug", "Sep", "Oct", "Nov", "Dec", 
                                "營收分類", "紀錄類型", "說明"
                            ) VALUES (
                                :p, :m1, :m2, :m3, :m4, :m5, :m6, 
                                :m7, :m8, :m9, :m10, :m11, :m12, 
                                :cat, :type, :desc
                            )
                        """)
                        
                        # 💡 關鍵修正：將所有月份資料透過 clean_num 轉換，解決 "-" 報錯問題
                        s.execute(sql, {
                            "p": p_name, 
                            "cat": str(r.get('營收分類', '其他')),
                            "type": str(r.get('紀錄類型', '未知')), 
                            "desc": str(r.get('說明', '')),
                            "m1": clean_num(r.get('Jan')), "m2": clean_num(r.get('Feb')), 
                            "m3": clean_num(r.get('Mar')), "m4": clean_num(r.get('Apr')), 
                            "m5": clean_num(r.get('May')), "m6": clean_num(r.get('Jun')), 
                            "m7": clean_num(r.get('Jul')), "m8": clean_num(r.get('Aug')), 
                            "m9": clean_num(r.get('Sep')), "m10": clean_num(r.get('Oct')), 
                            "m11": clean_num(r.get('Nov')), "m12": clean_num(r.get('Dec'))
                        })
                    s.commit()
                st.toast("數據清洗並匯入成功！")
                st.rerun()
        except Exception as e:
            st.error(f"解析失敗: {e}")

    st.divider()
    st.subheader("2. 系統管理")
    if st.button("🗑️ 清空資料庫", type="primary", use_container_width=True):
        with conn.session as s:
            s.execute(text("TRUNCATE TABLE financials;"))
            s.commit()
        st.rerun()

# ==========================================
# 🏠 主畫面內容
# ==========================================
st.title("📊 車聯網事業本部 - 營收管理平台")
df = get_clean_db_data()

if df.empty:
    st.info("👋 你好！目前尚無數據，請由左側上傳專案收支 Excel。")
else:
    # 數值計算
    months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
    for m in months:
        df[m] = pd.to_numeric(df[m], errors='coerce').fillna(0)
    df['年度總額'] = df[months].sum(axis=1)

    tab1, tab2, tab3 = st.tabs(["🚀 專案看板", "📈 分類彙整", "📝 原始明細"])

    with tab1:
        st.markdown("<h3 class='project-header'>各專案目標推進率</h3>", unsafe_allow_html=True)
        # 移除重複與無效專案
        projects = [p for p in df['專案說明'].unique() if p and str(p) != 'None' and '序號' not in str(p)]
        
        cols = st.columns(2)
        for i, project in enumerate(projects):
            p_rows = df[df['專案說明'] == project]
            
            # 定義目標與預估 (根據你的 Excel，通常第一列是目標收入，第二列是預估)
            # 我們嘗試用名稱區分，若名稱相同則按順序取
            income_rows = p_rows[p_rows['紀錄類型'].str.contains('收入', na=False)]
            
            if len(income_rows) >= 2:
                target = income_rows.iloc[0]['年度總額']
                est = income_rows.iloc[1]['年度總額']
            elif len(income_rows) == 1:
                target = income_rows.iloc[0]['年度總額']
                est = 0
            else:
                target, est = 0, 0

            rate = (est / target) if target > 0 else 0
            
            with cols[i % 2]:
                with st.container(border=True):
                    c1, c2 = st.columns([3, 1])
                    c1.subheader(project)
                    c1.caption(f"營收分類：{p_rows['營收分類'].iloc[0] if not p_rows.empty else 'N/A'}")
                    
                    # 達成率標示
                    color = "#28a745" if rate >= 1 else "#ff8c00"
                    c2.markdown(f"<h2 style='text-align:right; color:{color};'>{rate:.0%}</h2>", unsafe_allow_html=True)
                    
                    m1, m2, m3 = st.columns(3)
                    m1.metric("目標 (收入)", f"${target:,.0f}")
                    m2.metric("預估 (預估)", f"${est:,.0f}")
                    diff = est - target
                    m3.metric("差異", f"${diff:,.0f}", delta=diff)
                    st.progress(min(rate, 1.0) if rate >= 0 else 0)

    with tab2:
        st.markdown("<h3 class='project-header'>營收分類分析表</h3>", unsafe_allow_html=True)
        pivot = df.pivot_table(index='營收分類', columns='紀錄類型', values='年度總額', aggfunc='sum').fillna(0)
        st.dataframe(pivot.style.format("{:,.0f}"), use_container_width=True)

    with tab3:
        st.dataframe(df, use_container_width=True)
