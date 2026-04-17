import streamlit as st
import pandas as pd
from sqlalchemy import text

# 1. 頁面配置
st.set_page_config(page_title="車聯網營收戰情室", layout="wide")

# 強制修正 Sidebar 顏色與字體，解決看不清楚的問題
st.markdown("""
    <style>
    /* 側邊欄背景與文字顏色強制設定 */
    [data-testid="stSidebar"] {
        background-color: #262730 !important; /* 深灰色背景 */
        color: white !important;
    }
    [data-testid="stSidebar"] .stMarkdown p, [data-testid="stSidebar"] label {
        color: white !important;
    }
    /* 主畫面 Metric 樣式 */
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
    st.error("❌ 連線失敗")
    st.stop()

# --- 核心函數 ---
def get_clean_data():
    df = conn.query('SELECT * FROM financials ORDER BY id ASC', ttl=0)
    if not df.empty:
        for col in df.select_dtypes(['object']).columns:
            df[col] = df[col].astype(str).str.strip()
    return df

def calculate_yearly_sum(df):
    months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
    temp_df = df.copy()
    for m in months:
        temp_df[m] = pd.to_numeric(temp_df[m], errors='coerce').fillna(0)
    return temp_df[months].sum(axis=1)

# ==========================================
# ⬅️ 左側邊欄：功能控制區
# ==========================================
with st.sidebar:
    st.title("📂 數據中心")
    
    st.subheader("1. 匯入 Excel 報表")
    uploaded_file = st.file_uploader("選擇 2026 專案收支 Excel", type=["xlsx"])
    
    if uploaded_file:
        try:
            # 💡 關鍵修正：針對你的檔案，標題通常在第 3 列 (index 2)
            # 我們先讀前幾行來判斷標題在哪裡
            temp_df = pd.read_excel(uploaded_file, nrows=10)
            header_row = 0
            for i, row in temp_df.iterrows():
                if "專案說明" in str(row.values):
                    header_row = i + 1 # 找到「專案說明」那一列作為標題
                    break
            
            # 重新以正確的標題列讀取
            new_df = pd.read_excel(uploaded_file, header=header_row)
            new_df.columns = [str(c).strip() for c in new_df.columns]
            
            # 處理「紀錄類型」那一欄（它通常在專案說明右邊，且沒有名字）
            if "專案說明" in new_df.columns:
                proj_idx = new_df.columns.get_loc("專案說明")
                # 假設右邊那一欄就是紀錄類型（收入/支出）
                type_col_name = new_df.columns[proj_idx + 1]
                new_df = new_df.rename(columns={type_col_name: "紀錄類型"})
            
            # 處理合併儲存格：向下填充
            new_df['專案說明'] = new_df['專案說明'].ffill()
            if "營收分類" in new_df.columns:
                new_df['營收分類'] = new_df['營收分類'].ffill()
            
            st.success("✅ 檔案讀取成功")
            st.write("資料預覽：", new_df[['專案說明', '紀錄類型']].head(6))
            
            if st.button("🚀 確認更新資料庫", use_container_width=True):
                with conn.session as s:
                    for _, r in new_df.iterrows():
                        if pd.isna(r.get('專案說明')) or "序號" in str(r.get('專案說明')): continue
                        
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
                        s.execute(sql, {
                            "p": str(r.get('專案說明')), "cat": str(r.get('營收分類', '其他')),
                            "type": str(r.get('紀錄類型', '未知')), "desc": str(r.get('說明', '')),
                            "m1": r.get('Jan', 0), "m2": r.get('Feb', 0), "m3": r.get('Mar', 0),
                            "m4": r.get('Apr', 0), "m5": r.get('May', 0), "m6": r.get('Jun', 0),
                            "m7": r.get('Jul', 0), "m8": r.get('Aug', 0), "m9": r.get('Sep', 0),
                            "m10": r.get('Oct', 0), "m11": r.get('Nov', 0), "m12": r.get('Dec', 0)
                        })
                    s.commit()
                st.toast("數據已更新！")
                st.rerun()
        except Exception as e:
            st.error(f"解析失敗: {e}")

    st.divider()
    st.subheader("2. 管理操作")
    if st.button("🗑️ 刪除全部資料", type="primary", use_container_width=True):
        with conn.session as s:
            s.execute(text("TRUNCATE TABLE financials;"))
            s.commit()
        st.rerun()

# ==========================================
# 🏠 主畫面內容
# ==========================================
st.title("📊 車聯網事業本部 - 營收戰情室")
df = get_clean_data()

if df.empty:
    st.info("👋 你好！請由左側側邊欄上傳 Excel 檔案開始分析。")
else:
    df['年度總額'] = calculate_yearly_sum(df)
    tab1, tab2, tab3 = st.tabs(["🚀 專案看板", "📈 分類彙整", "📝 明細檢視"])

    with tab1:
        st.markdown("<h3 class='project-header'>各專案目標推進率</h3>", unsafe_allow_html=True)
        projects = [p for p in df['專案說明'].unique() if p and str(p).lower() != 'nan']
        
        cols = st.columns(2)
        for i, project in enumerate(projects):
            p_rows = df[df['專案說明'] == project]
            
            # 目標 (收入) vs 預估 (含有預估字眼的收入)
            target = p_rows[p_rows['紀錄類型'] == '收入']['年度總額'].sum()
            # 根據你提供的資料，有的紀錄類型叫「收入」，但顏色標註不同，
            # 這裡我們預設 DB 裡會區分「收入」與「收入預估」
            est = p_rows[p_rows['紀錄類型'].str.contains('預估', na=False) & p_rows['紀錄類型'].str.contains('收入', na=False)]['年度總額'].sum()
            
            # 如果資料庫裡兩個都叫「收入」，我們就取第一筆當目標，第二筆當預估（針對你的特殊格式）
            if est == 0 and len(p_rows[p_rows['紀錄類型'] == '收入']) > 1:
                rev_rows = p_rows[p_rows['紀錄類型'] == '收入']
                target = rev_rows.iloc[0]['年度總額']
                est = rev_rows.iloc[1]['年度總額']

            rate = (est / target) if target > 0 else 0
            
            with cols[i % 2]:
                with st.container(border=True):
                    c1, c2 = st.columns([3, 1])
                    c1.subheader(project)
                    c2.markdown(f"## {rate:.0%}")
                    
                    m1, m2, m3 = st.columns(3)
                    m1.metric("目標 (灰底)", f"${target:,.0f}")
                    m2.metric("預估 (粉底)", f"${est:,.0f}")
                    m3.metric("差異", f"${est-target:,.0f}")
                    st.progress(min(rate, 1.0) if rate >= 0 else 0)

    with tab2:
        st.markdown("<h3 class='project-header'>分類彙整分析表</h3>", unsafe_allow_html=True)
        # 這裡會根據你的營收分類進行加總
        pivot = df.pivot_table(index='營收分類', columns='紀錄類型', values='年度總額', aggfunc='sum').fillna(0)
        st.dataframe(pivot.style.format("{:,.0f}"), use_container_width=True)

    with tab3:
        st.dataframe(df, use_container_width=True)
