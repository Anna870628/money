import streamlit as st
import pandas as pd
from sqlalchemy import text

# 1. 頁面配置
st.set_page_config(page_title="車聯網專案營收管理", layout="wide")

# --- CSS 視覺強化設計 ---
st.markdown("""
    <style>
    /* 全域背景色 */
    .stApp { background-color: #f4f7f9; }
    
    /* 強制側邊欄顏色：深色背景，亮色文字 */
    [data-testid="stSidebar"] {
        background-color: #1e2124 !important;
        color: #ffffff !important;
    }
    [data-testid="stSidebar"] .stMarkdown p, [data-testid="stSidebar"] label, [data-testid="stSidebar"] .stHeader {
        color: #ffffff !important;
    }

    /* 專案卡片樣式：灰色底色以利區分白色背景 */
    .project-card {
        background-color: #e9ecef; 
        padding: 20px;
        border-radius: 12px;
        border: 1px solid #ced4da;
        margin-bottom: 20px;
        box-shadow: 2px 2px 8px rgba(0,0,0,0.05);
    }
    
    /* Metric 數值框樣式優化 */
    div[data-testid="stMetric"] {
        background-color: #ffffff;
        padding: 10px;
        border-radius: 8px;
        border: 1px solid #dee2e6;
    }
    
    .project-title { color: #0d47a1; font-weight: bold; font-size: 1.5rem; }
    .category-label { background-color: #6c757d; color: white; padding: 2px 8px; border-radius: 4px; font-size: 0.8rem; }
    </style>
""", unsafe_allow_html=True)

# 建立資料庫連線
try:
    conn = st.connection("postgresql", type="sql")
except Exception as e:
    st.error("❌ 資料庫連線失敗")
    st.stop()

# --- 核心工具：資料清洗 ---
def clean_num(val):
    v = pd.to_numeric(val, errors='coerce')
    return float(v) if not pd.isna(v) else 0.0

def get_db_data():
    df = conn.query('SELECT * FROM financials ORDER BY id ASC', ttl=0)
    if not df.empty:
        for col in df.select_dtypes(['object']).columns:
            df[col] = df[col].astype(str).str.strip()
    return df

# ==========================================
# ⬅️ 左側邊欄：控制台
# ==========================================
with st.sidebar:
    st.title("📂 營收數據中心")
    
    uploaded_file = st.file_uploader("匯入 2026 專案 Excel", type=["xlsx"])
    if uploaded_file:
        try:
            # 偵測標題列並讀取
            temp_df = pd.read_excel(uploaded_file, nrows=10)
            header_row = 0
            for i, row in temp_df.iterrows():
                if "專案說明" in str(row.values):
                    header_row = i + 1
                    break
            
            new_df = pd.read_excel(uploaded_file, header=header_row)
            new_df.columns = [str(c).strip() for c in new_df.columns]
            
            # 處理無標題的「紀錄類型」
            if "專案說明" in new_df.columns:
                p_idx = new_df.columns.get_loc("專案說明")
                if p_idx + 1 < len(new_df.columns):
                    new_df = new_df.rename(columns={new_df.columns[p_idx + 1]: "紀錄類型"})
            
            # 合併儲存格填充
            new_df['專案說明'] = new_df['專案說明'].ffill()
            if "營收分類" in new_df.columns:
                new_df['營收分類'] = new_df['營營分類'].ffill()

            if st.button("🚀 確認寫入資料庫", use_container_width=True):
                with conn.session as s:
                    for _, r in new_df.iterrows():
                        p_name = str(r.get('專案說明', ''))
                        if pd.isna(r.get('專案說明')) or "序號" in p_name: continue
                        
                        sql = text("""
                            INSERT INTO financials ("專案說明", "Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec", "營收分類", "紀錄類型", "說明")
                            VALUES (:p, :m1, :m2, :m3, :m4, :m5, :m6, :m7, :m8, :m9, :m10, :m11, :m12, :cat, :type, :desc)
                        """)
                        s.execute(sql, {
                            "p": p_name, "cat": str(r.get('營收分類', '其他')),
                            "type": str(r.get('紀錄類型', '收入')), "desc": str(r.get('說明', '')),
                            "m1": clean_num(r.get('Jan')), "m2": clean_num(r.get('Feb')), "m3": clean_num(r.get('Mar')),
                            "m4": clean_num(r.get('Apr')), "m5": clean_num(r.get('May')), "m6": clean_num(r.get('Jun')),
                            "m7": clean_num(r.get('Jul')), "m8": clean_num(r.get('Aug')), "m9": clean_num(r.get('Sep')),
                            "m10": clean_num(r.get('Oct')), "m11": clean_num(r.get('Nov')), "m12": clean_num(r.get('Dec'))
                        })
                    s.commit()
                st.toast("數據匯入成功！")
                st.rerun()
        except Exception as e:
            st.error(f"解析失敗: {e}")

    st.divider()
    if st.button("🗑️ 清空資料庫", type="primary", use_container_width=True):
        with conn.session as s:
            s.execute(text("TRUNCATE TABLE financials;"))
            s.commit()
        st.rerun()

# ==========================================
# 🏠 主畫面內容
# ==========================================
st.title("📊 車聯網事業本部 - 專案彙整")
df = get_db_data()

if df.empty:
    st.info("👋 目前尚無數據，請先由左側上傳 Excel。")
else:
    # 數值預處理
    months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
    for m in months:
        df[m] = pd.to_numeric(df[m], errors='coerce').fillna(0)
    df['年度總額'] = df[months].sum(axis=1)

    tab1, tab2, tab3 = st.tabs(["🚀 專案推進卡片", "📉 營收分類分析", "📝 原始明細檢視"])

    # --- Tab 1: 各專案推進營收 ---
    with tab1:
        projects = [p for p in df['專案說明'].unique() if p and str(p) != 'None']
        cols = st.columns(2)
        
        for i, project in enumerate(projects):
            p_rows = df[df['專案說明'] == project]
            # 目標 (收入 灰底)
            target_rev = p_rows[p_rows['紀錄類型'] == '收入']['年度總額'].sum()
            # 預估 (收入預估 粉底)
            est_rev = p_rows[p_rows['紀錄類型'].str.contains('預估', na=False) & p_rows['紀錄類型'].str.contains('收入', na=False)]['年度總額'].sum()
            # 推進率
            rate = (est_rev / target_rev) if target_rev > 0 else 0
            
            with cols[i % 2]:
                # 使用自定義 project-card 類別加深底色
                st.markdown(f"""
                <div class="project-card">
                    <div class="project-title">{project}</div>
                    <span class="category-label">{p_rows['營收分類'].iloc[0]}</span>
                    <hr>
                </div>
                """, unsafe_allow_html=True)
                
                m1, m2, m3 = st.columns(3)
                m1.metric("目標營收 (灰)", f"${target_rev:,.0f}")
                m2.metric("預估收入 (粉)", f"${est_rev:,.0f}")
                m3.metric("目標推進率", f"{rate:.1%}")
                st.progress(min(rate, 1.0) if rate >= 0 else 0)

    # --- Tab 2: 營收分類分析表 ---
    with tab2:
        st.subheader("📊 營收分類分析表")
        
        # 執行加總計算 (嚴格執行你的財務公式)
        cat_summary = []
        categories = df['營收分類'].unique()
        
        for cat in categories:
            c_df = df[df['營收分類'] == cat]
            
            target_in = c_df[c_df['紀錄類型'] == '收入']['年度總額'].sum()
            est_in = c_df[c_df['紀錄類型'].str.contains('收入', na=False) & c_df['紀錄類型'].str.contains('預估', na=False)]['年度總額'].sum()
            target_out = c_df[c_df['紀錄類型'] == '支出']['年度總額'].sum()
            est_out = c_df[c_df['紀錄類型'].str.contains('支出', na=False) & c_df['紀錄類型'].str.contains('預估', na=False)]['年度總額'].sum()
            
            # 1. 目標毛利 = 目標收入 - 支出 (白底)
            target_gp = target_in - target_out
            # 2. 預估毛利 = 收入預估 - 支出預估 (粉底)
            est_gp = est_in - est_out
            # 3. 預估毛利率 = 預估毛利 / 預估收入
            est_gp_rate = (est_gp / est_in) if est_in != 0 else 0
            # 4. 差異 = 目標收入 - 預估收入
            diff = target_in - est_in
            
            cat_summary.append({
                "營收分類": cat,
                "目標收入": target_in,
                "預估收入": est_in,
                "目標毛利": target_gp,
                "預估毛利": est_gp,
                "預估毛利率": est_gp_rate,
                "差異 (目標-預估)": diff
            })
        
        report_df = pd.DataFrame(cat_summary)
        
        # 樣式設定：差異大於 0 (表示預估不如目標) 則標紅
        st.dataframe(
            report_df.style.format({
                "目標收入": "${:,.0f}", "預估收入": "${:,.0f}",
                "目標毛利": "${:,.0f}", "預估毛利": "${:,.0f}",
                "預估毛利率": "{:.1%}", "差異 (目標-預估)": "${:,.0f}"
            }).map(lambda x: 'color: red' if x > 0 else 'color: green', subset=['差異 (目標-預估)']),
            use_container_width=True
        )

    # --- Tab 3: 原始數據 ---
    with tab3:
        st.dataframe(df, use_container_width=True)
