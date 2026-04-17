import streamlit as st
import pandas as pd
from sqlalchemy import text
import numpy as np

# 1. 頁面配置與深色系視覺優化
st.set_page_config(page_title="車聯網營收戰情室", layout="wide")

# 強制自定義 CSS：解決看板底色太白、側邊欄看不清的問題
st.markdown("""
    <style>
    /* 全域背景稍微帶一點灰，增加對比 */
    .stApp {
        background-color: #f4f7f9;
    }
    /* 側邊欄強制深色 */
    [data-testid="stSidebar"] {
        background-color: #1e1e2d !important;
        color: white !important;
    }
    [data-testid="stSidebar"] .stMarkdown p, [data-testid="stSidebar"] label {
        color: #cfcfd8 !important;
    }
    /* 看板卡片樣式：深藍色邊框與陰影，確保白色底色也看得很清楚 */
    .project-card {
        background-color: #ffffff;
        padding: 20px;
        border-radius: 12px;
        border-left: 8px solid #1f4e79;
        box-shadow: 0 4px 6px rgba(0,0,0,0.1);
        margin-bottom: 20px;
    }
    .stMetric {
        background-color: #f8f9fa;
        border: 1px solid #eee;
        border-radius: 8px;
        padding: 10px;
    }
    </style>
""", unsafe_allow_html=True)

# 建立資料庫連線
try:
    conn = st.connection("postgresql", type="sql")
except Exception as e:
    st.error("❌ 資料庫連線失敗，請檢查 Secrets 設定。")
    st.stop()

# --- 核心工具：數據清理 ---
def clean_num(val):
    """將 Excel 中的 '-' 或非數字字元轉為 0.0"""
    v = pd.to_numeric(val, errors='coerce')
    return float(v) if not pd.isna(v) else 0.0

def get_db_data():
    """抓取資料並清理字串空白"""
    df = conn.query('SELECT * FROM financials ORDER BY id ASC', ttl=0)
    if not df.empty:
        for col in df.select_dtypes(['object']).columns:
            df[col] = df[col].astype(str).str.strip()
    return df

# ==========================================
# ⬅️ 側邊欄：匯入與管理
# ==========================================
with st.sidebar:
    st.title("🗄️ 營收管理後台")
    
    st.subheader("📥 匯入 Excel 報表")
    uploaded_file = st.file_uploader("選擇 2026 專案 Excel", type=["xlsx"])
    
    if uploaded_file:
        try:
            # 💡 根據你的 Excel 結構：標題通常在第 3 列 (Index 2)
            # 自動偵測「專案說明」所在行
            temp_df = pd.read_excel(uploaded_file, nrows=15)
            header_idx = 0
            for i, row in temp_df.iterrows():
                if "專案說明" in [str(v).strip() for v in row.values]:
                    header_idx = i + 1
                    break
            
            # 正式讀取
            raw_df = pd.read_excel(uploaded_file, header=header_idx)
            raw_df.columns = [str(c).strip() for c in raw_df.columns]
            
            # 處理「紀錄類型」 (專案說明右邊那欄)
            if "專案說明" in raw_df.columns:
                p_idx = raw_df.columns.get_loc("專案說明")
                type_col = raw_df.columns[p_idx + 1]
                raw_df = raw_df.rename(columns={type_col: "紀錄類型"})
            
            # 處理合併儲存格：專案名稱與分類向下填充
            raw_df['專案說明'] = raw_df['專案說明'].ffill()
            if "營收分類" in raw_df.columns:
                raw_df['營收分類'] = raw_df['營收分類'].ffill()

            st.success(f"✅ 檔案讀取成功 (Header 行數: {header_idx})")
            
            if st.button("🚀 確認並覆蓋資料庫", use_container_width=True):
                with conn.session as s:
                    s.execute(text("TRUNCATE TABLE financials;")) # 每次匯入先清空，維持資料一致
                    for _, r in raw_df.iterrows():
                        p_name = str(r.get('專案說明', ''))
                        if "序號" in p_name or pd.isna(r.get('專案說明')): continue
                        
                        sql = text("""
                            INSERT INTO financials ("專案說明", "Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec", "營收分類", "紀錄類型", "說明")
                            VALUES (:p, :m1, :m2, :m3, :m4, :m5, :m6, :m7, :m8, :m9, :m10, :m11, :m12, :cat, :type, :desc)
                        """)
                        s.execute(sql, {
                            "p": p_name, "cat": str(r.get('營收分類', '其他')),
                            "type": str(r.get('紀錄類型', '未知')), "desc": str(r.get('說明', '')),
                            "m1": clean_num(r.get('Jan')), "m2": clean_num(r.get('Feb')), 
                            "m3": clean_num(r.get('Mar')), "m4": clean_num(r.get('Apr')), 
                            "m5": clean_num(r.get('May')), "m6": clean_num(r.get('Jun')), 
                            "m7": clean_num(r.get('Jul')), "m8": clean_num(r.get('Aug')), 
                            "m9": clean_num(r.get('Sep')), "m10": clean_num(r.get('Oct')), 
                            "m11": clean_num(r.get('Nov')), "m12": clean_num(r.get('Dec'))
                        })
                    s.commit()
                st.rerun()
        except Exception as e:
            st.error(f"解析失敗: {e}")

    st.divider()
    if st.button("🗑️ 清空所有資料庫內容", type="primary", use_container_width=True):
        with conn.session as s:
            s.execute(text("TRUNCATE TABLE financials;"))
            s.commit()
        st.rerun()

# ==========================================
# 🏠 主畫面：看板與彙整分析
# ==========================================
st.title("📊 車聯網專案營收戰情室")
df = get_db_data()

if df.empty:
    st.info("👋 目前尚無數據，請先從左側匯入 Excel。")
else:
    # 預運算月份數值
    months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
    for m in months:
        df[m] = pd.to_numeric(df[m], errors='coerce').fillna(0)
    df['年度總額'] = df[months].sum(axis=1)

    tab1, tab2, tab3 = st.tabs(["🚀 專案推進看板", "📈 營收分類彙整", "📝 原始數據細目"])

    # --- Tab 1: 專案推進看板 (CSS 加強版) ---
    with tab1:
        projects = [p for p in df['專案說明'].unique() if p and "序號" not in str(p)]
        cols = st.columns(2)
        
        for i, project in enumerate(projects):
            p_data = df[df['專案說明'] == project]
            # 邏輯：同專案下第一筆「收入」為目標，第二筆為預估 (對應你的 Excel 順序)
            income_rows = p_data[p_data['紀錄類型'].str.contains('收入', na=False)]
            
            target_rev = income_rows.iloc[0]['年度總額'] if len(income_rows) > 0 else 0
            est_rev = income_rows.iloc[1]['年度總額'] if len(income_rows) > 1 else target_rev
            
            rate = (est_rev / target_rev) if target_rev > 0 else 0
            
            with cols[i % 2]:
                st.markdown(f"""
                <div class="project-card">
                    <h3 style='margin-top:0;'>{project}</h3>
                    <p style='color:gray;'>分類：{p_data['營收分類'].iloc[0] if not p_data.empty else 'N/A'}</p>
                </div>
                """, unsafe_allow_html=True)
                
                m1, m2, m3 = st.columns(3)
                m1.metric("目標營收 (灰底)", f"${target_rev:,.0f}")
                m2.metric("預估收入 (粉底)", f"${est_rev:,.0f}")
                m3.metric("達成率", f"{rate:.1%}")
                st.progress(min(rate, 1.0) if rate >= 0 else 0)
                st.write("") # 間距

    # --- Tab 2: 營收分類彙整表 (精確毛利計算) ---
    with tab2:
        st.subheader("各類別營收毛利彙整")
        
        summary_list = []
        for cat in df['營收分類'].unique():
            cat_df = df[df['營營分類'] == cat]
            
            # 初始化各項指標
            target_rev_sum = 0
            est_rev_sum = 0
            target_exp_sum = 0
            est_exp_sum = 0
            
            # 按專案分組計算，以對應「第1筆是目標、第2筆是預估」的結構
            for proj in cat_df['專案說明'].unique():
                proj_df = cat_df[cat_df['專案說明'] == proj]
                
                # 收入類
                incomes = proj_df[proj_df['紀錄類型'].str.contains('收入', na=False)]
                target_rev_sum += incomes.iloc[0]['年度總額'] if len(incomes) > 0 else 0
                est_rev_sum += incomes.iloc[1]['年度總額'] if len(incomes) > 1 else (incomes.iloc[0]['年度總額'] if len(incomes) > 0 else 0)
                
                # 支出類
                exps = proj_df[proj_df['紀錄類型'].str.contains('支出', na=False)]
                target_exp_sum += exps.iloc[0]['年度總額'] if len(exps) > 0 else 0
                est_exp_sum += exps.iloc[1]['年度總額'] if len(exps) > 1 else (exps.iloc[0]['年度總額'] if len(exps) > 0 else 0)
            
            # 計算財務指標
            target_gp = target_rev_sum - target_exp_sum
            est_gp = est_rev_sum - est_exp_sum
            est_margin = (est_gp / est_rev_sum) if est_rev_sum != 0 else 0
            diff = target_rev_sum - est_rev_sum
            
            summary_list.append({
                "營收分類": cat,
                "目標收入": target_rev_sum,
                "預估收入": est_rev_sum,
                "目標毛利": target_gp,
                "預估毛利": est_gp,
                "預估毛利率": est_margin,
                "差異(目標-預估)": diff
            })
        
        report_df = pd.DataFrame(summary_list)
        
        # 格式化呈現
        st.dataframe(
            report_df.style.format({
                "目標收入": "{:,.0f}",
                "預估收入": "{:,.0f}",
                "目標毛利": "{:,.0f}",
                "預估毛利": "{:,.0f}",
                "預估毛利率": "{:.2%}",
                "差異(目標-預估)": "{:,.0f}"
            }).background_gradient(cmap="RdYlGn_r", subset=["差異(目標-預估)"]),
            use_container_width=True
        )

    with tab3:
        st.dataframe(df, use_container_width=True)
