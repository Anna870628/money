import streamlit as st
import pandas as pd
from sqlalchemy import text

# 1. 頁面配置 (使用系統預設顏色)
st.set_page_config(page_title="車聯網專案營收戰情室", layout="wide")

# 建立資料庫連線
try:
    conn = st.connection("postgresql", type="sql")
except Exception as e:
    st.error("❌ 資料庫連線失敗，請檢查側邊欄或 Secrets 設定。")
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
    st.title("🗄️ 數據中心")
    
    uploaded_file = st.file_uploader("匯入 2026 專案 Excel", type=["xlsx"])
    
    if uploaded_file:
        try:
            # 自動偵測「專案說明」所在行 (通常是第 3 或 4 行)
            temp_df = pd.read_excel(uploaded_file, nrows=15)
            header_idx = 0
            for i, row in temp_df.iterrows():
                if "專案說明" in [str(v).strip() for v in row.values]:
                    header_idx = i + 1
                    break
            
            raw_df = pd.read_excel(uploaded_file, header=header_idx)
            raw_df.columns = [str(c).strip() for c in raw_df.columns]
            
            # 處理「紀錄類型」 (專案說明右邊那欄)
            if "專案說明" in raw_df.columns:
                p_idx = raw_df.columns.get_loc("專案說明")
                type_col = raw_df.columns[p_idx + 1]
                raw_df = raw_df.rename(columns={type_col: "紀錄類型"})
            
            # 處理合併儲存格
            raw_df['專案說明'] = raw_df['專案說明'].ffill()
            if "營收分類" in raw_df.columns:
                raw_df['營收分類'] = raw_df['營收分類'].ffill()

            st.success("✅ 讀取成功")
            
            if st.button("🚀 更新資料庫", use_container_width=True):
                with conn.session as s:
                    s.execute(text("TRUNCATE TABLE financials;"))
                    for _, r in raw_df.iterrows():
                        p_name = str(r.get('專案說明', ''))
                        if "序號" in p_name or pd.isna(r.get('專案說明')) or p_name == 'nan': 
                            continue
                        
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
    if st.button("🗑️ 清空資料庫", type="primary", use_container_width=True):
        with conn.session as s:
            s.execute(text("TRUNCATE TABLE financials;"))
            s.commit()
        st.rerun()

# ==========================================
# 🏠 主畫面：看板與彙整分析
# ==========================================
st.title("📈 營收管理平台")
df = get_db_data()

if df.empty:
    st.info("👋 目前尚無數據，請先從左側匯入 Excel。")
else:
    # 預運算
    months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
    for m in months:
        df[m] = pd.to_numeric(df[m], errors='coerce').fillna(0)
    df['年度總額'] = df[months].sum(axis=1)

    tab1, tab2, tab3 = st.tabs(["🚀 專案推進看板", "📊 營收分類彙整", "📝 原始數據細目"])

    # --- Tab 1: 專案推進看板 (沿用正確邏輯) ---
    with tab1:
        projects = [p for p in df['專案說明'].unique() if p and "序號" not in str(p) and str(p) != 'nan']
        cols = st.columns(3) # 改回 3 欄
        
        for i, project in enumerate(projects):
            p_data = df[df['專案說明'] == project]
            income_rows = p_data[p_data['紀錄類型'].str.contains('收入', na=False)]
            
            # 卡片邏輯：第1筆收入是目標(灰底)，第2筆收入是預估(粉底)
            target_rev = income_rows.iloc[0]['年度總額'] if len(income_rows) > 0 else 0
            est_rev = income_rows.iloc[1]['年度總額'] if len(income_rows) > 1 else target_rev
            
            rate = (est_rev / target_rev) if target_rev > 0 else 0
            
            with cols[i % 3]:
                with st.container(border=True):
                    st.subheader(project)
                    st.caption(f"分類：{p_data['營收分類'].iloc[0]}")
                    m1, m2 = st.columns(2)
                    m1.metric("目標營收", f"${target_rev:,.0f}")
                    m2.metric("預估收入", f"${est_rev:,.0f}")
                    st.progress(min(rate, 1.0) if rate >= 0 else 0)
                    st.write(f"達成率：{rate:.1%}")

    # --- Tab 2: 營收分類彙整表 (修正 Typo 與 邏輯同步) ---
    with tab2:
        summary_list = []
        unique_cats = [c for c in df['營收分類'].unique() if c and str(c) != 'nan']
        
        for cat in unique_cats:
            # 修正之前的 typo: 營營分類 -> 營收分類
            cat_df = df[df['營收分類'] == cat]
            
            cat_target_rev = 0
            cat_est_rev = 0
            cat_target_exp = 0
            cat_est_exp = 0
            
            # 按專案細分計算再加總，確保符合卡片逻辑
            for proj in cat_df['專案說明'].unique():
                proj_df = cat_df[cat_df['專案說明'] == proj]
                
                # 收入類 (第1筆目標, 第2筆預估)
                incs = proj_df[proj_df['紀錄類型'].str.contains('收入', na=False)]
                p_target_rev = incs.iloc[0]['年度總額'] if len(incs) > 0 else 0
                p_est_rev = incs.iloc[1]['年度總額'] if len(incs) > 1 else p_target_rev
                
                # 支出類 (第1筆目標支出[白底], 第2筆預估支出[粉底])
                exps = proj_df[proj_df['紀錄類型'].str.contains('支出', na=False)]
                p_target_exp = exps.iloc[0]['年度總額'] if len(exps) > 0 else 0
                p_est_exp = exps.iloc[1]['年度總額'] if len(exps) > 1 else p_target_exp
                
                cat_target_rev += p_target_rev
                cat_est_rev += p_est_rev
                cat_target_exp += p_target_exp
                cat_est_exp += p_est_exp
            
            # 計算指標 (按照您的要求)
            target_profit = cat_target_rev - cat_target_exp
            est_profit = cat_est_rev - cat_est_exp
            est_margin = (est_profit / cat_est_rev) if cat_est_rev != 0 else 0
            diff = cat_target_rev - cat_est_rev
            
            summary_list.append({
                "營收分類": cat,
                "目標收入": cat_target_rev,
                "預估收入": cat_est_rev,
                "目標毛利": target_profit,
                "預估毛利": est_profit,
                "預估毛利率": est_margin,
                "差異(目標-預估)": diff
            })
        
        if summary_list:
            report_df = pd.DataFrame(summary_list)
            st.dataframe(
                report_df.style.format({
                    "目標收入": "{:,.0f}",
                    "預估收入": "{:,.0f}",
                    "目標毛利": "{:,.0f}",
                    "預估毛利": "{:,.0f}",
                    "預估毛利率": "{:.2%}",
                    "差異(目標-預估)": "{:,.0f}"
                }),
                use_container_width=True
            )
        else:
            st.warning("無分類數據。")

    with tab3:
        st.dataframe(df, use_container_width=True)
