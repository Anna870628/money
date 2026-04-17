import streamlit as st
import pandas as pd
import numpy as np

# --- 1. 初始化 PostgreSQL 連線 ---
# 使用 st.connection 自動從 Secrets 讀取 [connections.postgresql]
conn = st.connection("postgresql", type="sql")

st.set_page_config(layout="wide", page_title="營收整理系統")

# --- 2. 資料存取函數 ---
def fetch_data():
    # ttl=0 確保每次切換分頁或重新整理都抓取最新資料
    return conn.query('SELECT * FROM financials', ttl=0)

def upload_to_db(df):
    # 移除自動生成的 id 或建立時間再寫入
    df_to_save = df.drop(columns=['id', '建立時間'], errors='ignore')
    with conn.session as session:
        # 依需求決定是否先清空舊資料
        session.execute("DELETE FROM financials")
        df_to_save.to_sql(
            "financials", 
            con=conn.engine, 
            if_exists="append", 
            index=False,
            method="multi"
        )
        session.commit()

# --- 3. 側邊欄：檔案匯入與解析邏輯 ---
with st.sidebar:
    st.title("📂 數據上傳中心")
    uploaded_file = st.file_uploader("匯入 2026 專案收支 Excel", type=["xlsx"])
    
    if uploaded_file:
        if st.button("🚀 確認解析並覆蓋資料庫"):
            # 讀取 Excel，不設定標題，由程式邏輯解析
            raw_df = pd.read_excel(uploaded_file, header=None)
            total_cols = raw_df.shape[1]
            records = []
            
            # 解析邏輯：從第 3 列 (index 2) 開始，每 6 列為一個專案組
            for i in range(2, len(raw_df), 6):
                # 安全讀取專案名稱 (索引 1)
                proj_name = raw_df.iloc[i, 1] if total_cols > 1 and pd.notna(raw_df.iloc[i, 1]) else None
                if not proj_name: continue

                # 安全讀取分類 (索引 19) 與 說明 (索引 20)
                cat = raw_df.iloc[i, 19] if total_cols > 19 and pd.notna(raw_df.iloc[i, 19]) else "未分類"
                desc = raw_df.iloc[i, 20] if total_cols > 20 and pd.notna(raw_df.iloc[i, 20]) else ""

                # 固定六列類型：收入(灰)、收入預估(粉)、支出(白)、支出預估(粉)、收入差異(藍)、支出差異(藍)
                row_labels = ["收入", "收入預估", "支出", "支出預估", "收入差異", "支出差異"]
                
                for idx, label in enumerate(row_labels):
                    if i + idx < len(raw_df):
                        # 取得 1~12 月數據 (欄位 D 到 O，索引 3 到 14)
                        # 使用填補 0 與資料轉換確保不會出錯
                        m_vals = raw_df.iloc[i + idx, 3:15].fillna(0).replace('-', 0).astype(float).tolist()
                        
                        # 若欄位不足 12 個月，補齊至 12 個
                        while len(m_vals) < 12: m_vals.append(0.0)
                        
                        records.append({
                            "專案說明": proj_name,
                            "Jan": m_vals[0], "Feb": m_vals[1], "Mar": m_vals[2], 
                            "Apr": m_vals[3], "May": m_vals[4], "Jun": m_vals[5],
                            "Jul": m_vals[6], "Aug": m_vals[7], "Sep": m_vals[8], 
                            "Oct": m_vals[9], "Nov": m_vals[10], "Dec": m_vals[11],
                            "營收分類": cat,
                            "紀錄類型": label,
                            "說明": desc
                        })
            
            if records:
                upload_to_db(pd.DataFrame(records))
                st.success("✅ 資料解析成功並已更新資料庫！")
                st.rerun()

# --- 4. 主要分頁 UI ---
tab1, tab2, tab3 = st.tabs(["📊 各專案推進營收", "🛠️ 資料庫管理", "📈 營收分類彙整表"])

df = fetch_data()

# 定義月份欄位清單
months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]

if not df.empty:
    # 預先計算每行的年度總計
    df['年度小計'] = df[months].sum(axis=1)

    with tab1:
        st.subheader("專案目標達成狀況")
        projects = df['專案說明'].unique()
        cols = st.columns(3)
        for idx, p in enumerate(projects):
            p_df = df[df['專案說明'] == p]
            
            # 目標營收 = 紀錄類型為 '收入' 的年度小計
            target_rev = p_df[p_df['紀錄類型'] == "收入"]['年度小計'].sum()
            # 預估收入 = 紀錄類型為 '收入預估' 的年度小計
            est_rev = p_df[p_df['紀錄類型'] == "收入預估"]['年度小計'].sum()
            category = p_df['營營收分類'].iloc[0] if '營收分類' in p_df.columns else "未分類"
            
            rate = (est_rev / target_rev * 100) if target_rev != 0 else 0
            
            with cols[idx % 3]:
                st.markdown(f"""
                <div style="padding:15px; border-radius:10px; border:1px solid #eee; margin-bottom:15px;">
                    <small style="color:gray;">{category}</small>
                    <h4 style="margin:5px 0;">{p}</h4>
                    <p style="margin:2px; font-size:14px;">目標：{target_rev:,.0f}</p>
                    <p style="margin:2px; font-size:14px;">預估：{est_rev:,.0f}</p>
                    <h3 style="color:{'#d32f2f' if rate < 100 else '#2e7d32'}; margin-top:10px;">{rate:.1f}% 達成</h3>
                </div>
                """, unsafe_allow_html=True)

    with tab2:
        st.subheader("資料庫原始明細")
        # 標示紅字邏輯 (小於 0 為紅)
        def style_negative(val):
            color = 'red' if isinstance(val, (int, float)) and val < 0 else 'black'
            return f'color: {color}'

        edited_df = st.data_editor(
            df, 
            num_rows="dynamic", 
            key="db_editor",
            column_config={"id": None, "建立時間": None} # 隱藏系統欄位
        )
        
        if st.button("💾 儲存所有變更"):
            upload_to_db(edited_df)
            st.success("變更已同步至 Supabase！")
            st.rerun()
        st.caption("提示：可在表格中直接修改數值，或勾選最左側列按 Delete 刪除。")

    with tab3:
        st.subheader("營收分類彙整報表")
        # 依照分類與紀錄類型加總
        agg = df.groupby(['營收分類', '紀錄類型'])['年度小計'].sum().unstack(fill_value=0)
        
        # 建立彙整表並計算指標
        summary = pd.DataFrame(index=agg.index)
        summary['目標收入'] = agg.get('收入', 0)
        summary['預估收入'] = agg.get('收入預估', 0)
        summary['目標毛利'] = agg.get('收入', 0) - agg.get('支出', 0)
        summary['預估毛利'] = agg.get('收入預估', 0) - agg.get('支出預估', 0)
        summary['預估毛利率'] = (summary['預估毛利'] / summary['預估收入']).replace([np.inf, -np.inf], 0).fillna(0)
        summary['差異(目標-預估)'] = summary['目標收入'] - summary['預估收入']
        
        # 套用美化與紅字標示
        st.dataframe(
            summary.style.format({
                '目標收入': '{:,.0f}', '預估收入': '{:,.0f}',
                '目標毛利': '{:,.0f}', '預估毛利': '{:,.0f}',
                '預估毛利率': '{:.2%}', '差異(目標-預估)': '{:,.0f}'
            }).applymap(lambda x: 'color: red' if isinstance(x, (int, float)) and x < 0 else '', 
                       subset=['差異(目標-預估)', '預估毛利'])
        )
else:
    st.info("👋 歡迎！請先從左側邊欄匯入專案收支 Excel 檔案以開始分析。")
