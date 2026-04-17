import streamlit as st
import pandas as pd
import numpy as np

# --- 1. 初始化 PostgreSQL 連線 ---
# Streamlit 會自動讀取 Secrets 裡的 [connections.postgresql]
conn = st.connection("postgresql", type="sql")

st.set_page_config(layout="wide", page_title="車聯網營收整理系統")

# --- 2. 資料處理函數 ---
def fetch_data():
    # ttl=0 確保手動更新後能立即看到結果
    return conn.query('SELECT * FROM financials', ttl=0)

def upload_to_db(df):
    # 移除自動生成的系統欄位再寫入
    df_to_save = df.drop(columns=['id', '建立時間'], errors='ignore')
    with conn.session as session:
        # 先清空舊資料再寫入新資料（批次更新邏輯）
        session.execute("DELETE FROM financials")
        df_to_save.to_sql(
            "financials", 
            con=conn.engine, 
            if_exists="append", 
            index=False,
            method="multi"
        )
        session.commit()

# --- 3. 側邊欄：Excel 解析邏輯 ---
with st.sidebar:
    st.title("📂 數據上傳")
    uploaded_file = st.file_uploader("匯入 2026 專案收支 Excel", type=["xlsx"])
    
    if uploaded_file:
        if st.button("🚀 執行解析並更新資料庫"):
            # 讀取 Excel
            raw_df = pd.read_excel(uploaded_file, header=None)
            total_cols = raw_df.shape[1]
            records = []
            
            # 每 6 列為一組專案，從第 3 列 (index 2) 開始
            for i in range(2, len(raw_df), 6):
                # 專案說明在 index 1
                proj_name = raw_df.iloc[i, 1] if total_cols > 1 and pd.notna(raw_df.iloc[i, 1]) else None
                if not proj_name: continue

                # 營收分類在 index 19，說明在 index 20
                cat = raw_df.iloc[i, 19] if total_cols > 19 and pd.notna(raw_df.iloc[i, 19]) else "未分類"
                desc = raw_df.iloc[i, 20] if total_cols > 20 and pd.notna(raw_df.iloc[i, 20]) else ""

                # 固定六列標籤
                row_labels = ["收入", "收入預估", "支出", "支出預估", "收入差異", "支出差異"]
                
                for idx, label in enumerate(row_labels):
                    if i + idx < len(raw_df):
                        # 月份數據在 index 3~14 (D到O欄)
                        m_vals = raw_df.iloc[i + idx, 3:15].fillna(0).replace('-', 0).astype(float).tolist()
                        
                        # 確保月份數據長度正確
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
                st.success("✅ 資料庫更新成功！")
                st.rerun()

# --- 4. 主分頁 UI 邏輯 ---
tab1, tab2, tab3 = st.tabs(["📊 各專案推進營收", "🛠️ 資料庫管理", "📈 營收分類彙整表"])

df = fetch_data()
months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]

if not df.empty:
    df['年度小計'] = df[months].sum(axis=1)

    with tab1:
        st.subheader("各專案進度卡片")
        projects = df['專案說明'].unique()
        cols = st.columns(3)
        for idx, p in enumerate(projects):
            p_df = df[df['專案說明'] == p]
            
            target_rev = p_df[p_df['紀錄類型'] == "收入"]['年度小計'].sum()
            est_rev = p_df[p_df['紀錄類型'] == "收入預估"]['年度小計'].sum()
            
            # 修正先前的打字錯誤
            category = p_df['營收分類'].iloc[0] if '營收分類' in p_df.columns else "未分類"
            
            rate = (est_rev / target_rev * 100) if target_rev != 0 else 0
            
            with cols[idx % 3]:
                st.markdown(f"""
                <div style="padding:15px; border-radius:10px; border:1px solid #eee; margin-bottom:15px; background-color: #fcfcfc;">
                    <small style="color:gray;">{category}</small>
                    <h4 style="margin:5px 0;">{p}</h4>
                    <p style="margin:2px; font-size:13px;">目標營收：{target_rev:,.0f}</p>
                    <p style="margin:2px; font-size:13px;">預估收入：{est_rev:,.0f}</p>
                    <h3 style="color:{'#d32f2f' if rate < 100 else '#2e7d32'}; margin-top:8px;">{rate:.1f}% <small style="font-size:12px;">推進率</small></h3>
                </div>
                """, unsafe_allow_html=True)

    with tab2:
        st.subheader("手動更新與批次管理")
        edited_df = st.data_editor(
            df, 
            num_rows="dynamic", 
            key="db_editor",
            column_config={"id": None, "建立時間": None}
        )
        
        if st.button("💾 儲存所有變更"):
            upload_to_db(edited_df)
            st.success("變更已同步至資料庫")
            st.rerun()

    with tab3:
        st.subheader("營收分類彙整報表")
        # 依分類加總
        agg = df.groupby(['營收分類', '紀錄類型'])['年度小計'].sum().unstack(fill_value=0)
        
        summary = pd.DataFrame(index=agg.index)
        summary['目標收入'] = agg.get('收入', 0)
        summary['預估收入'] = agg.get('收入預估', 0)
        summary['目標毛利'] = agg.get('收入', 0) - agg.get('支出', 0)
        summary['預估毛利'] = agg.get('收入預估', 0) - agg.get('支出預估', 0)
        summary['預估毛利率'] = (summary['預估毛利'] / summary['預估收入']).replace([np.inf, -np.inf], 0).fillna(0)
        summary['差異(目標-預估)'] = summary['目標收入'] - summary['預估收入']
        
        st.dataframe(
            summary.style.format({
                '目標收入': '{:,.0f}', '預估收入': '{:,.0f}',
                '目標毛利': '{:,.0f}', '預估毛利': '{:,.0f}',
                '預估毛利率': '{:.2%}', '差異(目標-預估)': '{:,.0f}'
            }).applymap(lambda x: 'color: red' if isinstance(x, (int, float)) and x < 0 else '', 
                       subset=['差異(目標-預估)', '預估毛利'])
        )
else:
    st.info("尚未匯入數據，請利用左側面板上傳 Excel。")
