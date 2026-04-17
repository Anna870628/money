import streamlit as st
import pandas as pd
import numpy as np

# --- 1. 初始化連線 ---
conn = st.connection("postgresql", type="sql")

st.set_page_config(layout="wide", page_title="營收管理系統")

# --- 2. 資料處理函數 ---
def fetch_data():
    # 注意：SQL 中的中文欄位建議用雙引號，但 conn.query 處理 DataFrame 時會自動對應
    return conn.query('SELECT * FROM financials', ttl=0)

def upload_to_db(df):
    # 移除自動生成的 id 與時間欄位再寫入
    df_to_save = df.drop(columns=['id', '建立時間'], errors='ignore')
    with conn.session as session:
        session.execute("DELETE FROM financials")
        df_to_save.to_sql(
            "financials", 
            con=conn.engine, 
            if_exists="append", 
            index=False,
            method="multi"
        )
        session.commit()

# --- 3. 側邊欄匯入邏輯 ---
with st.sidebar:
    st.title("數據上傳")
    uploaded_file = st.file_uploader("匯入 Excel", type=["xlsx"])
    if uploaded_file:
        if st.button("確認解析並覆蓋資料庫"):
            raw_df = pd.read_excel(uploaded_file, header=None)
            records = []
            
            # 假設數據從第 3 列開始，每 6 列一組專案
            for i in range(2, len(raw_df), 6):
                proj_name = raw_df.iloc[i, 1] if pd.notna(raw_df.iloc[i, 1]) else None
                cat = raw_df.iloc[i, 19] if pd.notna(raw_df.iloc[i, 19]) else "未分類"
                desc = raw_df.iloc[i, 20] if pd.notna(raw_df.iloc[i, 20]) else ""
                
                if not proj_name: continue
                
                # 紀錄類型對應
                types = ["收入", "收入預估", "支出", "支出預估", "收入差異", "支出差異"]
                
                for idx, t_label in enumerate(types):
                    if i + idx < len(raw_df):
                        # 1-12月數據 (欄位 D 到 O, index 3-14)
                        m_vals = raw_df.iloc[i + idx, 3:15].fillna(0).replace('-', 0).astype(float).tolist()
                        
                        records.append({
                            "專案說明": proj_name,
                            "Jan": m_vals[0], "Feb": m_vals[1], "Mar": m_vals[2], 
                            "Apr": m_vals[3], "May": m_vals[4], "Jun": m_vals[5],
                            "Jul": m_vals[6], "Aug": m_vals[7], "Sep": m_vals[8], 
                            "Oct": m_vals[9], "Nov": m_vals[10], "Dec": m_vals[11],
                            "營收分類": cat,
                            "紀錄類型": t_label,
                            "說明": desc
                        })
            
            if records:
                upload_to_db(pd.DataFrame(records))
                st.success("資料已成功存入 financials 表！")
                st.rerun()

# --- 4. 分頁 UI ---
tab1, tab2, tab3 = st.tabs(["各專案推進營收", "資料庫明細", "營收分類彙整"])
df = fetch_data()

with tab1:
    if not df.empty:
        # 計算年度總額供卡片顯示
        months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
        df['年度小計'] = df[months].sum(axis=1)
        
        projects = df['專案說明'].unique()
        cols = st.columns(3)
        for idx, p in enumerate(projects):
            p_df = df[df['專案說明'] == p]
            target = p_df[p_df['紀錄類型'] == "收入"]['年度小計'].sum()
            est = p_df[p_df['紀錄類型'] == "收入預估"]['年度小計'].sum()
            rate = (est/target*100) if target != 0 else 0
            
            with cols[idx % 3]:
                st.metric(label=p, value=f"{est:,.0f}", delta=f"目標達成率 {rate:.1f}%")

with tab2:
    if not df.empty:
        # 使用 data_editor 手動更新
        st.subheader("手動更正欄位")
        edited_df = st.data_editor(df, num_rows="dynamic", key="main_editor")
        if st.button("儲存修改"):
            upload_to_db(edited_df)
            st.success("變更已存入資料庫")

with tab3:
    if not df.empty:
        # 依營收分類彙整
        months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
        df['年度小計'] = df[months].sum(axis=1)
        
        agg = df.groupby(['營收分類', '紀錄類型'])['年度小計'].sum().unstack(fill_value=0)
        
        # 計算彙整表欄位
        res = pd.DataFrame(index=agg.index)
        res['目標收入'] = agg.get('收入', 0)
        res['預估收入'] = agg.get('收入預估', 0)
        res['目標毛利'] = agg.get('收入', 0) - agg.get('支出', 0)
        res['預估毛利'] = agg.get('收入預估', 0) - agg.get('支出預估', 0)
        res['預估毛利率'] = (res['預估毛利'] / res['預估收入']).replace([np.inf, -np.inf], 0).fillna(0)
        res['差異'] = res['目標收入'] - res['預估收入']
        
        st.dataframe(res.style.format({'預估毛利率': '{:.2%}'}))
