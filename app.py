import streamlit as st
import pandas as pd
from supabase import create_client, Client
import numpy as np

# --- 1. 初始化 Supabase 連線 ---
url = st.secrets["SUPABASE_URL"]
key = st.secrets["SUPABASE_KEY"]
supabase: Client = create_client(url, key)

st.set_page_config(layout="wide", page_title="營收整理系統")

# --- 2. 資料處理函數 ---
def fetch_data():
    res = supabase.table("revenue_records").select("*").execute()
    return pd.DataFrame(res.data)

def upload_to_supabase(df):
    # 先清空舊資料 (依需求決定是否覆蓋)
    # supabase.table("revenue_records").delete().neq("id", 0).execute()
    data = df.to_dict(orient="records")
    supabase.table("revenue_records").insert(data).execute()

# --- 3. 側邊欄：匯入功能 ---
with st.sidebar:
    st.title("營收系統控制")
    uploaded_file = st.file_uploader("匯入每月收支 Excel", type=["xlsx"])
    
    if uploaded_file:
        if st.button("執行解析並存入 Supabase"):
            # 根據您的檔案結構解析 (每 6 列一組)
            raw_df = pd.read_excel(uploaded_file, header=None)
            all_records = []
            
            # 解析邏輯 (假設從第 3 列開始是數據)
            for i in range(2, len(raw_df), 6):
                proj_name = raw_df.iloc[i, 1] if pd.notna(raw_df.iloc[i, 1]) else "未命名專案"
                category = raw_df.iloc[i, 19] if pd.notna(raw_df.iloc[i, 19]) else "未分類"
                
                # 六列類型定義
                row_labels = ["收入", "收入預估", "支出", "支出預估", "收入差異", "支出差異"]
                
                for idx, label in enumerate(row_labels):
                    if i + idx < len(raw_df):
                        row_vals = raw_df.iloc[i + idx, 3:15].fillna(0).replace('-', 0).astype(float).tolist()
                        record = {
                            "project_name": proj_name,
                            "category": category,
                            "row_type": label,
                            "m1": row_vals[0], "m2": row_vals[1], "m3": row_vals[2], "m4": row_vals[3],
                            "m5": row_vals[4], "m6": row_vals[5], "m7": row_vals[6], "m8": row_vals[7],
                            "m9": row_vals[8], "m10": row_vals[9], "m11": row_vals[10], "m12": row_vals[11],
                            "total": sum(row_vals)
                        }
                        all_records.append(record)
            
            upload_to_supabase(pd.DataFrame(all_records))
            st.success("資料已成功上傳至 Supabase！")
            st.rerun()

# --- 4. 主分頁 UI ---
tab1, tab2, tab3 = st.tabs(["📊 各專案推進營收", "🛠️ 資料庫管理", "📈 營收分類彙整表"])

df = fetch_data()

# --- 分頁 1: 各專案推進營收 (卡片呈現) ---
with tab1:
    if not df.empty:
        projects = df['project_name'].unique()
        cols = st.columns(3)
        for idx, p in enumerate(projects):
            p_df = df[df['project_name'] == p]
            
            target_inc = p_df[p_df['row_type'] == "收入"]['total'].sum()
            est_inc = p_df[p_df['row_type'] == "收入預估"]['total'].sum()
            cat = p_df['category'].iloc[0]
            rate = (est_inc / target_inc * 100) if target_inc != 0 else 0
            
            with cols[idx % 3]:
                st.markdown(f"""
                <div style="padding:15px; border-radius:10px; border:1px solid #E0E0E0; background-color:#F8F9FA; margin-bottom:10px">
                    <small>{cat}</small>
                    <h3 style="margin:0px">{p}</h3>
                    <hr>
                    <p style="margin:2px">目標營收：<b>{target_inc:,.0f}</b></p>
                    <p style="margin:2px">預估收入：<b>{est_inc:,.0f}</b></p>
                    <h4 style="color:{'#D32F2F' if rate < 100 else '#388E3C'}">推進率：{rate:.1f}%</h4>
                </div>
                """, unsafe_allow_html=True)
    else:
        st.info("尚未從側邊欄匯入資料。")

# --- 分頁 2: 資料庫管理 ---
with tab2:
    if not df.empty:
        st.subheader("手動更新與批次刪除")
        # 使用 data_editor 編輯
        edited_df = st.data_editor(
            df, 
            num_rows="dynamic", 
            key="db_editor",
            disabled=["id", "created_at"]
        )
        
        c1, c2 = st.columns(2)
        if c1.button("儲存所有更改"):
            # 這裡簡化處理：全刪除再全寫入，或可撰寫 UPSERT 邏輯
            supabase.table("revenue_records").delete().neq("id", 0).execute()
            # 移除自動生成的 id 欄位再寫入
            save_df = edited_df.drop(columns=['id', 'created_at'], errors='ignore')
            upload_to_supabase(save_df)
            st.success("Supabase 資料庫已更新")
            st.rerun()
            
        st.caption("提示：點擊列最左側可選取並按 Delete 鍵刪除。若數值為負，系統將自動在報表中標示紅字。")

# --- 分頁 3: 營收分類彙整表 ---
with tab3:
    if not df.empty:
        # 轉換資料結構以利計算
        agg = df.groupby(['category', 'row_type'])['total'].sum().unstack(fill_value=0)
        
        # 確保必要欄位存在
        for col in ["收入", "收入預估", "支出", "支出預估"]:
            if col not in agg.columns: agg[col] = 0
            
        report = pd.DataFrame(index=agg.index)
        report['目標收入'] = agg['收入']
        report['預估收入'] = agg['收入預估']
        report['目標毛利'] = agg['收入'] - agg['支出']
        report['預估毛利'] = agg['收入預估'] - agg['支出預估']
        report['預估毛利率'] = (report['預估毛利'] / report['預估收入']).replace([np.inf, -np.inf], 0).fillna(0)
        report['差異(目標-預估)'] = report['目標收入'] - report['預估收入']
        
        # 樣式設定：負數標紅字
        def color_negative_red(val):
            if isinstance(val, (int, float)) and val < 0:
                return 'color: red'
            return ''

        st.dataframe(
            report.style.format({
                '目標收入': '{:,.0f}', '預估收入': '{:,.0f}',
                '目標毛利': '{:,.0f}', '預估毛利': '{:,.0f}',
                '預估毛利率': '{:.2%}', '差異(目標-預估)': '{:,.0f}'
            }).applymap(color_negative_red)
        )
