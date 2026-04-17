import streamlit as st
import pandas as pd
from sqlalchemy import text
import io

# 1. 頁面基礎設定
st.set_page_config(page_title="營收管理系統 v2", layout="wide")

# --- [自動診斷區] 幫助你檢查 Secrets 到底有沒有讀到 ---
with st.expander("🛠️ 系統連線診斷 (若連線失敗請點開)", expanded=False):
    import os
    st.write(f"目前工作目錄: `{os.getcwd()}`")
    st.write("是否有 .streamlit 資料夾:", os.path.exists(".streamlit"))
    st.write("是否有 secrets.toml:", os.path.exists(".streamlit/secrets.toml"))
    st.write("已讀取到的 Secrets 區塊:", list(st.secrets.keys()))

# 2. 建立資料庫連線
try:
    conn = st.connection("postgresql", type="sql")
except Exception as e:
    st.error("❌ 資料庫連線配置失敗，請檢查 secrets.toml 格式。")
    st.stop()

# --- 核心函數 ---
def fetch_data():
    """從 DB 抓取最新資料"""
    return conn.query('SELECT * FROM financials ORDER BY "建立時間" DESC', ttl=0)

def calculate_row_total(df):
    """計算 1-12 月的加總"""
    months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
    # 確保所有月份欄位都存在且為數值
    temp_df = df.copy()
    for m in months:
        if m not in temp_df.columns:
            temp_df[m] = 0
        temp_df[m] = pd.to_numeric(temp_df[m], errors='coerce').fillna(0)
    return temp_df[months].sum(axis=1)

# --- UI 介面 ---
tabs = st.tabs(["📊 各專案推進營收", "⚙️ 數據管理庫", "📈 營收分類彙整"])

# --- Tab 1: 各專案推進營收 ---
with tabs[0]:
    st.header("專案營收推進看板")
    df = fetch_data()
    if not df.empty:
        df['年度總額'] = calculate_row_total(df)
        projects = df['專案說明'].unique()
        
        cols = st.columns(3)
        for i, project in enumerate(projects):
            p_data = df[df['專案說明'] == project]
            
            # 取得目標營收 (灰底：收入) 與 預估收入 (粉底：收入預估)
            target_val = p_data[p_data['紀錄類型'] == '收入']['年度總額'].sum()
            est_val = p_data[p_data['紀錄類型'] == '收入預估']['年度總額'].sum()
            cat = p_data['營收分類'].iloc[0] if not p_data.empty else "N/A"
            
            # 計算推進率
            prog_rate = (est_val / target_val) if target_val > 0 else 0
            
            with cols[i % 3]:
                with st.container(border=True):
                    st.subheader(project)
                    st.caption(f"分類: {cat}")
                    st.metric("目標營收 (收入)", f"${target_val:,.0f}")
                    st.metric("預估收入", f"${est_val:,.0f}", f"{prog_rate:.1%}")
                    st.progress(min(prog_rate, 1.0))
    else:
        st.info("目前無資料，請至管理庫匯入。")

# --- Tab 2: 數據管理庫 ---
with tabs[1]:
    st.header("資料維護中心")
    
    # 刪除功能區
    with st.expander("🗑️ 危險操作：刪除資料"):
        c1, c2 = st.columns(2)
        with c1:
            st.write("👉 **批次刪除勾選內容**")
            df_for_del = fetch_data()
            if not df_for_del.empty:
                to_delete = st.multiselect("選擇要刪除的 ID", df_for_del['id'].tolist())
                if st.button("確認刪除選中項", type="primary") and to_delete:
                    with conn.session as s:
                        s.execute(text("DELETE FROM financials WHERE id IN :ids"), {"ids": tuple(to_delete)})
                        s.commit()
                    st.rerun()
        with c2:
            st.write("👉 **清空全表**")
            if st.button("🔥 一鍵清空所有資料", help="注意！此操作不可逆"):
                with conn.session as s:
                    s.execute(text("TRUNCATE TABLE financials;"))
                    s.commit()
                st.rerun()

    st.divider()

    # 匯入功能
    st.subheader("📥 匯入 Excel 資料")
    uploaded_file = st.file_uploader("選擇 Excel 檔案", type=["xlsx"])
    if uploaded_file:
        raw_df = pd.read_excel(uploaded_file)
        raw_df.columns = raw_df.columns.str.strip() # 去除標題空格
        st.write("匯入預覽:", raw_df.head(2))
        
        if st.button("確認寫入資料庫"):
            with conn.session as s:
                for _, row in raw_df.iterrows():
                    sql = text("""
                        INSERT INTO financials ("專案說明", "Jan", "Feb", "Mar", "Apr", "May", "Jun", 
                                            "Jul", "Aug", "Sep", "Oct", "Nov", "Dec", 
                                            "營收分類", "紀錄類型", "說明")
                        VALUES (:p, :m1, :m2, :m3, :m4, :m5, :m6, :m7, :m8, :m9, :m10, :m11, :m12, :cat, :type, :desc)
                    """)
                    s.execute(sql, {
                        "p": row.get('專案說明'), "cat": row.get('營收分類'), "type": row.get('紀錄類型'), "desc": row.get('說明'),
                        "m1": row.get('Jan', 0), "m2": row.get('Feb', 0), "m3": row.get('Mar', 0),
                        "m4": row.get('Apr', 0), "m5": row.get('May', 0), "m6": row.get('Jun', 0),
                        "m7": row.get('Jul', 0), "m8": row.get('Aug', 0), "m9": row.get('Sep', 0),
                        "m10": row.get('Oct', 0), "m11": row.get('Nov', 0), "m12": row.get('Dec', 0)
                    })
                s.commit()
            st.success("寫入成功！")
            st.rerun()

    # 資料編輯區
    st.subheader("📝 手動更新與標記")
    edit_df = fetch_data()
    if not edit_df.empty:
        # 使用 st.data_editor，標示紅字邏輯會在顯示時套用
        st.info("小撇步：雙擊儲存格可直接修改值。")
        st.data_editor(edit_df, key="main_editor", use_container_width=True)

# --- Tab 3: 營收分類彙整表 ---
with tabs[2]:
    st.header("營收分類彙整表")
    sum_df = fetch_data()
    if not sum_df.empty:
        sum_df['年度總額'] = calculate_row_total(sum_df)
        
        # 樞紐分析邏輯
        pivot = sum_df.pivot_table(
            index='營收分類', 
            columns='紀錄類型', 
            values='年度總額', 
            aggfunc='sum'
        ).fillna(0)
        
        # 確保必要欄位都存在
        for c in ['收入', '收入預估', '支出', '支出預估']:
            if c not in pivot.columns: pivot[c] = 0
            
        # 計算指標
        res = pd.DataFrame(index=pivot.index)
        res['目標收入'] = pivot['收入']
        res['預估收入'] = pivot['收入預估']
        res['目標毛利'] = pivot['收入'] - pivot['支出']
        res['預估毛利'] = pivot['收入預估'] - pivot['支出預估']
        res['預估毛利率'] = (res['預估毛利'] / res['預估收入']).fillna(0)
        res['差異'] = res['目標收入'] - res['預估收入']
        
        # 格式化顯示
        format_dict = {
            '目標收入': '{:,.0f}', '預估收入': '{:,.0f}', 
            '目標毛利': '{:,.0f}', '預估毛利': '{:,.0f}', 
            '差異': '{:,.0f}', '預估毛利率': '{:.2%}'
        }
        
        # 套用樣式：如果「差異」大於 0（目標 > 預估），標示為紅色
        def highlight_diff(val):
            color = 'red' if val > 0 else 'white'
            return f'color: {color}'

        st.dataframe(res.style.format(format_dict).applymap(highlight_diff, subset=['差異']))
    else:
        st.warning("暫無資料。")
