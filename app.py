import streamlit as st
import pandas as pd
from sqlalchemy import text

# 頁面配置
st.set_page_config(page_title="專案營收管理系統", layout="wide")

# 建立資料庫連線
try:
    conn = st.connection("postgresql", type="sql")
except Exception as e:
    st.error("❌ 無法連線至資料庫，請檢查 Secrets 設定。")
    st.stop()

# --- 核心運算函數 ---
def get_data():
    """從資料庫抓取最新資料"""
    return conn.query('SELECT * FROM financials ORDER BY id DESC', ttl=0)

def calculate_yearly_sum(df):
    """將 1-12 月欄位加總為年度總額"""
    months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
    # 轉換數值，避免 Excel 匯入時產生字串
    temp_df = df.copy()
    for m in months:
        temp_df[m] = pd.to_numeric(temp_df[m], errors='coerce').fillna(0)
    return temp_df[months].sum(axis=1)

# --- 主介面 Tabs ---
tab1, tab2, tab3 = st.tabs(["📊 各專案推進營收", "⚙️ 數據管理庫", "📈 營收分類彙整表"])

# --- Tab 1: 各專案推進營收 (卡片呈現) ---
with tab1:
    st.header("專案進度看板")
    df = get_data()
    if not df.empty:
        df['年度總額'] = calculate_yearly_sum(df)
        projects = df['專案說明'].unique()
        
        cols = st.columns(3)
        for i, project in enumerate(projects):
            with cols[i % 3]:
                p_data = df[df['專案說明'] == project]
                # 篩選：目標 (收入) 與 預估 (收入預估)
                goal_rev = p_data[p_data['紀錄類型'] == '收入']['年度總額'].sum()
                est_rev = p_data[p_data['紀錄類型'] == '收入預估']['年度總額'].sum()
                category = p_data['營收分類'].iloc[0] if not p_data.empty else "未分類"
                
                prog_rate = (est_rev / goal_rev) if goal_rev != 0 else 0
                
                with st.container(border=True):
                    st.subheader(f"📁 {project}")
                    st.caption(f"分類：{category}")
                    # 使用小圖標標示類別
                    st.metric("目標營收 (收入)", f"${goal_rev:,.0f}")
                    st.metric("預估收入 (收入預估)", f"${est_rev:,.0f}", f"{prog_rate:.1%}")
                    st.progress(min(prog_rate, 1.0) if prog_rate >= 0 else 0)
    else:
        st.info("目前尚無資料，請至管理庫匯入 Excel。")

# --- Tab 2: 數據管理庫 (編輯與刪除) ---
with tab2:
    st.header("資料庫管理中心")
    
    col_up, col_del = st.columns([2, 1])
    
    with col_up:
        # A. 匯入功能
        with st.expander("📥 匯入新專案 Excel"):
            uploaded_file = st.file_uploader("選擇檔案", type="xlsx")
            if uploaded_file:
                new_df = pd.read_excel(uploaded_file)
                new_df.columns = new_df.columns.str.strip() # 清理空白
                st.write("預覽資料：", new_df.head(3))
                if st.button("確認寫入 DB"):
                    with conn.session as s:
                        for _, r in new_df.iterrows():
                            query = text("""
                                INSERT INTO financials ("專案說明", "Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec", "營收分類", "紀錄類型", "說明")
                                VALUES (:p, :m1, :m2, :m3, :m4, :m5, :m6, :m7, :m8, :m9, :m10, :m11, :m12, :cat, :type, :desc)
                            """)
                            s.execute(query, {
                                "p": r.get('專案說明'), "cat": r.get('營收分類'), "type": r.get('紀錄類型'), "desc": r.get('說明',''),
                                "m1":r.get('Jan',0), "m2":r.get('Feb',0), "m3":r.get('Mar',0), "m4":r.get('Apr',0), "m5":r.get('May',0), "m6":r.get('Jun',0),
                                "m7":r.get('Jul',0), "m8":r.get('Aug',0), "m9":r.get('Sep',0), "m10":r.get('Oct',0), "m11":r.get('Nov',0), "m12":r.get('Dec',0)
                            })
                        s.commit()
                    st.success("資料匯入成功！")
                    st.rerun()

    with col_del:
        # B. 危險操作區
        with st.expander("⚠️ 危險操作"):
            if st.button("🔥 清空資料庫所有資料", type="primary"):
                with conn.session as s:
                    s.execute(text("TRUNCATE TABLE financials;"))
                    s.commit()
                st.rerun()

    # C. 即時編輯與紅字標記
    st.divider()
    st.subheader("📝 數據編輯與批次管理")
    manage_df = get_data()
    
    if not manage_df.empty:
        # 建立紅字標記邏輯：如果紀錄類型包含「預估」或「支出」，在表格中標示
        def style_rows(row):
            if '預估' in str(row['紀錄類型']):
                return ['color: #e91e63'] * len(row) # 粉紅字
            elif '支出' in str(row['紀錄類型']):
                return ['color: #1e88e5'] * len(row) # 藍字
            return [''] * len(row)

        # 批次刪除勾選
        c1, c2 = st.columns([3, 1])
        with c2:
            selected_ids = st.multiselect("勾選 ID 進行批次刪除", manage_df['id'].tolist())
            if st.button("🗑️ 執行刪除", use_container_width=True) and selected_ids:
                with conn.session as s:
                    s.execute(text("DELETE FROM financials WHERE id IN :ids"), {"ids": tuple(selected_ids)})
                    s.commit()
                st.rerun()
        
        with c1:
            st.info("💡 提示：可以直接在下方表格修改數據，預估類型將以紅/藍字區分。")
            st.data_editor(
                manage_df, 
                key="db_editor", 
                num_rows="dynamic",
                use_container_width=True
            )

# --- Tab 3: 營收分類彙整表 ---
with tab3:
    st.header("營收彙整分析表")
    sum_df = get_data()
    if not sum_df.empty:
        sum_df['年度總額'] = calculate_yearly_sum(sum_df)
        
        # 透過 Pivot 整理數據，將「紀錄類型」轉為欄位
        pivot = sum_df.pivot_table(
            index='營收分類', 
            columns='紀錄類型', 
            values='年度總額', 
            aggfunc='sum'
        ).fillna(0)
        
        # 確保公式中需要的 4 個核心欄位都存在
        for col in ['收入', '收入預估', '支出', '支出預估']:
            if col not in pivot.columns: pivot[col] = 0

        # --- 執行財務公式運算 ---
        report = pd.DataFrame(index=pivot.index)
        report['目標收入'] = pivot['收入']
        report['預估收入'] = pivot['收入預估']
        
        # 目標毛利 = 收入(灰) - 支出(白)
        report['目標毛利'] = pivot['收入'] - pivot['支出']
        
        # 預估毛利 = 收入預估(粉) - 支出預估(粉)
        report['預估毛利'] = pivot['收入預估'] - pivot['支出預估']
        
        # 預估毛利率 = 預估毛利 / 預估收入
        report['預估毛利率'] = (report['預估毛利'] / report['預估收入']).replace([float('inf'), -float('inf')], 0).fillna(0)
        
        # 差異 = 目標收入 - 預估收入
        report['差異'] = report['目標收入'] - report['預估收入']

        # --- 樣式美化與格式化 ---
        # 差異 > 0 標示紅字 (代表目標比預估高，存在缺口)
        def highlight_diff(val):
            color = 'red' if val > 0 else '#00c853'
            return f'color: {color}; font-weight: bold'

        st.dataframe(
            report.style.format({
                '目標收入': '{:,.0f}', '預估收入': '{:,.0f}',
                '目標毛利': '{:,.0f}', '預估毛利': '{:,.0f}',
                '預估毛利率': '{:.2%}', '差異': '{:,.0f}'
            }).applymap(highlight_diff, subset=['差異']),
            use_container_width=True,
            height=400
        )
        
        # 增加一個小結語
        total_diff = report['差異'].sum()
        if total_diff > 0:
            st.error(f"⚠️ 目前總體營收缺口 (差異加總)：${total_diff:,.0f}")
        else:
            st.success(f"✅ 目前預估營收達成狀況良好！")
            
    else:
        st.warning("暫無資料可進行彙整計算。")
