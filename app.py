import streamlit as st
import pandas as pd
from sqlalchemy import text

# 1. 頁面配置
st.set_page_config(page_title="Carmax 營收管理系統 v5.2", layout="wide")

# 自定義介面樣式
st.markdown("""
    <style>
    .stMetric { background-color: #ffffff; padding: 15px; border-radius: 10px; border: 1px solid #eee; box-shadow: 0 2px 4px rgba(0,0,0,0.05); }
    [data-testid="stSidebar"] { background-color: #f8f9fa; }
    .project-header { color: #1f4e79; border-left: 5px solid #1f4e79; padding-left: 10px; margin-bottom: 20px; font-weight: bold; }
    </style>
""", unsafe_allow_html=True)

# 建立資料庫連線
try:
    conn = st.connection("postgresql", type="sql")
except Exception as e:
    st.error("❌ 連線失敗，請檢查側邊欄或 Secrets 設定")
    st.stop()

# --- 核心運算函數 ---
def get_clean_data():
    """從資料庫抓取並清理數據"""
    df = conn.query('SELECT * FROM financials ORDER BY id ASC', ttl=0)
    if not df.empty:
        # 清理字串空白
        for col in df.select_dtypes(['object']).columns:
            df[col] = df[col].astype(str).str.strip()
    return df

def calculate_yearly_sum(df):
    """將 1-12 月欄位加總"""
    months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
    temp_df = df.copy()
    for m in months:
        temp_df[m] = pd.to_numeric(temp_df[m], errors='coerce').fillna(0)
    return temp_df[months].sum(axis=1)

# ==========================================
# ⬅️ 左側邊欄：控制中心 (匯入與刪除)
# ==========================================
with st.sidebar:
    st.title("🎛️ 營收數據中心")
    
    st.subheader("📥 匯入最新 Excel")
    uploaded_file = st.file_uploader("請選擇 Excel 檔案", type=["xlsx"])
    
    if uploaded_file:
        try:
            # 讀取 Excel 並處理合併儲存格
            new_df = pd.read_excel(uploaded_file)
            new_df.columns = new_df.columns.str.strip()
            
            # --- 處理合併儲存格 (核心修正) ---
            if '專案說明' in new_df.columns:
                new_df['專案說明'] = new_df['專案說明'].ffill()
            if '營收分類' in new_df.columns:
                new_df['營收分類'] = new_df['營收分類'].ffill()
            
            st.success("✅ 檔案已成功填充合併儲存格")
            st.write("預覽 (前 6 列)：", new_df[['專案說明', '紀錄類型']].head(6))
            
            if st.button("🚀 確認寫入資料庫", use_container_width=True):
                with conn.session as s:
                    for _, r in new_df.iterrows():
                        if pd.isna(r.get('專案說明')) or str(r.get('專案說明')) == 'nan':
                            continue
                        
                        sql_cmd = text("""
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
                        
                        params = {
                            "p": str(r.get('專案說明')),
                            "m1": r.get('Jan', 0), "m2": r.get('Feb', 0), "m3": r.get('Mar', 0),
                            "m4": r.get('Apr', 0), "m5": r.get('May', 0), "m6": r.get('Jun', 0),
                            "m7": r.get('Jul', 0), "m8": r.get('Aug', 0), "m9": r.get('Sep', 0),
                            "m10": r.get('Oct', 0), "m11": r.get('Nov', 0), "m12": r.get('Dec', 0),
                            "cat": str(r.get('營收分類', '未分類')),
                            "type": str(r.get('紀錄類型', '未知')),
                            "desc": str(r.get('說明', ''))
                        }
                        s.execute(sql_cmd, params)
                    s.commit()
                st.toast("數據匯入成功！", icon="🎉")
                st.rerun()
        except Exception as e:
            st.error(f"解析失敗: {e}")

    st.divider()
    st.subheader("🗑️ 清理舊數據")
    current_df = get_clean_data()
    if not current_df.empty:
        selected_ids = st.multiselect("選擇欲刪除的 ID", sorted(current_df['id'].unique().tolist()))
        if st.button("🗑️ 刪除選中資料", type="primary", use_container_width=True) and selected_ids:
            with conn.session as s:
                s.execute(text("DELETE FROM financials WHERE id IN :ids"), {"ids": tuple(selected_ids)})
                s.commit()
            st.rerun()
        
        if st.checkbox("開啟『一鍵清空』功能"):
            if st.button("🔥 清空所有數據", use_container_width=True):
                with conn.session as s:
                    s.execute(text("TRUNCATE TABLE financials;"))
                    s.commit()
                st.rerun()

# ==========================================
# 🏠 主畫面：分析與報表
# ==========================================
st.title("📈 營收管理戰情室")
df = get_clean_data()

if df.empty:
    st.info("👋 你好！目前資料庫為空，請從左側側邊欄匯入 Excel 檔案。")
else:
    df['年度總額'] = calculate_yearly_sum(df)
    tab1, tab2, tab3 = st.tabs(["🚀 專案進度看板", "📈 分類彙整報表", "📝 資料細目檢視"])

    # --- Tab 1: 專案進度看板 ---
    with tab1:
        st.markdown("<h3 class='project-header'>各專案目標推進率</h3>", unsafe_allow_html=True)
        all_cats = sorted(df['營收分類'].unique().tolist())
        selected_cat = st.multiselect("🔍 篩選營收分類", all_cats, default=all_cats)
        
        display_df = df[df['營收分類'].isin(selected_cat)]
        projects = [p for p in display_df['專案說明'].unique() if p and str(p).lower() != 'nan']
        
        cols = st.columns(2)
        for i, project in enumerate(projects):
            p_rows = display_df[display_df['專案說明'] == project]
            
            # 目標 (收入) vs 預估 (收入預估)
            target_rev = p_rows[p_rows['紀錄類型'].str.contains('^收入$', na=False)]['年度總額'].sum()
            est_rev = p_rows[p_rows['紀錄類型'].str.contains('預估', na=False) & p_rows['紀錄類型'].str.contains('收入', na=False)]['年度總額'].sum()
            
            rate = (est_rev / target_rev) if target_rev > 0 else 0
            
            with cols[i % 2]:
                with st.container(border=True):
                    c1, c2 = st.columns([3, 1])
                    c1.subheader(project)
                    c1.caption(f"分類：{p_rows['營收分類'].iloc[0]}")
                    c2.markdown(f"<h2 style='text-align:right;'>{rate:.0%}</h2>", unsafe_allow_html=True)
                    
                    m1, m2, m3 = st.columns(3)
                    m1.metric("目標營收", f"${target_rev:,.0f}")
                    m2.metric("預估收入", f"${est_rev:,.0f}")
                    m3.metric("差異", f"${est_rev-target_rev:,.0f}")
                    st.progress(min(rate, 1.0) if rate >= 0 else 0)

    # --- Tab 2: 彙整分析 ---
    with tab2:
        st.markdown("<h3 class='project-header'>營收分類分析表</h3>", unsafe_allow_html=True)
        
        # 樞紐分析
        pivot = df.pivot_table(index='營收分類', columns='紀錄類型', values='年度總額', aggfunc='sum').fillna(0)
        
        # 建立分析表
        rep = pd.DataFrame(index=pivot.index)
        
        # 動態抓取列名 (解決 Excel 紀錄類型名稱微差)
        col_target = [c for c in pivot.columns if c == '收入']
        col_est_in = [c for c in pivot.columns if '收入' in c and '預估' in c]
        col_target_ex = [c for c in pivot.columns if c == '支出']
        col_est_ex = [c for c in pivot.columns if '支出' in c and '預估' in c]
        
        rep['目標收入'] = pivot[col_target[0]] if col_target else 0
        rep['預估收入'] = pivot[col_est_in[0]] if col_est_in else 0
        target_exp = pivot[col_target_ex[0]] if col_target_ex else 0
        est_exp = pivot[col_est_ex[0]] if col_est_ex else 0
        
        rep['目標毛利'] = rep['目標收入'] - target_exp
        rep['預估毛利'] = rep['預估收入'] - est_exp
        rep['預估毛利率'] = (rep['預估毛利'] / rep['預估收入']).fillna(0)
        rep['營收差異'] = rep['預估收入'] - rep['目標收入']

        # --- 修正 AttributeError 的地方：將 applymap 改為 map ---
        styled_rep = rep.style.format({
            '目標收入': '{:,.0f}', '預估收入': '{:,.0f}',
            '目標毛利': '{:,.0f}', '預估毛利': '{:,.0f}',
            '預估毛利率': '{:.1%}', '營收差異': '{:,.0f}'
        }).map(lambda x: 'color: red' if x < 0 else 'color: green', subset=['營收差異'])
        
        st.dataframe(styled_rep, use_container_width=True)

    # --- Tab 3: 管理檢視 ---
    with tab3:
        st.markdown("<h3 class='project-header'>明細數據檢核</h3>", unsafe_allow_html=True)
        st.data_editor(df, use_container_width=True)
