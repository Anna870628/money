import streamlit as st
import pandas as pd
from sqlalchemy import text

# 1. 頁面配置
st.set_page_config(page_title="營收管理系統 v5.3", layout="wide")

# 自定義樣式
st.markdown("""
    <style>
    .stMetric { background-color: #ffffff; padding: 15px; border-radius: 10px; border: 1px solid #eee; }
    [data-testid="stSidebar"] { background-color: #f8f9fa; }
    .project-header { color: #1f4e79; border-left: 5px solid #1f4e79; padding-left: 10px; margin-bottom: 20px; font-weight: bold; }
    </style>
""", unsafe_allow_html=True)

# 建立連線
try:
    conn = st.connection("postgresql", type="sql")
except Exception as e:
    st.error("❌ 資料庫連線失敗")
    st.stop()

# --- 核心運算函數 ---
def get_clean_data():
    df = conn.query('SELECT * FROM financials ORDER BY id ASC', ttl=0)
    if not df.empty:
        for col in df.select_dtypes(['object']).columns:
            df[col] = df[col].astype(str).str.strip()
    return df

def calculate_yearly_sum(df):
    months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
    temp_df = df.copy()
    for m in months:
        temp_df[m] = pd.to_numeric(temp_df[m], errors='coerce').fillna(0)
    return temp_df[months].sum(axis=1)

# ==========================================
# ⬅️ 左側邊欄：控制中心
# ==========================================
with st.sidebar:
    st.title("🎛️ 營收數據管理")
    
    st.subheader("📥 匯入 Excel")
    uploaded_file = st.file_uploader("選擇檔案", type=["xlsx"])
    
    if uploaded_file:
        try:
            # 讀取 Excel (預設第 0 列為標題，若標題在下方可加 header=1)
            new_df = pd.read_excel(uploaded_file)
            
            # 清理欄位名稱空白
            new_df.columns = [str(c).strip() for c in new_df.columns]
            
            # --- 強大版：自動偵測可能的欄位名稱 ---
            proj_col = next((c for c in new_df.columns if "專案" in c), None)
            type_col = next((c for c in new_df.columns if "紀錄" in c or "類型" in c), None)
            cat_col = next((c for c in new_df.columns if "分類" in c), None)
            
            if not proj_col or not type_col:
                st.error(f"❌ 找不到必要欄位！")
                st.warning(f"Excel 內的欄位目前是：{list(new_df.columns)}")
                st.info("請確認標題列是否有『專案說明』與『紀錄類型』這兩個字眼。")
                st.stop()

            # --- 處理合併儲存格 (ffill) ---
            new_df[proj_col] = new_df[proj_col].ffill()
            if cat_col:
                new_df[cat_col] = new_df[cat_col].ffill()
            
            st.success(f"✅ 成功辨識欄位：{proj_col}")
            st.write("預覽修復後數據：", new_df[[proj_col, type_col]].head(6))
            
            if st.button("🚀 確認寫入資料庫", use_container_width=True):
                with conn.session as s:
                    for _, r in new_df.iterrows():
                        # 跳過完全空值的列
                        if pd.isna(r.get(proj_col)) or str(r.get(proj_col)) == 'nan':
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
                            "p": str(r.get(proj_col)),
                            "m1": r.get('Jan', 0), "m2": r.get('Feb', 0), "m3": r.get('Mar', 0),
                            "m4": r.get('Apr', 0), "m5": r.get('May', 0), "m6": r.get('Jun', 0),
                            "m7": r.get('Jul', 0), "m8": r.get('Aug', 0), "m9": r.get('Sep', 0),
                            "m10": r.get('Oct', 0), "m11": r.get('Nov', 0), "m12": r.get('Dec', 0),
                            "cat": str(r.get(cat_col, '未分類')) if cat_col else '未分類',
                            "type": str(r.get(type_col, '未知')),
                            "desc": str(r.get('說明', ''))
                        }
                        s.execute(sql_cmd, params)
                    s.commit()
                st.toast("數據同步完成！", icon="🎉")
                st.rerun()
        except Exception as e:
            st.error(f"解析發生錯誤：{e}")

    st.divider()
    st.subheader("🗑️ 清理資料")
    db_df = get_clean_data()
    if not db_df.empty:
        selected_ids = st.multiselect("勾選 ID 刪除", sorted(db_df['id'].unique().tolist()))
        if st.button("🗑️ 刪除勾選內容", type="primary", use_container_width=True) and selected_ids:
            with conn.session as s:
                s.execute(text("DELETE FROM financials WHERE id IN :ids"), {"ids": tuple(selected_ids)})
                s.commit()
            st.rerun()

# ==========================================
# 🏠 主畫面內容
# ==========================================
st.title("📊 營收管理戰情室")
df = get_clean_data()

if df.empty:
    st.info("👋 你好！請先上傳 Excel 檔案。")
else:
    df['年度總額'] = calculate_yearly_sum(df)
    tab1, tab2, tab3 = st.tabs(["🚀 專案看板", "📈 分類彙整", "📝 資料檢視"])

    # --- Tab 1: 看板 ---
    with tab1:
        st.markdown("<h3 class='project-header'>各專案推進率</h3>", unsafe_allow_html=True)
        all_cats = sorted(df['營收分類'].unique().tolist())
        selected_cat = st.multiselect("🔍 篩選分類", all_cats, default=all_cats)
        
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
                    m1.metric("目標", f"${target_rev:,.0f}")
                    m2.metric("預估", f"${est_rev:,.0f}")
                    m3.metric("差異", f"${est_rev-target_rev:,.0f}")
                    st.progress(min(rate, 1.0) if rate >= 0 else 0)

    # --- Tab 2: 彙整 ---
    with tab2:
        st.markdown("<h3 class='project-header'>營收分類分析表</h3>", unsafe_allow_html=True)
        pivot = df.pivot_table(index='營收分類', columns='紀錄類型', values='年度總額', aggfunc='sum').fillna(0)
        
        rep = pd.DataFrame(index=pivot.index)
        col_map = {
            '目標': [c for c in pivot.columns if c == '收入'],
            '預估': [c for c in pivot.columns if '收入' in c and '預估' in c],
            '支出': [c for c in pivot.columns if c == '支出'],
            '支預': [c for c in pivot.columns if '支出' in c and '預估' in c]
        }
        
        rep['目標收入'] = pivot[col_map['目標'][0]] if col_map['目標'] else 0
        rep['預估收入'] = pivot[col_map['預估'][0]] if col_map['預估'] else 0
        target_exp = pivot[col_map['支出'][0]] if col_map['支出'] else 0
        est_exp = pivot[col_map['支預'][0]] if col_map['支預'] else 0
        
        rep['目標毛利'] = rep['目標收入'] - target_exp
        rep['預估毛利'] = rep['預估收入'] - est_exp
        rep['預估毛利率'] = (rep['預估毛利'] / rep['預估收入']).fillna(0)
        rep['營收差異'] = rep['預估收入'] - rep['目標收入']

        # 這裡使用新版 Pandas 支援的 .map
        styled_rep = rep.style.format({
            '目標收入': '{:,.0f}', '預估收入': '{:,.0f}',
            '目標毛利': '{:,.0f}', '預估毛利': '{:,.0f}',
            '預估毛利率': '{:.1%}', '營收差異': '{:,.0f}'
        }).map(lambda x: 'color: red' if x < 0 else 'color: green', subset=['營收差異'])
        
        st.dataframe(styled_rep, use_container_width=True)

    # --- Tab 3: 管理 ---
    with tab3:
        st.data_editor(df, use_container_width=True)
