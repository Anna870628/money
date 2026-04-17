import streamlit as st
import pandas as pd
from sqlalchemy import text

# ==========================================
# 0. 登入與基礎設定
# ==========================================
def check_password():
    def password_entered():
        if st.session_state["password"] == st.secrets["passwords"]["admin_password"]:
            st.session_state["password_correct"] = True
            del st.session_state["password"]
        else:
            st.session_state["password_correct"] = False
    if "password_correct" not in st.session_state:
        st.title("🔒 營收管理系統 - 登入")
        st.text_input("請輸入密碼", type="password", on_change=password_entered, key="password")
        return False
    return st.session_state.get("password_correct", False)

if check_password():
    st.set_page_config(page_title="車聯網營收戰情系統 v8", layout="wide")
    conn = st.connection("postgresql", type="sql")

    def load_data():
        try:
            df = conn.query("SELECT * FROM financials", ttl="0")
            months = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec']
            for m in months:
                df[m] = pd.to_numeric(df[m], errors='coerce').fillna(0)
            return df
        except:
            return pd.DataFrame(columns=['專案說明', '紀錄類型', '營收分類', '說明'] + ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'])

    def save_to_supabase(df):
        with conn.session as session:
            session.execute(text("DROP TABLE IF EXISTS financials")) 
            df.to_sql('financials', conn.engine, if_exists='replace', index=False)
            session.commit()

    # ==========================================
    # 2. 介面分頁
    # ==========================================
    st.title("📊 專案營收戰情系統")
    tab_cards, tab_edit, tab_summary = st.tabs(["🎴 專案戰情卡片", "📝 數據編輯", "📈 分類匯總"])

    # --- 讀取最新資料 ---
    data = load_data()

    # ==========================================
    # 分頁 1: 卡片式摘要 (Card View)
    # ==========================================
    with tab_cards:
        st.subheader("💡 專案績效一覽表")
        
        if not data.empty:
            # 依營收分類過濾
            all_cats = ["全部分類"] + list(data['營收分類'].unique())
            sel_cat = st.selectbox("篩選營收分類", all_cats)
            
            display_data = data if sel_cat == "全部分類" else data[data['營收分類'] == sel_cat]
            projects = display_data['專案說明'].unique()
            
            # 每列顯示 2 張卡片
            cols = st.columns(2)
            for idx, proj in enumerate(projects):
                with cols[idx % 2]:
                    p_df = display_data[display_data['專案說明'] == proj]
                    months = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec']
                    
                    # 計算核心指標
                    target_rev = p_df[p_df['紀錄類型'] == '收入'][months].sum().sum()
                    est_rev = p_df[p_df['紀錄類型'] == '收入預估'][months].sum().sum()
                    est_exp = p_df[p_df['紀錄類型'] == '支出預估'][months].sum().sum()
                    
                    profit = est_rev - est_exp
                    margin = (profit / est_rev) if est_rev != 0 else 0
                    reach_rate = (est_rev / target_rev) if target_rev != 0 else 0
                    
                    # 卡片 UI 設計
                    with st.container(border=True):
                        st.markdown(f"### {proj}")
                        st.caption(f"營收分類: {p_df['營收分類'].iloc[0]}")
                        
                        m1, m2, m3 = st.columns(3)
                        m1.metric("目標營收", f"{target_rev:,.0f}")
                        m2.metric("預估營收", f"{est_rev:,.0f}", f"{est_rev-target_rev:,.0f}")
                        m3.metric("預估毛利率", f"{margin:.1%}")
                        
                        # 進度條展示達成率
                        st.write(f"**目標達成率: {reach_rate:.1%}**")
                        st.progress(min(reach_rate, 1.0))
                        
                        with st.expander("查看 1-12 月預估明細"):
                            st.table(p_df[p_df['紀錄類型'].isin(['收入預估', '支出預估'])][['紀錄類型'] + months])

    # ==========================================
    # 分頁 2: 數據編輯 (保留原本功能)
    # ==========================================
    with tab_edit:
        st.info("提示：點擊下方的『儲存變更』才會同步到雲端。")
        edited_df = st.data_editor(data, num_rows="dynamic", use_container_width=True, height=500)
        if st.button("💾 儲存變更", type="primary"):
            save_to_supabase(edited_df)
            st.success("同步完成！")
            st.rerun()

    # ==========================================
    # 分頁 3: 分類匯總 (保留原本功能)
    # ==========================================
    with tab_summary:
        st.subheader("📋 營收分類匯總分析")
        # (此處放之前的 Summary 計算邏輯...)
