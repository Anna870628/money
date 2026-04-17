import streamlit as st
import pandas as pd
from sqlalchemy import text

# ==========================================
# 0. 安全驗證與基礎設定
# ==========================================
def check_password():
    def password_entered():
        if st.session_state["password"] == st.secrets["passwords"]["admin_password"]:
            st.session_state["password_correct"] = True
            del st.session_state["password"]
        else:
            st.session_state["password_correct"] = False

    if "password_correct" not in st.session_state:
        st.title("🔒 車聯網營收管理系統 - 登入")
        st.text_input("請輸入存取密碼", type="password", on_change=password_entered, key="password")
        return False
    elif not st.session_state.get("password_correct", False):
        st.title("🔒 車聯網營收管理系統 - 登入")
        st.text_input("請輸入存取密碼", type="password", on_change=password_entered, key="password")
        st.error("😕 密碼錯誤")
        return False
    return True

if check_password():
    st.set_page_config(page_title="車聯網營收戰情系統 最終版", layout="wide")
    
    # 建立雲端資料庫連線
    conn = st.connection("postgresql", type="sql")

    # ==========================================
    # 1. 資料處理核心邏輯 (極速寫入 & 防死結版)
    # ==========================================
    def load_data():
        try:
            df = conn.query("SELECT * FROM financials", ttl="0")
            months = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec']
            for m in months:
                if m in df.columns:
                    df[m] = pd.to_numeric(df[m], errors='coerce').fillna(0)
            return df
        except:
            return pd.DataFrame(columns=['專案說明', '紀錄類型', '營收分類', '說明'] + ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'])

    def save_to_supabase(df):
        # 移除了容易造成死結的 session DROP，全權交給 Pandas 引擎處理覆蓋寫入
        # 啟用 chunksize 與 method='multi' 大幅提升上傳速度
        df.to_sql('financials', conn.engine, if_exists='replace', index=False, chunksize=500, method='multi')

    def process_imported_file(uploaded_file):
        """讀取並清洗 Excel/CSV 資料，加入智能分頁辨識與極速清理"""
        target_sheet = "預設"
        if uploaded_file.name.endswith(('.xlsx', '.xls')):
            excel_file = pd.ExcelFile(uploaded_file)
            sheet_names = excel_file.sheet_names
            
            # 智能尋找資料分頁
            target_sheet = sheet_names[0]
            for s in sheet_names:
                if '營收' in s or '預估' in s or '收支' in s:
                    target_sheet = s
                    break
            if target_sheet == '專案說明' and len(sheet_names) > 1:
                target_sheet = [s for s in sheet_names if s != '專案說明'][0]
                
            df = pd.read_excel(uploaded_file, sheet_name=target_sheet, skiprows=2)
        else:
            df = pd.read_csv(uploaded_file, skiprows=2)
        
        # 清除欄位名稱前後空白
        df.columns = [str(col).strip() for col in df.columns]
        
        if '專案說明' not in df.columns:
            found_cols = ", ".join([str(c) for c in df.columns])
            raise ValueError(f"在分頁『{target_sheet}』中找不到『專案說明』欄位。\n讀到的欄位有：{found_cols}")

        df.rename(columns={df.columns[2]: '紀錄類型'}, inplace=True)
        
        # 剔除 Excel 幽靈空白列，加速上傳
        df['專案說明'] = df['專案說明'].replace(r'^\s*$', pd.NA, regex=True).ffill()
        df = df.dropna(subset=['專案說明'])

        # 營收分類處理
        cat_col = [c for c in df.columns if '營收分類' in str(c)]
        if cat_col:
            df.rename(columns={cat_col[0]: '營收分類'}, inplace=True)
            df['營收分類'] = df['營收分類'].ffill().fillna("其他")
        else:
            df['營收分類'] = "其他"

        months = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec']
        for m in months:
            if m in df.columns:
                df[m] = pd.to_numeric(df[m], errors='coerce').fillna(0)
            else:
                df[m] = 0.0
                
        if '說明' not in df.columns:
            df['說明'] = ""
            
        target_cols = ['專案說明', '紀錄類型'] + months + ['營收分類', '說明']
        
        # 剔除沒有紀錄類型的無效行
        df = df.dropna(subset=['紀錄類型'])
        return df[target_cols]

    # ==========================================
    # 2. 介面分頁規劃
    # ==========================================
    st.title("📊 車聯網事業本部 - 營收管理戰情室")
    
    tab_cards, tab_edit, tab_summary = st.tabs([
        "🎴 專案戰情卡片", 
        "📝 原始數據管理", 
        "📈 分類分析摘要"
    ])

    data = load_data()

    # --- TAB 1: 專案戰情卡片 ---
    with tab_cards:
        if data.empty:
            st.warning("目前資料庫無資料，請先至『原始數據管理』分頁匯入 Excel 檔案。")
        else:
            st.subheader("💡 專案績效概覽")
            cats = ["全部分類"] + list(data['營收分類'].unique())
            sel_cat = st.selectbox("依分類篩選", cats, key="card_filter")
            
            card_df = data if sel_cat == "全部分類" else data[data['營收分類'] == sel_cat]
            projects = card_df['專案說明'].unique()
            
            cols = st.columns(2)
            for idx, proj in enumerate(projects):
                with cols[idx % 2]:
                    p_df = card_df[card_df['專案說明'] == proj]
                    months = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec']
                    
                    target_rev = p_df[p_df['紀錄類型'] == '收入'][months].sum().sum()
                    est_rev = p_df[p_df['紀錄類型'] == '收入預估'][months].sum().sum()
                    est_exp = p_df[p_df['紀錄類型'] == '支出預估'][months].sum().sum()
                    
                    profit = est_rev - est_exp
                    margin = (profit / est_rev) if est_rev != 0 else 0
                    diff = est_rev - target_rev
                    
                    with st.container(border=True):
                        st.markdown(f"#### {proj}")
                        st.caption(f"營收分類：{p_df['營收分類'].iloc[0]}")
                        
                        m1, m2, m3 = st.columns(3)
                        m1.metric("年度目標營收", f"${target_rev:,.0f}")
                        m2.metric("年度預估營收", f"${est_rev:,.0f}", f"{diff:,.0f}")
                        m3.metric("預估毛利率", f"{margin:.1%}")
                        
                        reach = (est_rev / target_rev) if target_rev != 0 else 0
                        st.write(f"**目標達成率: {reach:.1%}**")
                        st.progress(min(reach, 1.0))
                        
                        with st.expander("展開查看 1-12 月數據"):
                            st.dataframe(p_df, use_container_width=True)

    # --- TAB 2: 原始數據管理 ---
    with tab_edit:
        st.info("提示：此處可直接編輯數值、下拉選單，或上傳新檔案。")
        
        with st.sidebar:
            st.header("📂 數據匯入")
            file = st.file_uploader("匯入 Excel 檔案 (.xlsx)", type=["xlsx", "csv"])
            if file and st.button("🚀 開始解析並覆蓋雲端"):
                with st.spinner("極速處理中..."):
                    try:
                        new_df = process_imported_file(file)
                        save_to_supabase(new_df)
                        st.success("匯入成功！")
                        st.rerun()
                    except Exception as e:
                        st.error(f"解析失敗！\n{e}")
            
            st.divider()
            
            # 使用 engine 執行避免死結
            if st.button("⚠️ 清空資料庫 (重來)"):
                with conn.engine.begin() as db_conn:
                    db_conn.execute(text("DROP TABLE IF EXISTS financials"))
                st.warning("資料庫已清空，請重新匯入檔案")
                st.rerun()
                
            st.divider()
            if st.button("🚪 安全登出"):
                st.session_state["password_correct"] = False
                st.rerun()

        edited_df = st.data_editor(
            data,
            num_rows="dynamic",
            use_container_width=True,
            height=500,
            column_config={
                "營收分類": st.column_config.SelectboxColumn("營收分類", options=["24DCM開發/維運", "TOYOTA聯網服務", "LEXUS聯網服務", "其他"]),
                "紀錄類型": st.column_config.SelectboxColumn("紀錄類型", options=["收入", "收入預估", "支出", "支出預估", "收入差異", "支出差異"]),
                **{m: st.column_config.NumberColumn(format="%.0f") for m in ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec']}
            }
        )
        
        if st.button("💾 儲存變更", type="primary"):
            with st.spinner("同步中..."):
                save_to_supabase(edited_df)
                st.success("雲端資料已更新！")
                st.rerun()

    # --- TAB 3: 分類分析摘要 ---
    with tab_summary:
        st.header("📋 營收匯總與分析")
        if not data.empty:
            df_sum = data.copy()
            months = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec']
            df_sum['年度總計'] = df_sum[months].sum(axis=1)

            summary = pd.DataFrame({
                '目標營收': df_sum[df_sum['紀錄類型'] == '收入'].groupby('營收分類')['年度總計'].sum(),
                '預估營收': df_sum[df_sum['紀錄類型'] == '收入預估'].groupby('營收分類')['年度總計'].sum(),
                '實際支出': df_sum[df_sum['紀錄類型'] == '支出'].groupby('營收分類')['年度總計'].sum(),
                '預估支出': df_sum[df_sum['紀錄類型'] == '支出預估'].groupby('營收分類')['年度總計'].sum(),
            }).fillna(0)

            summary['營收差異'] = summary['目標營收'] - summary['預估營收']
            summary['預估毛利'] = summary['預估營收'] - summary['預估支出']
            summary['預估毛利率'] = (summary['預估毛利'] / summary['預估營收']).replace([float('inf'), -float('inf')], 0).fillna(0)

            def style_negative_red(val):
                if isinstance(val, (int, float)) and val < 0:
                    return 'color: red'
                return ''

            st.subheader("📊 各分類績效總覽")
            st.dataframe(
                summary.style.format({'預估毛利率': '{:.2%}', '目標營收': '{:,.0f}', '預估營收': '{:,.0f}', '營收差異': '{:,.0f}'})
                .map(style_negative_red, subset=['營收差異', '預估毛利']),
                use_container_width=True
            )
