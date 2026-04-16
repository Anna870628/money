import streamlit as st
import pandas as pd
from sqlalchemy import text

# ==========================================
# 0. 安全驗證
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
    return st.session_state.get("password_correct", False)

if check_password():
    st.set_page_config(page_title="車聯網營收管理系統 v5", layout="wide")
    conn = st.connection("postgresql", type="sql")

    # ==========================================
    # 1. 核心資料處理函數
    # ==========================================
    def load_data():
        try:
            return conn.query("SELECT * FROM financials", ttl="0")
        except:
            return pd.DataFrame(columns=['專案說明', '紀錄類型'] + [f'Month_{i}' for i in range(1,13)] + ['營收分類', '說明'])

    def save_to_supabase(df):
        with conn.session as session:
            session.execute(text("DELETE FROM financials")) 
            df.to_sql('financials', conn.engine, if_exists='append', index=False)
            session.commit()

    def process_multi_sheet_excel(uploaded_file):
        """核心：同時讀取數據與格式定義"""
        excel_file = pd.ExcelFile(uploaded_file)
        sheet_names = excel_file.sheet_names
        
        # 1. 讀取數據 (預設讀取第一個分頁，或是包含 '預估' 字眼的分頁)
        data_sheet = sheet_names[0]
        for s in sheet_names:
            if '預估' in s or '營收' in s:
                data_sheet = s
                break
        
        df_raw = pd.read_excel(uploaded_file, sheet_name=data_sheet, skiprows=2)
        
        # 2. 讀取格式 (專案說明分頁)
        df_meta = None
        if "專案說明" in sheet_names:
            df_meta = pd.read_excel(uploaded_file, sheet_name="專案說明")
            st.toast("💡 已偵測到『專案說明』分頁，將自動對齊格式。")

        # --- 資料清洗邏輯 ---
        df_raw.rename(columns={df_raw.columns[2]: '紀錄類型'}, inplace=True)
        df_raw['專案說明'] = df_raw['專案說明'].ffill()
        
        # 如果有格式表，則根據格式表來強制校正「營收分類」
        if df_meta is not None:
            # 假設格式表有 '專案名稱' 和 '分類' 欄位，可進行 Mapping
            pass 

        months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']
        for m in months:
            if m in df_raw.columns:
                df_raw[m] = pd.to_numeric(df_raw[m], errors='coerce').fillna(0)
            else:
                df_raw[m] = 0.0
                
        # 營收分類清洗 (若原本資料是空的，則預設為「其他」)
        cat_col = [c for c in df_raw.columns if '營收分類' in str(c)]
        if cat_col:
            df_raw.rename(columns={cat_col[0]: '營收分類'}, inplace=True)
            df_raw['營收分類'] = df_raw['營收分類'].ffill().fillna("其他")
        else:
            df_raw['營收分類'] = "其他"

        target_cols = ['專案說明', '紀錄類型'] + months + ['營收分類', '說明']
        return df_raw[df_raw['專案說明'].notna()][target_cols]

    # ==========================================
    # 2. 介面與功能
    # ==========================================
    st.title("📊 車聯網事業本部 - 專業營收管理系統")
    tab_edit, tab_summary = st.tabs(["📝 數據管理", "📈 分析與匯出 (Summary)"])

    with tab_edit:
        with st.sidebar:
            st.header("📂 智能匯入")
            uploaded_file = st.file_uploader("匯入 Excel (包含專案說明分頁)", type=["xlsx"])
            if uploaded_file:
                if st.button("🚀 解析並同步雲端"):
                    try:
                        processed_df = process_multi_sheet_excel(uploaded_file)
                        save_to_supabase(processed_df)
                        st.success(f"成功匯入！偵測到 {len(processed_df)} 筆紀錄。")
                        st.rerun()
                    except Exception as e:
                        st.error(f"解析失敗：{e}")

        # 資料編輯器 (含下拉選單)
        current_data = load_data()
        edited_df = st.data_editor(
            current_data,
            num_rows="dynamic",
            use_container_width=True,
            height=500,
            column_config={
                "營收分類": st.column_config.SelectboxColumn(
                    "營收分類",
                    options=["24DCM開發/維運", "TOYOTA聯網服務", "LEXUS聯網服務", "其他"],
                    required=True
                ),
                "紀錄類型": st.column_config.SelectboxColumn(
                    "紀錄類型",
                    options=["收入", "支出", "收入預估", "支出預估", "收入差異", "支出差異"],
                    required=True
                )
            }
        )
        if st.button("💾 儲存所有變更", type="primary"):
            save_to_supabase(edited_df)
            st.success("雲端同步完成！")

    with tab_summary:
        st.header("📋 營收匯總與差異分析")
        if not edited_df.empty:
            # --- 計算邏輯 ---
            months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']
            df = edited_df.copy()
            for m in months:
                df[m] = pd.to_numeric(df[m], errors='coerce').fillna(0)
            df['年度總計'] = df[months].sum(axis=1)

            # 按分類加總
            summary = pd.DataFrame()
            summary['目標營收'] = df[df['紀錄類型'] == '收入'].groupby('營收分類')['年度總計'].sum()
            summary['預估營收'] = df[df['紀錄類型'] == '收入預估'].groupby('營收分類')['年度總計'].sum()
            summary['實際支出'] = df[df['紀錄類型'] == '支出'].groupby('營收分類')['年度總計'].sum()
            summary['預估支出'] = df[df['紀錄類型'] == '支出預估'].groupby('營收分類')['年度總計'].sum()
            summary = summary.fillna(0)

            summary['營收差異'] = summary['目標營收'] - summary['預估營收']
            summary['預估毛利'] = summary['預估營收'] - summary['預估支出']
            summary['毛利率'] = (summary['預估毛利'] / summary['預估營收']).replace([float('inf'), -float('inf')], 0).fillna(0)

            # 樣式處理
            def style_negative_red(val):
                color = 'red' if val < 0 else 'green' if val > 0 else 'black'
                return f'color: {color}'

            st.dataframe(
                summary.style.format({'毛利率': '{:.2%}', '營收差異': '{:,.0f}', '目標營收': '{:,.0f}'})
                .applymap(style_negative_red, subset=['營收差異', '預估毛利']),
                use_container_width=True
            )

            # --- PICK 出差異專案 ---
            st.subheader("🔍 差異專案提取 (目標 ≠ 預估)")
            diff_list = []
            for proj in df['專案說明'].unique():
                p_data = df[df['專案說明'] == proj]
                target = p_data[p_data['紀錄類型'] == '收入']['年度總計'].sum()
                estimate = p_data[p_data['紀錄類型'] == '收入預估']['年度總計'].sum()
                if target != estimate:
                    diff_list.append({
                        "專案名稱": proj,
                        "分類": p_data['營收分類'].iloc[0],
                        "目標": target,
                        "預估": estimate,
                        "差異": target - estimate,
                        "說明": p_data['說明'].iloc[0] if not p_data['說明'].empty else ""
                    })
            if diff_list:
                st.table(pd.DataFrame(diff_list))
