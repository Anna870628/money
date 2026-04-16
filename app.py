import streamlit as st
import pandas as pd
from sqlalchemy import text

# ==========================================
# 0. 系統環境設定與密碼檢查
# ==========================================
def check_password():
    def password_entered():
        if st.session_state["password"] == st.secrets["passwords"]["admin_password"]:
            st.session_state["password_correct"] = True
            del st.session_state["password"]
        else:
            st.session_state["password_correct"] = False

    if "password_correct" not in st.session_state:
        st.title("🔒 營收管理系統 - 身份驗證")
        st.text_input("請輸入存取密碼", type="password", on_change=password_entered, key="password")
        return False
    elif not st.session_state["password_correct"]:
        st.title("🔒 營收管理系統 - 身份驗證")
        st.text_input("請輸入存取密碼", type="password", on_change=password_entered, key="password")
        st.error("😕 密碼錯誤")
        return False
    return True

if check_password():
    st.set_page_config(page_title="車聯網營收管理系統 v3", layout="wide")
    
    # 建立雲端 PostgreSQL 連線
    conn = st.connection("postgresql", type="sql")

    # ==========================================
    # 1. 資料庫與計算邏輯
    # ==========================================
    def load_data():
        try:
            df = conn.query("SELECT * FROM financials", ttl="0")
            return df
        except Exception:
            return pd.DataFrame(columns=['專案說明', '紀錄類型'] + ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'] + ['營收分類', '說明'])

    def save_to_supabase(df):
        with conn.session as session:
            session.execute(text("DELETE FROM financials")) 
            df.to_sql('financials', conn.engine, if_exists='append', index=False)
            session.commit()

    # ==========================================
    # 2. 介面呈現與分頁
    # ==========================================
    st.title("📊 車聯網事業本部 - 專業營收管理系統")
    
    # 使用 Tab 區分「編輯介面」與「分析摘要」
    tab_edit, tab_summary = st.tabs(["📝 數據編輯與匯入", "📈 營收分析摘要 (Summary Sheet)"])

    with tab_edit:
        st.info("🟢 已連線至 Supabase | 提示：編輯完畢請點擊下方儲存按鈕")
        
        # 側邊欄：匯入功能
        with st.sidebar:
            st.header("📂 數據匯入")
            uploaded_file = st.file_uploader("匯入報表 (CSV/Excel)", type=["csv", "xlsx"])
            # (此處保留之前的匯入 process 邏輯，為節省空間簡化)
            
        data = load_data()
        
        # 設定編輯器的下拉選單與顏色格式
        edited_df = st.data_editor(
            data,
            num_rows="dynamic",
            use_container_width=True,
            height=500,
            column_config={
                "紀錄類型": st.column_config.SelectboxColumn(
                    "紀錄類型",
                    options=["收入", "支出", "收入預估", "支出預估", "收入差異", "支出差異"],
                    required=True
                ),
                "營收分類": st.column_config.SelectboxColumn(
                    "營收分類",
                    options=["24DCM開發/維運", "TOYOTA聯網服務", "LEXUS聯網服務", "其他"],
                    required=True
                ),
                # 設定月份顯示格式，不顯示小數點
                **{m: st.column_config.NumberColumn(format="%.0f") for m in ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec']}
            }
        )

        if st.button("💾 儲存所有變更至雲端", type="primary"):
            save_to_supabase(edited_df)
            st.success("資料已永久儲存！")
            st.rerun()

    with tab_summary:
        st.header("📋 營收分類匯總分析")
        
        if not edited_df.empty:
            df = edited_df.copy()
            months = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec']
            df['年度總計'] = df[months].sum(axis=1)

            # --- 計算邏輯 ---
            # 1. 目標營收 (收入)
            target_rev = df[df['紀錄類型'] == '收入'].groupby('營收分類')['年度總計'].sum()
            # 2. 預估營收 (收入預估)
            est_rev = df[df['紀錄類型'] == '收入預估'].groupby('營收分類')['年度總計'].sum()
            # 3. 實際支出
            actual_exp = df[df['紀錄類型'] == '支出'].groupby('營營分類')['年度總計'].sum()
            # 4. 預估支出
            est_exp = df[df['紀錄類型'] == '支出預估'].groupby('營收分類')['年度總計'].sum()

            # 合併成摘要表
            summary = pd.DataFrame({
                '目標營收': target_rev,
                '預估營收': est_rev,
                '實際支出': actual_exp,
                '預估支出': est_exp
            }).fillna(0)

            summary['營收差異 (目標-預估)'] = summary['目標營收'] - summary['預估營收']
            summary['原毛利'] = summary['目標營收'] - summary['實際支出']
            summary['預估毛利'] = summary['預估營收'] - summary['預估支出']
            
            # 計算毛利率 (預估毛利 / 預估營收)
            summary['預估毛利率'] = (summary['預估毛利'] / summary['預估營收']).replace([float('inf'), -float('inf')], 0).fillna(0)

            # --- 顏色處理函數 ---
            def color_diff(val):
                color = 'red' if val < 0 else 'green' if val > 0 else 'black'
                return f'color: {color}'

            # 呈現摘要表格
            st.subheader("📊 分類統計總覽")
            st.dataframe(
                summary.style.format({
                    '目標營收': '{:,.0f}', '預估營收': '{:,.0f}', 
                    '營收差異 (目標-預估)': '{:,.0f}', '原毛利': '{:,.0f}', 
                    '預估毛利': '{:,.0f}', '預估毛利率': '{:.2%}'
                }).applymap(color_diff, subset=['營收差異 (目標-預估)', '預估毛利']),
                use_container_width=True
            )

            # --- 差異項目 PICK 出來 ---
            st.divider()
            st.subheader("🔍 營收差異項目明細 (目標 ≠ 預估)")
            
            # 找出同一專案下，收入與收入預估不等的項目
            diff_items = []
            for project in df['專案說明'].unique():
                proj_data = df[df['專案說明'] == project]
                t_val = proj_data[proj_data['紀錄類型'] == '收入']['年度總計'].sum()
                e_val = proj_data[proj_data['紀錄類型'] == '收入預估']['年度總計'].sum()
                
                if t_val != e_val:
                    diff_items.append({
                        '專案名稱': project,
                        '營收分類': proj_data['營收分類'].iloc[0],
                        '目標營收': t_val,
                        '預估營收': e_val,
                        '差異值': t_val - e_val,
                        '說明': proj_data['說明'].iloc[0] if not proj_data['說明'].empty else ""
                    })
            
            if diff_items:
                diff_df = pd.DataFrame(diff_items)
                st.table(diff_df)
            else:
                st.write("✅ 目前所有專案目標與預估皆一致。")
