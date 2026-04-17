import streamlit as st
import pandas as pd
from sqlalchemy import text
import openpyxl

# ==========================================
# 0. 密碼驗證 (CMX_BPT)
# ==========================================
def check_password():
    def password_entered():
        if st.session_state["password"] == st.secrets["passwords"]["admin_password"]:
            st.session_state["password_correct"] = True
            del st.session_state["password"]
        else:
            st.session_state["password_correct"] = False
    if "password_correct" not in st.session_state:
        st.title("🔒 專案管理系統 - 登入")
        st.text_input("請輸入存取密碼", type="password", on_change=password_entered, key="password")
        return False
    return st.session_state.get("password_correct", False)

if check_password():
    st.set_page_config(page_title="車聯網營收系統 v7.1", layout="wide")
    conn = st.connection("postgresql", type="sql")

    # ==========================================
    # 1. 核心資料處理：包含顏色抓取
    # ==========================================
    def load_data():
        try:
            df = conn.query("SELECT * FROM financials", ttl="0")
            # 【修復2】確保從舊資料庫讀取的資料，如果沒有顏色欄位，自動補上
            if '顏色標記' not in df.columns:
                df['顏色標記'] = '無底色'
            return df
        except Exception as e:
            st.warning(f"資料庫讀取提示: {e}")
            return pd.DataFrame(columns=['專案說明', '紀錄類型'] + ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'] + ['營收分類', '顏色標記', '說明'])

    def save_to_supabase(df):
        with conn.session as session:
            # 使用 DROP TABLE 可以確保資料庫結構(Schema)跟著我們的新欄位一起更新
            session.execute(text("DROP TABLE IF EXISTS financials")) 
            df.to_sql('financials', conn.engine, if_exists='replace', index=False)
            session.commit()

    def process_excel_with_colors(uploaded_file):
        """讀取 Excel 並抓取儲存格底色"""
        # 1. 使用 openpyxl 讀取顏色
        wb = openpyxl.load_workbook(uploaded_file, data_only=True)
        sheet = wb.active
        
        color_data = []
        for row in sheet.iter_rows(min_row=4):
            # 抓取「專案說明」那一格的顏色
            cell_color = row[1].fill.start_color.index
            color_data.append(str(cell_color) if cell_color != '00000000' else "無底色")

        # 2. 回到 pandas 讀取數值
        uploaded_file.seek(0)
        df = pd.read_excel(uploaded_file, skiprows=2)
        
        df['顏色標記'] = color_data[:len(df)]
        df.rename(columns={df.columns[2]: '紀錄類型'}, inplace=True)
        df['專案說明'] = df['專案說明'].ffill()
        
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
                
        # 【修復1】修正了這裡的錯字 (原本寫成 顏色標標記)
        target_cols = ['專案說明', '紀錄類型'] + months + ['營收分類', '顏色標記', '說明']
        return df[df['專案說明'].notna()][target_cols]

    # ==========================================
    # 2. 介面呈現
    # ==========================================
    tab_edit, tab_summary = st.tabs(["📝 數據編輯與匯入", "📈 顏色加總與分析"])

    with tab_edit:
        with st.sidebar:
            st.header("📂 數據匯入")
            file = st.file_uploader("匯入帶有底色的 Excel", type=["xlsx"])
            if file and st.button("🚀 開始解析顏色與數據"):
                with st.spinner("正在解析顏色與上傳資料庫..."):
                    processed_df = process_excel_with_colors(file)
                    save_to_supabase(processed_df)
                    st.success("匯入完成！資料庫結構已更新。")
                    st.rerun()
            
            st.divider()
            if st.button("⚠️ 清空資料庫 (重來)"):
                with conn.session as session:
                    session.execute(text("DROP TABLE IF EXISTS financials"))
                    session.commit()
                st.warning("資料庫已清空，請重新匯入檔案")
                st.rerun()

        data = load_data()
        st.subheader("📝 營收明細編輯 (支援刪除列)")
        st.caption("提示：點擊最左側選取列，按下鍵盤 Delete 鍵可刪除。")
        
        edited_df = st.data_editor(
            data,
            num_rows="dynamic",
            use_container_width=True,
            height=500,
            column_config={
                "營收分類": st.column_config.SelectboxColumn("營收分類", options=["24DCM開發/維運", "TOYOTA聯網服務", "LEXUS聯網服務", "其他"]),
                "紀錄類型": st.column_config.SelectboxColumn("紀錄類型", options=["收入", "支出", "收入預估", "支出預估"]),
                "顏色標記": st.column_config.TextColumn("Excel底色代碼", disabled=True),
                **{m: st.column_config.NumberColumn(format="%.0f") for m in ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec']}
            }
        )
        if st.button("💾 儲存變更", type="primary"):
            save_to_supabase(edited_df)
            st.success("雲端存儲成功！")

    with tab_summary:
        st.header("🎨 基於 Excel 底色的自動加總分析")
        if not edited_df.empty:
            months = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec']
            df_sum = edited_df.copy()
            
            # 【修復3】確保 df_sum 裡面一定有顏色標記欄位
            if '顏色標記' not in df_sum.columns:
                df_sum['顏色標記'] = '無底色'
                
            for m in months:
                df_sum[m] = pd.to_numeric(df_sum[m], errors='coerce').fillna(0)
            df_sum['年度總計'] = df_sum[months].sum(axis=1)

            # 針對顏色進行加總
            try:
                color_summary = df_sum.groupby(['顏色標記', '紀錄類型'])['年度總計'].sum().unstack().fillna(0)
                st.subheader("📊 不同顏色區塊的營收規模")
                st.dataframe(color_summary.style.format("{:,.0f}"), use_container_width=True)
            except Exception as e:
                st.error(f"顏色加總計算錯誤: {e}")
            
            # 差異 Pick-up
            st.divider()
            st.subheader("🔍 異常檢視 (目標 vs 預估)")
            diff_list = []
            for proj in df_sum['專案說明'].unique():
                p_data = df_sum[df_sum['專案說明'] == proj]
                t_val = p_data[p_data['紀錄類型'] == '收入']['年度總計'].sum()
                e_val = p_data[p_data['紀錄類型'] == '收入預估']['年度總計'].sum()
                if t_val != e_val:
                    diff_list.append({
                        "專案名稱": proj,
                        "分類": p_data['營收分類'].iloc[0] if not p_data['營收分類'].empty else "其他",
                        "顏色": p_data['顏色標記'].iloc[0] if not p_data['顏色標記'].empty else "無底色",
                        "目標": t_val,
                        "預估": e_val,
                        "差異": t_val - e_val
                    })
            if diff_list:
                st.table(pd.DataFrame(diff_list))
            else:
                st.success("✅ 目前所有專案目標與預估皆一致。")
