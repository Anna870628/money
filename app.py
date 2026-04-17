import streamlit as st
import pandas as pd
from sqlalchemy import text
import openpyxl

# ==========================================
# 0. 安全驗證 (密碼: CMX_BPT)
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
    elif not st.session_state.get("password_correct", False):
        st.title("🔒 營收管理系統 - 身份驗證")
        st.text_input("請輸入存取密碼", type="password", on_change=password_entered, key="password")
        st.error("😕 密碼錯誤，請重新輸入。")
        return False
    return True

# ==========================================
# 1. 核心邏輯區
# ==========================================
if check_password():
    st.set_page_config(page_title="車聯網營收戰情系統", layout="wide")
    conn = st.connection("postgresql", type="sql")

    def load_data():
        """讀取雲端資料，若無資料則回傳空架構"""
        try:
            df = conn.query("SELECT * FROM financials", ttl="0")
            months = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec']
            for m in months:
                if m in df.columns:
                    df[m] = pd.to_numeric(df[m], errors='coerce').fillna(0)
            return df
        except:
            return pd.DataFrame(columns=['專案說明', '紀錄類型', '營收分類', '顏色標記', '說明'] + ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'])

    def save_to_supabase(df):
        """極速寫入引擎：使用 DELETE 避開死結，並批量上傳"""
        with conn.session as session:
            # 建立表格(若不存在)，或清空現有資料列
            session.execute(text("CREATE TABLE IF NOT EXISTS financials (id SERIAL PRIMARY KEY)")) 
            session.execute(text("DELETE FROM financials"))
            session.commit()
        # 使用 append 模式配合 method='multi' 提升 10 倍速度
        df.to_sql('financials', conn.engine, if_exists='append', index=False, chunksize=500, method='multi')

    def process_imported_file(uploaded_file):
        """讀取 Excel、辨識顏色、清洗數據、填補合併儲存格"""
        # 1. 讀取顏色 (使用 openpyxl)
        wb = openpyxl.load_workbook(uploaded_file, data_only=True)
        # 尋找數據分頁
        sheet_names = wb.sheetnames
        target_s = sheet_names[0]
        for s in sheet_names:
            if any(k in s for k in ['營收', '預估', '收支']):
                target_s = s
                break
        sheet = wb[target_s]
        
        color_list = []
        # 假設資料從第 4 行開始 (對應 skiprows=2)
        for row in sheet.iter_rows(min_row=4):
            # 抓取專案說明格(Column B)的背景色
            fill = row[1].fill
            color_index = str(fill.start_color.index) if fill and fill.start_color else "無底色"
            color_list.append(color_index)

        # 2. 讀取數據 (使用 pandas)
        uploaded_file.seek(0)
        df = pd.read_excel(uploaded_file, sheet_name=target_s, skiprows=2)
        df.columns = [str(c).strip() for c in df.columns]
        
        # 數據清理
        df['顏色標記'] = color_list[:len(df)]
        df.rename(columns={df.columns[2]: '紀錄類型'}, inplace=True)
        df['專案說明'] = df['專案說明'].replace(r'^\s*$', pd.NA, regex=True).ffill()
        
        cat_col = [c for c in df.columns if '營收分類' in str(c)]
        if cat_col:
            df.rename(columns={cat_col[0]: '營收分類'}, inplace=True)
            df['營收分類'] = df['營收分類'].ffill().fillna("其他")
        
        months = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec']
        for m in months:
            if m in df.columns:
                df[m] = pd.to_numeric(df[m], errors='coerce').fillna(0)
            else:
                df[m] = 0.0
        
        if '說明' not in df.columns: df['說明'] = ""
            
        target_cols = ['專案說明', '紀錄類型'] + months + ['營收分類', '顏色標記', '說明']
        return df.dropna(subset=['專案說明', '紀錄類型'])[target_cols]

    # ==========================================
    # 2. 介面呈現
    # ==========================================
    st.title("📊 車聯網事業本部 - 專案營收戰情室")
    
    tabs = st.tabs(["🎴 專案卡片摘要", "📝 原始數據管理", "🎨 顏色加總分析"])
    data = load_data()

    # --- TAB 1: 卡片戰情室 ---
    with tabs[0]:
        if data.empty:
            st.warning("請先至數據管理分頁匯入資料。")
        else:
            projects = data['專案說明'].unique()
            cols = st.columns(2)
            for i, p in enumerate(projects):
                with cols[i % 2]:
                    p_df = data[data['專案說明'] == p]
                    months = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec']
                    
                    # 計算核心
                    target = p_df[p_df['紀錄類型'] == '收入'][months].sum().sum()
                    est_in = p_df[p_df['紀錄類型'] == '收入預估'][months].sum().sum()
                    est_out = p_df[p_df['紀錄類型'] == '支出預估'][months].sum().sum()
                    
                    profit = est_in - est_out
                    margin = (profit / est_in) if est_in != 0 else 0
                    
                    with st.container(border=True):
                        st.markdown(f"#### {p}")
                        st.caption(f"分類：{p_df['營收分類'].iloc[0]} | 標籤：{p_df['顏色標記'].iloc[0]}")
                        m1, m2, m3 = st.columns(3)
                        m1.metric("目標營收", f"${target:,.0f}")
                        m2.metric("預估營收", f"${est_in:,.0f}", f"{est_in-target:,.0f}")
                        m3.metric("預估毛利率", f"{margin:.1%}")
                        
                        reach = (est_in / target) if target != 0 else 0
                        st.write(f"**目標達成率: {reach:.1%}**")
                        st.progress(min(reach, 1.0))
                        with st.expander("查看 1-12 月數據"):
                            st.dataframe(p_df, use_container_width=True)

    # --- TAB 2: 數據管理 ---
    with tabs[1]:
        with st.sidebar:
            st.header("📂 匯入數據")
            f = st.file_uploader("選擇 Excel", type=["xlsx"])
            if f and st.button("🚀 開始解析並上傳"):
                try:
                    new_df = process_imported_file(f)
                    save_to_supabase(new_df)
                    st.success("匯入成功！")
                    st.rerun()
                except Exception as e:
                    st.error(f"解析失敗: {e}")
            
            st.divider()
            if st.button("⚠️ 清空資料庫"):
                with conn.session as s:
                    s.execute(text("DELETE FROM financials"))
                    s.commit()
                st.rerun()

        st.info("提示：此處可直接編輯，記得點擊下方儲存按鈕。")
        edited = st.data_editor(
            data, num_rows="dynamic", use_container_width=True, height=500,
            column_config={
                "營收分類": st.column_config.SelectboxColumn(options=["24DCM開發/維運", "TOYOTA聯網服務", "LEXUS聯網服務", "其他"]),
                "紀錄類型": st.column_config.SelectboxColumn(options=["收入", "收入預估", "支出", "支出預估", "收入差異", "支出差異"])
            }
        )
        if st.button("💾 儲存變更", type="primary"):
            save_to_supabase(edited)
            st.success("已更新至雲端！")

    # --- TAB 3: 顏色加總 ---
    with tabs[2]:
        if not data.empty:
            df_sum = data.copy()
            months = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec']
            df_sum['年度總計'] = df_sum[months].sum(axis=1)
            
            st.subheader("📊 依 Excel 底色加總營收規模")
            color_res = df_sum.groupby(['顏色標記', '紀錄類型'])['年度總計'].sum().unstack().fillna(0)
            st.dataframe(color_res.style.format("{:,.0f}"), use_container_width=True)
            
            st.divider()
            st.subheader("🔍 異常明細 (目標 ≠ 預估)")
            diff_list = []
            for p in df_sum['專案說明'].unique():
                pdf = df_sum[df_sum['專案說明'] == p]
                t = pdf[pdf['紀錄類型'] == '收入']['年度總計'].sum()
                e = pdf[pdf['紀錄類型'] == '收入預估']['年度總計'].sum()
                if t != e:
                    diff_list.append({"專案": p, "分類": pdf['營收分類'].iloc[0], "目標": t, "預估": e, "差異": t-e})
            if diff_list: st.table(pd.DataFrame(diff_list))
