import streamlit as st
import pandas as pd
import os

# 設定網頁標題與寬度
st.set_page_config(page_title="職評工具條碼盤點系統", layout="wide")
st.title("📋 職評工具條碼行動盤點系統 (Excel 直接讀取版)")

# 指定讀取桌面資料夾內的真正的 Excel 檔案
db_filename = '庫存表.xlsx'

@st.cache_data
def load_initial_data(filename):
    if os.path.exists(filename):
        try:
            # 讀取 Excel 的「庫存表」工作表
            # 如果您的工作表有名稱，可以在這裡指定 sheet_name='庫存表'
            df = pd.read_excel(filename, sheet_name=0)
            
            # 清理欄位名稱的空白
            df.columns = df.columns.str.strip()
            
            # 強制將條碼轉為字串並去除空白
            if '條碼編號' in df.columns:
                df['條碼編號'] = df['條碼編號'].astype(str).str.strip()
            else:
                st.error("❌ 找不到【條碼編號】欄位，請檢查 Excel 工作表中的標頭名稱是否正確。")
                return None
            
            # 初始化實際盤點欄位
            if '實際盤點' not in df.columns or df['實際盤點'].isnull().all():
                df['實際盤點'] = 0
            
            # 補齊因 Excel 合併儲存格導致空白的工具名稱與類別
            if '工具名稱' in df.columns:
                df['工具名稱'] = df['工具名稱'].ffill()
            if '工具類別' in df.columns:
                df['工具類別'] = df['工具類別'].ffill()
                
            return df
        except Exception as e:
            st.error(f"❌ 讀取 Excel 檔案時發生錯誤: {e}")
            return None
    return None

# 如果找不到檔案的防呆機制
if not os.path.exists(db_filename):
    st.warning(f"⚠️ 在目前資料夾中找不到 【{db_filename}】。")
    st.info("請確認：1. 您的 Excel 檔名是否確實為 `庫存表.xlsx`\n2. CMD 終端機路徑是否已切換至該資料夾。")
    uploaded_file = st.file_uploader("或者，您也可以直接把 Excel 檔拖曳到下方上傳：", type=["xlsx"])
    if uploaded_file is not None:
        with open("庫存表.xlsx", "wb") as f:
            f.write(uploaded_file.getbuffer())
        st.rerun()

# 初始化 Session State
if os.path.exists(db_filename) and 'inventory_df' not in st.session_state:
    st.session_state.inventory_df = load_initial_data(db_filename)
if 'scan_history' not in st.session_state:
    st.session_state.scan_history = []

# 開始渲染介面
if 'inventory_df' in st.session_state and st.session_state.inventory_df is not None:
    df = st.session_state.inventory_df

    # 介面排版
    col1, col2 = st.columns([1, 2])

    with col1:
        st.subheader("🔍 條碼掃描區")
        st.write("請將游標點擊下方輸入框，並使用掃描槍進行掃碼：")

        with st.form(key='barcode_form', clear_on_submit=True):
            barcode_input = st.text_input("條碼輸入 (或手動輸入後按 Enter)", key="barcode_field")
            submit_button = st.form_submit_button("送出")

        if barcode_input:
            scanned_code = barcode_input.strip().replace('*', '') # 清除前後的星號防呆
            
            # 比對條碼
            match = df[df['條碼編號'] == scanned_code]

            if not match.empty:
                idx = match.index[0]
                item_name = df.loc[idx, '工具名稱']
                item_content = df.loc[idx, '工具內容']
                sys_qty = df.loc[idx, '系統數量'] if pd.notna(df.loc[idx, '系統數量']) else 0
                
                # 實際盤點數量 + 1
                df.loc[idx, '實際盤點'] = int(df.loc[idx, '實際盤點']) + 1
                st.session_state.inventory_df = df
                
                st.success(f"✅ 掃描成功！\n\n**條碼**：{scanned_code}\n\n**品名**：{item_name} ({item_content})\n\n**目前累計盤點數**：{df.loc[idx, '實際盤點']} (系統庫存: {int(sys_qty)})")
                st.session_state.scan_history.insert(0, f"成功盤點：{scanned_code} - {item_name}({item_content})")
            else:
                st.error(f"❌ 錯誤：找不到此條碼【{scanned_code}】")
                st.session_state.scan_history.insert(0, f"❌ 掃描失敗：未知條碼 {scanned_code}")

        st.write("---")
        st.caption("📜 最近掃描紀錄：")
        for log in st.session_state.scan_history[:5]:
            st.text(log)

    with col2:
        st.subheader("📊 盤點即時狀態表")
        
        display_df = df.copy()
        display_df['系統數量'] = pd.to_numeric(display_df['系統數量']).fillna(0).astype(int)
        display_df['實際盤點'] = display_df['實際盤點'].astype(int)
        display_df['盤點盈虧'] = display_df['實際盤點'] - display_df['系統數量']
        
        # 自動篩選出有條碼的資料列
        display_df = display_df[display_df['條碼編號'].notna() & (display_df['條碼編號'] != 'nan') & (display_df['條碼編號'] != '')]
        
        # 彈性調整要顯示的欄位
        available_cols = [col for col in ['工具類別', '條碼編號', '工具名稱', '工具內容', '系統數量', '實際盤點', '盤點盈虧'] if col in display_df.columns]
        
        st.dataframe(display_df[available_cols], use_container_width=True, height=500)

        st.write("---")
        
        # 匯出盤點結果為 Excel 檔
        import io
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
            display_df[available_cols].to_excel(writer, index=False, sheet_name='盤點結果')
        
        st.download_button(
            label="💾 下載本次盤點結果報告 (Excel)",
            data=buffer.getvalue(),
            file_name='盤點結果報告.xlsx',
            mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        )
        
        if st.button("🔄 重設所有盤點數量為 0"):
            df['實際盤點'] = 0
            st.session_state.inventory_df = df
            st.session_state.scan_history = []
            st.rerun()