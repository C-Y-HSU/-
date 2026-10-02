import streamlit as st
import openpyxl
import pandas as pd
import tempfile
import os

# 1. 網頁標題與選單設定
st.set_page_config(page_title="浩然敬老院 - 院民人日數自動化工具", page_icon="🏥", layout="wide")

st.title("🏥 浩然敬老院 - 院民人日數自動化分析系統")
st.markdown("請在左側輸入統計月份並上傳每日報表（如 `0601.xlsx` ~ `0630.xlsx`），系統將自動計算各區人數與全月總人日數。")

# 2. 側邊欄控制區
with st.sidebar:
    st.header("📌 設定與檔案上傳")
    month_num = st.number_input("統計月份 (1~12)", min_value=1, max_value=12, value=6, step=1)
    uploaded_files = st.file_uploader(
        "📁 選擇/拖曳多個每日報表 (.xlsx)", 
        type=["xlsx"], 
        accept_multiple_files=True
    )
    btn_calculate = st.button("🚀 開始自動化分析", type="primary")

# 3. 核心計算邏輯
if btn_calculate:
    if not uploaded_files:
        st.warning("⚠️ 請先上傳每日報表 Excel 檔案！")
    else:
        target_cols = ['E', 'F', 'G', 'I', 'J', 'M', 'N', 'O']

        def get_cell_val(ws, coord):
            val = ws[coord].value
            if val is None:
                return 0
            try:
                return float(val)
            except (ValueError, TypeError):
                return 0

        # 依檔名排序 (0601.xlsx, 0602.xlsx...)
        sorted_files = sorted(uploaded_files, key=lambda x: x.name)
        all_monthly_rows = []

        with st.spinner("正在進行大數據解析與計算中..."):
            for file_obj in sorted_files:
                file_name = file_obj.name
                
                # 排除非預期的舊檔
                if file_name.startswith('~$') or '彙總表' in file_name or '總表' in file_name:
                    continue
                    
                wb = openpyxl.load_workbook(file_obj, data_only=True)
                ws = wb.active
                
                row_data = {"日期/檔名": file_name}
                daily_total_net = 0
                
                for col in target_cols:
                    col_16 = get_cell_val(ws, f"{col}16")
                    col_outs = sum(get_cell_val(ws, f"{col}{r}") for r in range(17, 23))
                    net_val = col_16 - col_outs
                    row_data[f"{col}欄_實際人數"] = int(net_val)
                    daily_total_net += net_val
                
                row_data["全院合計實際人數"] = int(daily_total_net)
                all_monthly_rows.append(row_data)

        if not all_monthly_rows:
            st.error("❌ 沒有找到有效的每日報表檔案，請確認檔名與格式！")
        else:
            df_summary = pd.DataFrame(all_monthly_rows)

            column_rename_map = {
                "E欄_實際人數": "致中組_中一",
                "F欄_實際人數": "致中組_中二",
                "G欄_實際人數": "致中組_中三",
                "I欄_實際人數": "致和組_三樓",
                "J欄_實際人數": "致和組_四樓",
                "M欄_實際人數": "保養組_1區",
                "N欄_實際人數": "保養組_2區",
                "O欄_實際人數": "保養組_3區",
            }
            df_summary.rename(columns=column_rename_map, inplace=True)

            # 全月人日數加總
            sum_row = {"日期/檔名": f"【{int(month_num)}月全月總人日數】"}
            for col_name in df_summary.columns:
                if col_name != "日期/檔名":
                    sum_row[col_name] = df_summary[col_name].sum()

            df_summary = pd.concat([df_summary, pd.DataFrame([sum_row])], ignore_index=True)

            # 4. 畫面上顯示成果
            st.success("🎉 分析完成！")
            st.subheader("📊 統計成果預覽 (含每日人數與全月人日數)")
            st.dataframe(df_summary, use_container_width=True)

            # 生成 Excel 下載檔
            temp_dir = tempfile.mkdtemp()
            out_file_path = os.path.join(temp_dir, f"{int(month_num)}月全院各區實際在院人數及人日數總彙總表.xlsx")
            df_summary.to_excel(out_file_path, index=False)

            with open(out_file_path, "rb") as f:
                st.download_button(
                    label="📥 下載全月彙總 Excel 檔",
                    data=f,
                    file_name=f"{int(month_num)}月全院各區實際在院人數及人日數總彙總表.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
