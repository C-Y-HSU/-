from datetime import datetime
import io
import os

from docxtpl import DocxTemplate
from google.oauth2.service_account import Credentials
import gspread
import pandas as pd
import streamlit as st

st.title("🏥 群聚事件管理系統 - 雲端多事件與通報單生成")
st.write(
    "支援住民資料庫、智慧帶入、雲端分頁同步，以及一鍵生成正式 Word 通報單功能。"
)

# 定義群聚事件日誌的標準欄位順序
DESIRED_COLS = [
    "事件名稱",
    "事件狀態",
    "指標個案",
    "發生日期",
    "姓名",
    "性別",
    "生日",
    "身份證字號",
    "確診管道",
    "後續處理",
    "記錄時間",
]

# 定義住民基本資料庫的欄位順序
RESIDENT_COLS = ["姓名", "房床號", "身份證字號", "出生日期", "入住日期"]


# ==========================================
# 雲端資料庫連線與讀寫函數
# ==========================================
@st.cache_resource
def init_google_spreadsheet():
  creds_dict = dict(st.secrets["gcp_service_account"])
  scope = [
      "https://www.googleapis.com/auth/spreadsheets",
      "https://www.googleapis.com/auth/drive",
  ]
  creds = Credentials.from_service_account_info(creds_dict, scopes=scope)
  client = gspread.authorize(creds)

  sheet_id = "16EGDmPEQnhjhYrJ3Y5Afc-s1ByM3Va7eHkBlCXSIBUU"
  return client.open_by_key(sheet_id)


def load_residents():
  try:
    spreadsheet = init_google_spreadsheet()
    try:
      sheet = spreadsheet.worksheet("住民基本資料")
    except gspread.exceptions.WorksheetNotFound:
      sheet = spreadsheet.add_worksheet(
          title="住民基本資料", rows=100, cols=10
      )
      sheet.update([RESIDENT_COLS])

    data = sheet.get_all_records()
    if not data:
      return pd.DataFrame(columns=RESIDENT_COLS)
    df = pd.DataFrame(data)
    for col in RESIDENT_COLS:
      if col not in df.columns:
        df[col] = ""
    return df
  except Exception as e:
    return pd.DataFrame(columns=RESIDENT_COLS)


def save_residents(df_res):
  spreadsheet = init_google_spreadsheet()
  try:
    sheet = spreadsheet.worksheet("住民基本資料")
  except gspread.exceptions.WorksheetNotFound:
    sheet = spreadsheet.add_worksheet(title="住民基本資料", rows=100, cols=10)

  sheet.clear()
  existing_cols = [col for col in RESIDENT_COLS if col in df_res.columns]
  df_res = df_res[existing_cols]
  sheet.update([df_res.columns.values.tolist()] + df_res.values.tolist())


def load_data():
  try:
    spreadsheet = init_google_spreadsheet()
    sheets = spreadsheet.worksheets()
    all_dfs = []

    for sheet in sheets:
      if sheet.title == "住民基本資料":
        continue
      data = sheet.get_all_records()
      if data:
        df = pd.DataFrame(data)
        df["事件名稱"] = sheet.title
        all_dfs.append(df)

    if not all_dfs:
      return pd.DataFrame(columns=DESIRED_COLS)

    df_logs = pd.concat(all_dfs, ignore_index=True)

    for col in DESIRED_COLS:
      if col not in df_logs.columns:
        df_logs[col] = ""

    df_logs["事件名稱"] = df_logs["事件名稱"].fillna("未命名群聚事件").astype(str)
    df_logs["事件狀態"] = df_logs["事件狀態"].fillna("進行中").astype(str)
    df_logs["指標個案"] = df_logs["指標個案"].fillna("").astype(str)

    existing_cols = [col for col in DESIRED_COLS if col in df_logs.columns]
    other_cols = [col for col in df_logs.columns if col not in DESIRED_COLS]
    return df_logs[existing_cols + other_cols]

  except Exception as e:
    st.error(f"⚠️ 無法讀取 Google 試算表，錯誤原因：{e}")
    return pd.DataFrame(columns=DESIRED_COLS)


def save_event_data(event_name, full_df):
  spreadsheet = init_google_spreadsheet()
  df_event = full_df[full_df["事件名稱"] == event_name]

  try:
    sheet = spreadsheet.worksheet(event_name)
  except gspread.exceptions.WorksheetNotFound:
    sheet = spreadsheet.add_worksheet(title=event_name, rows=100, cols=20)

  sheet.clear()
  if not df_event.empty:
    existing_cols = [col for col in DESIRED_COLS if col in df_event.columns]
    df_event = df_event[existing_cols]
    sheet.update(
        [df_event.columns.values.tolist()] + df_event.values.tolist()
    )
  else:
    sheet.update([DESIRED_COLS])


df_logs = load_data()
df_residents = load_residents()

# 建立對應字典，方便從身分證查出房床號
resident_room_map = {}
if not df_residents.empty:
  for _, r in df_residents.iterrows():
    id_key = str(r.get("身份證字號", "")).strip().upper()
    room_val = str(r.get("房床號", "")).strip()
    if id_key:
      resident_room_map[id_key] = room_val

# ==========================================
# 側邊欄：事件管理與住民資料庫維護
# ==========================================
st.sidebar.header("📁 事件管理與結案中心")

if not df_logs.empty and "事件名稱" in df_logs.columns:
  event_status_map = (
      df_logs.drop_duplicates(subset=["事件名稱"])
      .set_index("事件名稱")["事件狀態"]
      .to_dict()
  )
else:
  event_status_map = {}

all_events = list(event_status_map.keys())
active_events = [
    e for e, status in event_status_map.items() if status != "已結案"
]
closed_events = [
    e for e, status in event_status_map.items() if status == "已結案"
]

event_mode = st.sidebar.radio("請選擇操作模式", ["進行中事件", "新增並切換新事件"])

selected_event = ""
if event_mode == "進行中事件":
  view_closed = st.sidebar.checkbox("📂 顯示已結案的歷史事件")
  target_list = (
      (active_events + closed_events) if view_closed else active_events
  )

  if target_list:
    selected_event = st.sidebar.selectbox("選擇要處理的事件", target_list)
  else:
    selected_event = st.sidebar.text_input(
        "目前無進行中事件，請輸入新事件名稱", "預設群聚事件"
    )
else:
  selected_event = st.sidebar.text_input("輸入新事件名稱（例如：A機構群聚）", "")

if not selected_event:
  selected_event = "未命名群聚事件"

current_status = event_status_map.get(selected_event, "進行中")

st.sidebar.markdown("---")
st.sidebar.markdown(f"**📌 目前選定事件：**\n### `{selected_event}`")
if current_status == "已結案":
  st.sidebar.error("🔒 狀態：此事件已結案（唯獨封存）")
else:
  st.sidebar.success("🔥 狀態：進行中")

if not df_logs.empty and selected_event in event_status_map:
  st.sidebar.markdown("---")
  if current_status == "進行中":
    if st.sidebar.button("🔒 將此事件標記為結案"):
      df_logs.loc[df_logs["事件名稱"] == selected_event, "事件狀態"] = "已結案"
      save_event_data(selected_event, df_logs)
      st.sidebar.success("已成功將此事件結案並同步至專屬分頁！")
      st.rerun()
  else:
    if st.sidebar.button("🔓 重新啟動此事件"):
      df_logs.loc[df_logs["事件名稱"] == selected_event, "事件狀態"] = "進行中"
      save_event_data(selected_event, df_logs)
      st.sidebar.success("已成功重新啟動此事件！")
      st.rerun()

with st.sidebar.expander("👥 管理機構住民基本資料庫"):
  st.write("依序維護：姓名、房床號、身分證字號、出生日期、入住日期")
  edited_res_df = st.data_editor(
      df_residents,
      num_rows="dynamic",
      use_container_width=True,
      key="resident_editor",
  )
  if st.button("💾 儲存住民名冊至雲端"):
    save_residents(edited_res_df)
    st.success("✨ 住民基本資料已成功更新！")
    st.rerun()

if not df_logs.empty and "事件名稱" in df_logs.columns:
  df_current_event = df_logs[df_logs["事件名稱"] == selected_event].copy()

  if not df_current_event.empty and "指標個案" in df_current_event.columns:
    df_current_event["_temp_sort"] = df_current_event[
        "指標個案"
    ].str.contains("指標個案", na=False)
    df_current_event = df_current_event.sort_values(
        by="_temp_sort", ascending=False
    ).drop(columns=["_temp_sort"])
else:
  df_current_event = pd.DataFrame(columns=DESIRED_COLS)


# ==========================================
# 醒目功能：顯示當前事件的「指標個案英雄卡片」
# ==========================================
st.markdown("---")
if not df_current_event.empty and "指標個案" in df_current_event.columns:
  index_cases = df_current_event[
      df_current_event["指標個案"].str.contains("指標個案", na=False)
  ]

  if not index_cases.empty:
    ic = index_cases.iloc[0]
    st.success(
        f"### 🌟 【{selected_event}】之官方認定指標個案\n"
        f"- **姓名**：{ic['姓名']} （性別：{ic['性別']}）\n"
        f"- **身分證字號**：{ic['身份證字號']}\n"
        f"- **發生日期**：{ic['發生日期']} | **確診管道**：{ic['確診管道']}\n"
        f"- **後續處理**：{ic['後續處理']}"
    )
  else:
    st.info(
        f"💡 提示：目前事件【{selected_event}】尚未指定指標個案，請於下方清單中選取。"
    )


# ==========================================
# 主畫面 1：新增個案
# ==========================================
if current_status == "已結案":
  st.warning(
      "🔒 此事件目前為【已結案】狀態，若需新增或修改個案，請先至側邊欄點擊「"
      "重新啟動此事件」。"
  )
else:
  st.subheader(f"📝 新增確診者（事件：{selected_event}）")

  selected_resident_key = "-- 手動輸入 / 不從名冊帶入 --"
  if not df_residents.empty:
    resident_options = {
        f"[{r.get('房床號', '無房號')}] {r['姓名']} ({r['身份證字號']})": r
        for _, r in df_residents.iterrows()
        if str(r.get("姓名", "")).strip()
    }
    if resident_options:
      selected_resident_key = st.selectbox(
          "🔍 【智慧搜尋帶入】輸入房號、姓名或身分證字號關鍵字篩選住民",
          ["-- 手動輸入 / 不從名冊帶入 --"] + list(resident_options.keys()),
      )

  default_name = ""
  default_birthday = ""
  default_id = ""

  if selected_resident_key != "-- 手動輸入 / 不從名冊帶入 --":
    res_data = resident_options[selected_resident_key]
    default_name = str(res_data.get("姓名", ""))
    default_birthday = str(res_data.get("出生日期", ""))
    default_id = str(res_data.get("身份證字號", ""))

  with st.form("case_form"):
    case_date = st.date_input("1. 發生日期", value=datetime.today())

    col1, col2 = st.columns(2)
    with col1:
      name = st.text_input("2. 姓名", value=default_name)
      gender = st.selectbox("3. 性別", ["男", "女", "其他"])

    with col2:
      birthday = st.text_input(
          "4. 生日／出生日期",
          value=default_birthday,
          placeholder="例如：1991-03-08",
      )
      id_number = st.text_input(
          "5. 身份證字號", value=default_id, placeholder="例如：A123456789"
      )

    diagnosis_method = st.text_input(
        "6. 確診管道（例如：快篩陽性、發燒就醫、PCR）"
    )
    follow_up_action = st.text_area(
        "7. 後續處理（例如：安排單人隔離、通報疾管科）"
    )

    submitted = st.form_submit_button("送出並記錄個案日誌")

    if submitted:
      if not name or not id_number:
        st.warning("⚠️ 請務必填寫「姓名」與「身份證字號」！")
      else:
        now_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        new_data = pd.DataFrame([{
            "事件名稱": selected_event,
            "事件狀態": "進行中",
            "指標個案": "",
            "發生日期": str(case_date),
            "姓名": name,
            "性別": gender,
            "生日": birthday,
            "身份證字號": id_number.upper(),
            "確診管道": diagnosis_method,
            "後續處理": follow_up_action,
            "記錄時間": now_time,
        }])

        df_logs = pd.concat([df_logs, new_data], ignore_index=True)
        save_event_data(selected_event, df_logs)
        st.success(
            f"🎉 成功將確診者【{name}】記錄至雲端分頁【{selected_event}】！"
        )
        st.rerun()


# ==========================================
# 主畫面 2：目前事件的總日誌表格與通報單生成
# ==========================================
st.markdown("---")
st.subheader(f"📋 【{selected_event}】目前的確診個案總日誌")

if not df_current_event.empty:
  df_display = df_current_event.reset_index(drop=True).copy()
  df_display.index = df_display.index + 1
  st.dataframe(df_display, use_container_width=True)

  # ------------------------------------------
  # 獨立區塊：一鍵下載 Word 通報單
  # ------------------------------------------
  st.markdown("---")
  st.markdown("#### 📄 匯出正式 Word 通報單")
  report_options = {}
  for local_num, (idx, row) in enumerate(df_current_event.iterrows(), start=1):
    label = f"個案編號 {local_num}：{row['姓名']} ({row['身份證字號']})"
    report_options[label] = row

  selected_report_label = st.selectbox(
      "選擇要生成 Word 通報單的個案",
      list(report_options.keys()),
      key="report_select",
  )

  if st.button("📥 點擊生成此個案的 Word 通報單"):
    row_data = report_options[selected_report_label]
    id_upper = str(row_data["身份證字號"]).strip().upper()
    room_no = resident_room_map.get(id_upper, "未建檔房號")

    context = {
        "姓名": str(row_data["姓名"]),
        "房床號": room_no,
        "身份證字號": str(row_data["身份證字號"]),
        "生日": str(row_data["生日"]),
        "發生日期": str(row_data["發生日期"]),
        "確診管道": str(row_data["確診管道"]),
        "後續處理": str(row_data["後續處理"]),
    }

    template_path = "template.docx"
    if os.path.exists(template_path):
      try:
        doc = DocxTemplate(template_path)
        doc.render(context)

        file_stream = io.BytesIO()
        doc.save(file_stream)
        file_stream.seek(0)

        st.download_button(
            label=f"💾 下載 【{row_data['姓名']}】 的正式通報單.docx",
            data=file_stream,
            file_name=f"通報單_{row_data['姓名']}_{row_data['身份證字號']}.docx",
            mime=(
                "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            ),
        )
        st.success("✅ 通報單已成功生成，請點擊上方出現的下載按鈕！")
      except Exception as e:
        st.error(f"⚠️ 生成 Word 檔案時發生錯誤：{e}")
    else:
      st.error(
          "⚠️ 找不到 Word 範本檔案 (`template.docx`)，請確認是否已經上傳至 GitHub"
          " 專案根目錄！"
      )

  if current_status != "已結案":
    # ------------------------------------------
    # 子功能 A：設定指標個案
    # ------------------------------------------
    st.markdown("---")
    st.markdown("#### ⭐ 設定此事件的指標個案")
    case_options = {}
    for local_num, (idx, row) in enumerate(
        df_current_event.iterrows(), start=1
    ):
      label = f"個案編號 {local_num}：{row['姓名']} ({row['身份證字號']})"
      case_options[label] = idx

    selected_case_label = st.selectbox(
        "選擇要設為指標個案的確診者",
        list(case_options.keys()),
        key="index_select",
    )

    if st.button("🌟 確認將此人設為指標個案"):
      target_idx = case_options[selected_case_label]

      df_logs["指標個案"] = df_logs["指標個案"].astype(str)
      df_logs.loc[df_logs["事件名稱"] == selected_event, "指標個案"] = ""
      df_logs.loc[target_idx, "指標個案"] = "⭐ 指標個案"

      save_event_data(selected_event, df_logs)
      st.success(
          f"✨ 已成功指定【{selected_case_label}】為本群聚事件的指標個案，並同步至雲端分頁！"
      )
      st.rerun()

    # ------------------------------------------
    # 子功能 B：修改或刪除現有個案資料
    # ------------------------------------------
    st.markdown("---")
    st.markdown("#### ✏️ 修改或刪除現有個案資料")

    edit_options = {}
    for local_num, (idx, row) in enumerate(df_current_event.iterrows(), start=1):
      label = f"個案編號 {local_num}：{row['姓名']} ({row['身份證字號']})"
      edit_options[label] = idx

    selected_edit_label = st.selectbox(
        "選擇要修改或刪除的個案", list(edit_options.keys()), key="edit_select"
    )

    if selected_edit_label:
      target_idx = edit_options[selected_edit_label]
      target_row = df_logs.loc[target_idx]

      with st.form("edit_case_form"):
        st.markdown(
            f"正在編輯：**{target_row['姓名']}** （身分證："
            f"`{target_row['身份證字號']}`）"
        )

        try:
          default_date = datetime.strptime(
              str(target_row["發生日期"]), "%Y-%m-%d"
          ).date()
        except:
          default_date = datetime.today().date()

        edit_case_date = st.date_input("1. 發生日期", value=default_date)

        col_e1, col_e2 = st.columns(2)
        with col_e1:
          edit_name = st.text_input("2. 姓名", value=str(target_row["姓名"]))
          genders = ["男", "女", "其他"]
          current_g = str(target_row["性別"])
          g_idx = genders.index(current_g) if current_g in genders else 0
          edit_gender = st.selectbox("3. 性別", genders, index=g_idx)
        with col_e2:
          edit_birthday = st.text_input("4. 生日", value=str(target_row["生日"]))
          edit_id = st.text_input(
              "5. 身份證字號", value=str(target_row["身份證字號"])
          )

        edit_diag = st.text_input(
            "6. 確診管道", value=str(target_row["確診管道"])
        )
        edit_follow = st.text_area(
            "7. 後續處理", value=str(target_row["後續處理"])
        )

        col_sub1, col_sub2 = st.columns(2)
        with col_sub1:
          update_submitted = st.form_submit_button("💾 儲存修改至雲端分頁")
        with col_sub2:
          delete_submitted = st.form_submit_button(
              "🗑️ 從雲端分頁刪除此個案"
          )

        if update_submitted:
          if not edit_name or not edit_id:
            st.warning("⚠️ 請務必填寫「姓名」與「身份證字號」！")
          else:
            df_logs.loc[target_idx, "發生日期"] = str(edit_case_date)
            df_logs.loc[target_idx, "姓名"] = edit_name
            df_logs.loc[target_idx, "性別"] = edit_gender
            df_logs.loc[target_idx, "生日"] = edit_birthday
            df_logs.loc[target_idx, "身份證字號"] = edit_id.upper()
            df_logs.loc[target_idx, "確診管道"] = edit_diag
            df_logs.loc[target_idx, "後續處理"] = edit_follow

            save_event_data(selected_event, df_logs)
            st.success(f"✨ 成功更新個案【{edit_name}】並同步至雲端分頁！")
            st.rerun()

        if delete_submitted:
          deleted_name = target_row["姓名"]
          deleted_id = target_row["身份證字號"]
          df_logs = df_logs.drop(target_idx).reset_index(drop=True)
          save_event_data(selected_event, df_logs)
          st.success(f"🗑️ 已成功刪除個案【{deleted_name} ({deleted_id})】！")
          st.rerun()

  else:
    st.info("🔒 此事件已結案，無法再變更或刪除個案。")

else:
  st.info(f"事件【{selected_event}】目前尚無個案紀錄！")

# 管理員總表
with st.expander("🔍 管理員視角：檢視所有事件的總合併日誌（雲端同步）"):
  st.dataframe(df_logs, use_container_width=True)
