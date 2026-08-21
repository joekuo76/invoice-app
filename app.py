import io
import json
from google import genai
from google.genai import types
import openpyxl
from openpyxl.styles import Alignment, Border, Side
from PIL import Image
import streamlit as st

st.set_page_config(page_title="請款單自動生成系統", layout="wide")
st.title("🧾 請款單自動生成工具")

# 側邊欄：設定 API Key
with st.sidebar:
  st.header("⚙️ 系統設定")
  api_key = st.text_input("請輸入 Gemini API Key", type="password")
  applicant_name = st.text_input("請款人姓名", value="")
  department = st.text_input("請款單位", value="")


def parse_receipt(image_bytes: bytes, client: genai.Client) -> dict:
  """呼叫 Gemini 辨識發票/收據照片"""
  image = Image.open(io.BytesIO(image_bytes))
  prompt = """
    你是一位專業的會計助理，請辨識此單據/發票/收據照片的資訊。
    請嚴格依照 JSON 格式輸出（不要輸出 Markdown 區塊或額外文字）：
    {
        "date": "YYYY-MM-DD (若年份為民國年請轉為西元年，無日期則留空)",
        "invoice_no": "發票號碼 (兩碼英文+八碼數字，若無則留空)",
        "summary": "付款內容摘要 (例如: 機車加油, 汽車停車, 工地飲料, 耗材)",
        "amount": 金額數字 (整數),
        "payment_method": "現金 或 刷卡"
    }
    """
  response = client.models.generate_content(
      model="gemini-2.5-flash",
      contents=[image, prompt],
      config=types.GenerateContentConfig(
          response_mime_type="application/json",
          temperature=0.1,
      ),
  )
  return json.loads(response.text)


def write_to_excel(
    items: list,
    template_path: str,
    output_path: str,
    dept: str,
    applicant: str,
):
  """將辨識資料寫入 Excel 請款單"""
  wb = openpyxl.load_workbook(template_path)
  ws = wb.active

  # 填入基本資訊
  if dept:
    ws["B3"] = dept

  start_row = 5
  thin_border = Border(
      left=Side(style="thin"),
      right=Side(style="thin"),
      top=Side(style="thin"),
      bottom=Side(style="thin"),
  )
  center_align = Alignment(horizontal="center", vertical="center")
  right_align = Alignment(horizontal="right", vertical="center")

  # 寫入每筆明細
  for idx, item in enumerate(items):
    row = start_row + idx
    ws.cell(
        row=row, column=1, value=item.get("date", "")
    ).alignment = center_align
    ws.cell(
        row=row, column=2, value=item.get("invoice_no", "")
    ).alignment = center_align
    ws.cell(
        row=row, column=3, value=item.get("summary", "")
    ).alignment = center_align
    ws.cell(
        row=row, column=4, value=item.get("amount", 0)
    ).alignment = right_align
    ws.cell(
        row=row, column=5, value=item.get("payment_method", "現金")
    ).alignment = center_align

    for col in range(1, 8):
      ws.cell(row=row, column=col).border = thin_border

  # 合計欄位
  total_row = start_row + len(items)
  ws.cell(row=total_row, column=1, value="合    計").alignment = (
      center_align
  )
  ws.cell(
      row=total_row, column=4, value=f"=SUM(D{start_row}:D{total_row-1})"
  ).alignment = right_align

  for col in range(1, 8):
    ws.cell(row=total_row, column=col).border = thin_border

  # 底部簽核列
  sign_row = total_row + 2
  roles = ["總經理", "", "出納", "", "會計", "主管", "請款人"]
  for col_idx, role in enumerate(roles, start=1):
    cell = ws.cell(row=sign_row, column=col_idx, value=role)
    cell.alignment = center_align

  if applicant:
    ws.cell(row=sign_row + 1, column=7, value=applicant).alignment = (
        center_align
    )

  wb.save(output_path)


# 檔案上傳介面
uploaded_files = st.file_uploader(
    "上傳發票或收據照片 (可多選)",
    type=["jpg", "jpeg", "png"],
    accept_multiple_files=True,
)

if uploaded_files:
  if not api_key:
    st.warning("⚠️ 請先在左側欄位輸入 Gemini API Key！")
  else:
    if st.button("🚀 開始辨識單據並生成請款單"):
      client = genai.Client(api_key=api_key)
      parsed_results = []

      progress_bar = st.progress(0)
      for i, uploaded_file in enumerate(uploaded_files):
        with st.spinner(f"正在辨識第 {i+1}/{len(uploaded_files)} 張照片..."):
          try:
            res = parse_receipt(uploaded_file.getvalue(), client)
            parsed_results.append(res)
          except Exception as e:
            st.error(f"檔案 {uploaded_file.name} 辨識失敗: {e}")
        progress_bar.progress((i + 1) / len(uploaded_files))

      if parsed_results:
        st.session_state["parsed_results"] = parsed_results
        st.success("🎉 辨識完成！")

# 編輯與校對區
if "parsed_results" in st.session_state:
  st.subheader("📋 辨識結果確認與微調")
  edited_data = st.data_editor(
      st.session_state["parsed_results"], num_rows="dynamic", use_container_width=True
  )

  output_filename = "請款單_產出.xlsx"
  if st.button("💾 確認產出 Excel"):
    write_to_excel(
        edited_data, "template.xlsx", output_filename, department, applicant_name
    )

    with open(output_filename, "rb") as f:
      st.download_button(
          label="📥 下載完成的請款單 Excel",
          data=f,
          file_name="請款單.xlsx",
          mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      )