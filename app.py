import io
import os
import re
from datetime import datetime

import numpy as np
import openpyxl
import pandas as pd
import streamlit as st
from PIL import Image, ImageOps
import zxingcpp

st.set_page_config(page_title="請款單自動生成系統", page_icon="🧾", layout="wide")
st.title("🧾 請款單自動生成工具 V6")
st.caption("電子發票 QR｜代收繳費單／停車單條碼｜人工補登｜一鍵產生 Excel")
st.success("V6 輕量版：不使用 Gemini、不使用 EasyOCR、不使用 Torch。")

TEMPLATE_PATH = "template.xlsx"
MAX_ITEMS = 8

if "items_v6" not in st.session_state:
    st.session_state.items_v6 = []


# ---------- 共用 ----------
def load_image(data: bytes) -> Image.Image:
    img = Image.open(io.BytesIO(data))
    return ImageOps.exif_transpose(img).convert("RGB")


def decode_barcodes(img: Image.Image):
    """讀取 QR / Code128 / Code39 / EAN / ITF 等 zxing-cpp 支援格式。"""
    results = []
    try:
        decoded = zxingcpp.read_barcodes(np.array(img))
        for r in decoded:
            text = getattr(r, "text", "") or ""
            text = text.strip()
            if not text:
                continue

            fmt = str(getattr(r, "format", "Unknown"))
            results.append({
                "text": text,
                "format": fmt,
            })
    except Exception:
        pass
    return results


def roc_date_to_iso(v: str):
    digits = re.sub(r"\D", "", v or "")
    if len(digits) != 7:
        return ""
    try:
        y = int(digits[:3]) + 1911
        m = int(digits[3:5])
        d = int(digits[5:7])
        return datetime(y, m, d).strftime("%Y-%m-%d")
    except Exception:
        return ""


def parse_invoice_qr(text: str):
    """
    台灣電子發票左側 QR 常見前段：
    10碼發票號碼 + 7碼民國日期 + 4碼隨機碼
    + 8碼銷售額(hex) + 8碼總額(hex)
    """
    s = (text or "").replace("\n", "").strip()
    if len(s) < 37:
        return None

    invoice_no = s[:10].upper()
    roc_date = s[10:17]
    total_hex = s[29:37]

    if not re.fullmatch(r"[A-Z]{2}\d{8}", invoice_no):
        return None
    if not re.fullmatch(r"\d{7}", roc_date):
        return None

    try:
        amount = int(total_hex, 16)
    except Exception:
        return None

    date = roc_date_to_iso(roc_date)
    if not date:
        return None

    return {
        "document_type": "發票",
        "date": date,
        "reference_no": invoice_no,
        "summary": "",
        "amount": amount,
        "payment_method": "未知",
        "location": "",
        "plate_no": "",
        "barcode_raw": s,
        "source": "QR自動",
    }


def guess_date_from_text(text: str):
    """只從條碼內容本身猜日期，不讀取紙本印刷文字。"""
    t = re.sub(r"\s+", "", text or "")

    # YYYYMMDD
    for m in re.finditer(r"(?<!\d)(20\d{2})(0[1-9]|1[0-2])([0-2]\d|3[01])(?!\d)", t):
        try:
            return datetime(int(m.group(1)), int(m.group(2)), int(m.group(3))).strftime("%Y-%m-%d")
        except Exception:
            pass

    # ROC YYYMMDD
    for m in re.finditer(r"(?<!\d)(1\d{2})(0[1-9]|1[0-2])([0-2]\d|3[01])(?!\d)", t):
        date = roc_date_to_iso("".join(m.groups()))
        if date:
            return date

    return ""


def guess_amount_from_text(text: str):
    """
    保守模式：只有條碼文字明確含 amount/amt/金額 關鍵字時才抓，
    避免把帳號、停車編號誤判成金額。
    """
    t = text or ""
    patterns = [
        r"(?i)(?:amount|amt)\s*[:=]\s*([0-9]{1,7})",
        r"金額\s*[:：=]?\s*([0-9]{1,7})",
    ]
    for p in patterns:
        m = re.search(p, t)
        if m:
            try:
                return int(m.group(1))
            except Exception:
                pass
    return 0


def clean_reference(text: str):
    """把條碼原文整理成適合放 Excel 單號欄的內容。"""
    t = (text or "").strip()
    if len(t) <= 40:
        return t
    # 過長內容只取前 40 碼做顯示，完整內容仍保留 barcode_raw
    return t[:40]


def parse_general_barcode(decoded):
    """
    一般代收／停車單：
    不假裝知道各家私有格式，只把實際讀到的條碼帶入，
    日期／金額僅做保守猜測，最後讓使用者確認。
    """
    if not decoded:
        return None

    # 優先選最長的非 QR 條碼；若只有 QR 就用最長的
    ordered = sorted(
        decoded,
        key=lambda x: (("QRCode" not in x["format"]), len(x["text"])),
        reverse=True
    )
    primary = ordered[0]
    all_raw = " | ".join([f'{x["format"]}:{x["text"]}' for x in decoded])

    text_pool = " ".join(x["text"] for x in decoded)
    return {
        "document_type": "代收繳費單",
        "date": guess_date_from_text(text_pool),
        "reference_no": clean_reference(primary["text"]),
        "summary": "代收繳費",
        "amount": guess_amount_from_text(text_pool),
        "payment_method": "未知",
        "location": "",
        "plate_no": "",
        "barcode_raw": all_raw,
        "source": "條碼自動",
    }


def scan_document(data: bytes, preferred_type="自動判斷"):
    img = load_image(data)
    decoded = decode_barcodes(img)

    # 電子發票優先
    for d in decoded:
        inv = parse_invoice_qr(d["text"])
        if inv:
            return inv, decoded

    if decoded:
        item = parse_general_barcode(decoded)
        if preferred_type == "停車繳費單":
            item["document_type"] = "停車繳費單"
            item["summary"] = "停車費"
        elif preferred_type == "收據":
            item["document_type"] = "收據"
            item["summary"] = ""
        elif preferred_type == "其他":
            item["document_type"] = "其他"
            item["summary"] = ""
        return item, decoded

    return None, []


def normalize_item(x):
    x = dict(x)
    x["document_type"] = str(x.get("document_type", "其他") or "其他")
    x["date"] = str(x.get("date", "") or "").strip()
    x["reference_no"] = str(x.get("reference_no", "") or "").upper().strip()
    x["summary"] = str(x.get("summary", "") or "").strip()
    x["payment_method"] = str(x.get("payment_method", "未知") or "未知")
    x["location"] = str(x.get("location", "") or "").strip()
    x["plate_no"] = str(x.get("plate_no", "") or "").upper().strip()
    x["barcode_raw"] = str(x.get("barcode_raw", "") or "").strip()

    try:
        x["amount"] = int(float(x.get("amount", 0) or 0))
    except Exception:
        x["amount"] = 0

    return x


def write_excel(items, dept, applicant_name):
    if not os.path.exists(TEMPLATE_PATH):
        raise FileNotFoundError("找不到 template.xlsx")

    wb = openpyxl.load_workbook(TEMPLATE_PATH)
    ws = wb.active

    start_row, end_row, total_row = 5, 12, 13

    if len(items) > MAX_ITEMS:
        raise ValueError(f"目前範本最多 {MAX_ITEMS} 筆")

    ws["B3"] = dept or ""
    ws["F3"] = datetime.now().strftime("%Y-%m-%d")

    for r in range(start_row, end_row + 1):
        for c in range(1, 8):
            ws.cell(r, c).value = None

    for idx, raw in enumerate(items):
        item = normalize_item(raw)
        r = start_row + idx

        summary = item["summary"]
        if item["document_type"] == "停車繳費單" and not summary:
            summary = "停車費"
        elif item["document_type"] == "代收繳費單" and not summary:
            summary = "代收繳費"

        if item["location"] and item["document_type"] == "停車繳費單":
            summary = f"{summary or '停車費'}－{item['location']}"

        notes = []
        if item["plate_no"]:
            notes.append(f"車牌:{item['plate_no']}")
        if item["location"] and item["document_type"] != "停車繳費單":
            notes.append(item["location"])

        ws.cell(r, 1, item["date"])
        ws.cell(r, 2, item["reference_no"])
        ws.cell(r, 3, summary)
        ws.cell(r, 4, item["amount"]).number_format = '$#,##0'
        ws.cell(r, 5, item["payment_method"])
        ws.cell(r, 6, item["document_type"])
        ws.cell(r, 7, " / ".join(notes))

    ws.cell(total_row, 4, f"=SUM(D{start_row}:D{end_row})").number_format = '$#,##0'

    if applicant_name:
        ws["G16"] = applicant_name

    ws.print_area = "A1:G16"
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_margins.left = 0.25
    ws.page_margins.right = 0.25
    ws.page_margins.top = 0.35
    ws.page_margins.bottom = 0.35
    ws.print_options.horizontalCentered = True

    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()


# ---------- 側欄 ----------
with st.sidebar:
    st.header("⚙️ 請款資料")
    applicant = st.text_input("請款人姓名")
    department = st.text_input("請款單位")
    st.caption("Excel 範本最多 8 筆明細。")


# ---------- 主功能 ----------
tab1, tab2, tab3 = st.tabs([
    "📷 自動掃描",
    "✍️ 人工新增",
    "📋 請款明細",
])

with tab1:
    st.subheader("拍照 / 上傳單據")

    source_type = st.radio(
        "單據類型",
        ["自動判斷", "停車繳費單", "代收繳費單", "收據", "其他"],
        horizontal=True,
    )

    c1, c2 = st.columns(2)

    with c1:
        camera_file = st.camera_input("📸 手機直接拍照")

    with c2:
        upload_file = st.file_uploader(
            "🖼️ 或從相簿 / 電腦上傳",
            type=["jpg", "jpeg", "png", "webp"],
            key="v6_upload",
        )

    current_file = camera_file if camera_file is not None else upload_file

    if current_file is not None:
        st.image(current_file, caption="目前單據", width="stretch")

        if st.button("🔎 掃描 QR / 條碼", type="primary", width="stretch"):
            try:
                item, decoded = scan_document(current_file.getvalue(), source_type)

                if not decoded:
                    st.warning(
                        "這張照片沒有讀到 QR Code 或一維條碼。"
                        "V6 不使用大型 OCR，因此請改到「人工新增」輸入紙本上的文字。"
                    )
                elif item:
                    st.session_state.items_v6.append(item)
                    st.success(f"已讀到 {len(decoded)} 個 QR/條碼，並加入一筆資料。")
                    with st.expander("查看讀到的條碼內容"):
                        for i, d in enumerate(decoded, 1):
                            st.code(f"{i}. {d['format']}\n{d['text']}")
                    st.rerun()
            except Exception as exc:
                st.error(f"掃描失敗：{exc}")

    st.info(
        "V6 會掃描 QR Code 與一維條碼。"
        "若條碼沒有直接包含日期或金額，系統不會亂猜，請在明細表手動補上。"
    )


with tab2:
    st.subheader("人工新增")

    with st.form("manual_v6", clear_on_submit=True):
        doc_type = st.selectbox(
            "單據類型",
            ["停車繳費單", "代收繳費單", "收據", "發票", "其他"],
        )
        date = st.text_input("日期", placeholder="2026-09-09")
        reference_no = st.text_input("發票 / 單據 / 繳費編號")
        summary = st.text_input("付款內容摘要")
        amount = st.number_input("金額", min_value=0, step=1)
        payment_method = st.selectbox("付款方式", ["未知", "現金", "刷卡"])
        location = st.text_input("店家 / 停車場")
        plate_no = st.text_input("車牌號碼")

        submitted = st.form_submit_button("➕ 加入請款明細", type="primary", width="stretch")

        if submitted:
            if doc_type == "停車繳費單" and not summary:
                summary = "停車費"
            if doc_type == "代收繳費單" and not summary:
                summary = "代收繳費"

            st.session_state.items_v6.append({
                "document_type": doc_type,
                "date": date,
                "reference_no": reference_no,
                "summary": summary,
                "amount": int(amount),
                "payment_method": payment_method,
                "location": location,
                "plate_no": plate_no,
                "barcode_raw": "",
                "source": "人工輸入",
            })
            st.rerun()


with tab3:
    st.subheader("請款明細確認")

    if not st.session_state.items_v6:
        st.info("目前還沒有資料。")
    else:
        df = pd.DataFrame(st.session_state.items_v6)

        expected = [
            "document_type", "date", "reference_no", "summary", "amount",
            "payment_method", "location", "plate_no", "source", "barcode_raw",
        ]
        for col in expected:
            if col not in df.columns:
                df[col] = ""

        df = df[expected]

        edited = st.data_editor(
            df,
            hide_index=True,
            width="stretch",
            num_rows="dynamic",
            column_config={
                "document_type": st.column_config.SelectboxColumn(
                    "單據類型",
                    options=["發票", "停車繳費單", "代收繳費單", "收據", "其他"],
                ),
                "date": st.column_config.TextColumn("日期"),
                "reference_no": st.column_config.TextColumn("發票 / 單據編號"),
                "summary": st.column_config.TextColumn("付款摘要"),
                "amount": st.column_config.NumberColumn("金額", min_value=0, step=1, format="%d"),
                "payment_method": st.column_config.SelectboxColumn(
                    "付款方式", options=["未知", "現金", "刷卡"]
                ),
                "location": st.column_config.TextColumn("店家 / 停車場"),
                "plate_no": st.column_config.TextColumn("車牌"),
                "source": st.column_config.TextColumn("來源", disabled=True),
                "barcode_raw": st.column_config.TextColumn("原始條碼", disabled=True),
            },
            key="items_editor_v6",
        )

        st.session_state.items_v6 = edited.to_dict("records")

        total = sum(
            int(float(x.get("amount", 0) or 0))
            for x in st.session_state.items_v6
        )

        c1, c2 = st.columns(2)
        c1.metric("目前筆數", len(st.session_state.items_v6))
        c2.metric("請款合計", f"${total:,}")

        if len(st.session_state.items_v6) > MAX_ITEMS:
            st.error(f"目前 Excel 範本最多 {MAX_ITEMS} 筆，請刪除部分資料或分成兩張請款單。")
        else:
            try:
                excel = write_excel(
                    st.session_state.items_v6,
                    department,
                    applicant,
                )
                st.download_button(
                    "📥 下載請款單 Excel",
                    data=excel,
                    file_name=f"請款單_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary",
                    width="stretch",
                )
            except Exception as exc:
                st.error(f"Excel 產生失敗：{exc}")

        if st.button("🗑️ 清空本次請款資料", width="stretch"):
            st.session_state.items_v6 = []
            st.rerun()


with st.expander("ℹ️ V6 能做什麼"):
    st.markdown("""
- **電子發票**：QR Code 自動辨識發票號碼、日期、金額。
- **停車繳費單 / 代收繳費單**：掃描 QR Code 與一維條碼，將條碼內容自動帶入。
- **日期 / 金額**：只有條碼內容本身能明確判斷時才自動填，避免誤判。
- **紙本中文文字**：V6 不使用大型 OCR，因此沒有條碼的內容請人工補登。
- **手機使用**：支援直接拍照，也支援相簿上傳。
- **Excel**：固定使用 template.xlsx，最多 8 筆，A4 橫向單頁列印。
""")
