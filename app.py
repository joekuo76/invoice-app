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
st.title("🧾 請款單自動生成工具 V6.2")
st.caption("保留 V5 操作方式｜可批量上傳｜新增 QR / 一維條碼辨識")
st.success("輕量穩定版：不使用 Gemini、不使用 EasyOCR、不使用 Torch。")

st.markdown("""
<style>
/* 手機版整體留白與按鈕 */
@media (max-width: 768px) {
    .block-container {
        padding-top: 1rem;
        padding-left: 0.8rem;
        padding-right: 0.8rem;
        padding-bottom: 3rem;
    }
    h1 { font-size: 1.65rem !important; }
    h2, h3 { font-size: 1.25rem !important; }
    div[data-testid="stButton"] button,
    div[data-testid="stDownloadButton"] button {
        min-height: 3rem;
        width: 100%;
    }
}
/* 卡片 */
.expense-card {
    border: 1px solid rgba(128,128,128,.28);
    border-radius: 14px;
    padding: 14px 16px;
    margin: 10px 0;
}
.expense-card .amount {
    font-size: 1.25rem;
    font-weight: 700;
}
</style>
""", unsafe_allow_html=True)

TEMPLATE_PATH = "template.xlsx"
MAX_ITEMS = 8

if "items_v62" not in st.session_state:
    st.session_state.items_v62 = []


# ---------- 共用 ----------
def load_image(data: bytes):
    img = Image.open(io.BytesIO(data))
    return ImageOps.exif_transpose(img).convert("RGB")


def decode_barcodes(img):
    results = []
    try:
        decoded = zxingcpp.read_barcodes(np.array(img))
        for r in decoded:
            text = (getattr(r, "text", "") or "").strip()
            if text:
                results.append({
                    "text": text,
                    "format": str(getattr(r, "format", "Unknown")),
                })
    except Exception:
        pass
    return results


def roc_date_to_iso(v: str):
    digits = re.sub(r"\D", "", v or "")
    if len(digits) != 7:
        return ""
    try:
        return datetime(
            int(digits[:3]) + 1911,
            int(digits[3:5]),
            int(digits[5:7]),
        ).strftime("%Y-%m-%d")
    except Exception:
        return ""


def parse_invoice_qr(text):
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
        "source": "QR自動",
        "barcode_raw": s,
    }


def guess_date_from_barcode(text: str):
    t = re.sub(r"\s+", "", text or "")

    for m in re.finditer(r"(?<!\d)(20\d{2})(0[1-9]|1[0-2])([0-2]\d|3[01])(?!\d)", t):
        try:
            return datetime(int(m.group(1)), int(m.group(2)), int(m.group(3))).strftime("%Y-%m-%d")
        except Exception:
            pass

    for m in re.finditer(r"(?<!\d)(1\d{2})(0[1-9]|1[0-2])([0-2]\d|3[01])(?!\d)", t):
        d = roc_date_to_iso("".join(m.groups()))
        if d:
            return d

    return ""


def guess_amount_from_barcode(text: str):
    for p in [
        r"(?i)(?:amount|amt)\s*[:=]\s*([0-9]{1,7})",
        r"金額\s*[:：=]?\s*([0-9]{1,7})",
    ]:
        m = re.search(p, text or "")
        if m:
            try:
                return int(m.group(1))
            except Exception:
                pass
    return 0


def parse_general_barcode(decoded, preferred_type):
    if not decoded:
        return None

    # 優先使用非 QR 的一維條碼；否則使用最長內容
    ordered = sorted(
        decoded,
        key=lambda x: (("QRCode" not in x["format"]), len(x["text"])),
        reverse=True,
    )
    primary = ordered[0]
    pool = " ".join(x["text"] for x in decoded)

    doc_type = preferred_type
    if doc_type == "自動判斷":
        doc_type = "代收繳費單"

    summary = ""
    if doc_type == "停車繳費單":
        summary = "停車費"
    elif doc_type == "代收繳費單":
        summary = "代收繳費"

    ref = primary["text"].strip()
    if len(ref) > 40:
        ref = ref[:40]

    return {
        "document_type": doc_type,
        "date": guess_date_from_barcode(pool),
        "reference_no": ref,
        "summary": summary,
        "amount": guess_amount_from_barcode(pool),
        "payment_method": "未知",
        "location": "",
        "plate_no": "",
        "source": "條碼自動",
        "barcode_raw": " | ".join(f'{x["format"]}:{x["text"]}' for x in decoded),
    }


def scan_document(data, preferred_type="自動判斷"):
    img = load_image(data)
    decoded = decode_barcodes(img)

    for d in decoded:
        inv = parse_invoice_qr(d["text"])
        if inv:
            return inv, decoded

    if decoded:
        return parse_general_barcode(decoded, preferred_type), decoded

    return None, []


def normalize_item(x):
    x = dict(x)
    for k in [
        "document_type", "date", "reference_no", "summary",
        "payment_method", "location", "plate_no", "source", "barcode_raw"
    ]:
        x[k] = str(x.get(k, "") or "").strip()

    if not x["document_type"]:
        x["document_type"] = "其他"
    if not x["payment_method"]:
        x["payment_method"] = "未知"

    try:
        x["amount"] = int(float(x.get("amount", 0) or 0))
    except Exception:
        x["amount"] = 0

    return x


def write_excel(items, department, applicant):
    if not os.path.exists(TEMPLATE_PATH):
        raise FileNotFoundError("找不到 template.xlsx")

    wb = openpyxl.load_workbook(TEMPLATE_PATH)
    ws = wb.active

    start_row, end_row, total_row = 5, 12, 13

    if len(items) > MAX_ITEMS:
        raise ValueError(f"目前範本最多 {MAX_ITEMS} 筆")

    ws["B3"] = department or ""
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

    if applicant:
        ws["G16"] = applicant

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
    st.caption("目前 Excel 範本最多 8 筆明細。")


# ---------- V5 風格介面 ----------
tab1, tab2, tab3 = st.tabs([
    "📷 批量上傳辨識",
    "✍️ 人工新增",
    "📋 請款明細",
])

with tab1:
    st.subheader("批量上傳單據")
    st.write("不會自動開啟相機。手機可從檔案選擇器自行決定拍照或選相簿。")

    preferred_type = st.radio(
        "這批照片主要是哪一類？",
        ["自動判斷", "停車繳費單", "代收繳費單", "收據", "其他"],
        horizontal=True,
    )

    files = st.file_uploader(
        "選擇照片（可一次多選）",
        type=["jpg", "jpeg", "png", "webp"],
        accept_multiple_files=True,
        key="batch_upload_v62",
    )

    if files:
        st.caption(f"已選擇 {len(files)} 張照片")

        cols = st.columns(min(4, len(files)))
        for i, f in enumerate(files):
            with cols[i % len(cols)]:
                st.image(f, caption=f.name, width="stretch")

        if st.button("🔎 批量掃描 QR / 條碼", type="primary", width="stretch"):
            success = 0
            failed = 0

            for f in files:
                try:
                    item, decoded = scan_document(f.getvalue(), preferred_type)

                    if item:
                        item["filename"] = f.name
                        st.session_state.items_v62.append(item)
                        success += 1
                    else:
                        st.session_state.items_v62.append({
                            "document_type": preferred_type if preferred_type != "自動判斷" else "其他",
                            "date": "",
                            "reference_no": "",
                            "summary": "",
                            "amount": 0,
                            "payment_method": "未知",
                            "location": "",
                            "plate_no": "",
                            "source": "未辨識",
                            "barcode_raw": "",
                            "filename": f.name,
                        })
                        failed += 1
                except Exception:
                    failed += 1

            if success:
                st.success(f"成功辨識 {success} 張。")
            if failed:
                st.warning(f"{failed} 張沒有讀到可用 QR / 條碼，已建立空白資料供你手動補填。")

            st.rerun()

    st.info(
        "V6.2 不做大型 OCR。電子發票會讀 QR；代收／停車單會讀條碼。"
        "沒有讀到的照片仍會建立一筆空白明細，方便你直接補資料。"
    )


with tab2:
    st.subheader("人工新增")

    with st.form("manual_form_v62", clear_on_submit=True):
        doc_type = st.selectbox(
            "單據類型",
            ["停車繳費單", "代收繳費單", "收據", "發票", "其他"],
        )
        date = st.text_input("日期", placeholder="2026-09-09")
        reference_no = st.text_input("發票 / 單據 / 停車編號")
        summary = st.text_input("付款內容摘要")
        amount = st.number_input("金額", min_value=0, step=1)
        payment_method = st.selectbox("付款方式", ["未知", "現金", "刷卡"])
        location = st.text_input("店家 / 停車場")
        plate_no = st.text_input("車牌號碼")

        submitted = st.form_submit_button("➕ 加入請款明細", type="primary", width="stretch")

        if submitted:
            if doc_type == "停車繳費單" and not summary:
                summary = "停車費"
            elif doc_type == "代收繳費單" and not summary:
                summary = "代收繳費"

            st.session_state.items_v62.append({
                "document_type": doc_type,
                "date": date,
                "reference_no": reference_no,
                "summary": summary,
                "amount": int(amount),
                "payment_method": payment_method,
                "location": location,
                "plate_no": plate_no,
                "source": "人工輸入",
                "barcode_raw": "",
                "filename": "",
            })
            st.rerun()


with tab3:
    st.subheader("請款明細確認")

    if not st.session_state.items_v62:
        st.info("目前還沒有資料。")
    else:
        # 手機預設使用卡片模式；需要一次大量修改時可切到表格模式。
        view_mode = st.radio(
            "顯示方式",
            ["📱 卡片模式", "🖥️ 表格模式"],
            horizontal=True,
            help="手機建議卡片模式；電腦大量修改建議表格模式。",
        )

        if view_mode == "📱 卡片模式":
            delete_index = None

            for i, raw in enumerate(st.session_state.items_v62):
                item = normalize_item(raw)
                title = item["document_type"] or "其他"
                filename = str(raw.get("filename", "") or "")
                ref = item["reference_no"] or "尚未填寫"
                date_text = item["date"] or "尚未填寫"
                summary_text = item["summary"] or "尚未填寫"
                amount_text = f'${item["amount"]:,}'

                st.markdown(
                    f"""
                    <div class="expense-card">
                        <strong>第 {i+1} 筆｜{title}</strong><br>
                        <small>{filename}</small><br><br>
                        日期：{date_text}<br>
                        編號：{ref}<br>
                        摘要：{summary_text}<br>
                        <span class="amount">{amount_text}</span>
                    </div>
                    """,
                    unsafe_allow_html=True,
                )

                with st.expander(f"✏️ 編輯第 {i+1} 筆"):
                    doc_type = st.selectbox(
                        "單據類型",
                        ["發票", "停車繳費單", "代收繳費單", "收據", "其他"],
                        index=["發票", "停車繳費單", "代收繳費單", "收據", "其他"].index(item["document_type"])
                        if item["document_type"] in ["發票", "停車繳費單", "代收繳費單", "收據", "其他"] else 4,
                        key=f"card_type_{i}",
                    )
                    date = st.text_input("日期", value=item["date"], key=f"card_date_{i}")
                    ref_no = st.text_input("發票 / 單據編號", value=item["reference_no"], key=f"card_ref_{i}")
                    summary = st.text_input("付款摘要", value=item["summary"], key=f"card_summary_{i}")
                    amount = st.number_input("金額", min_value=0, step=1, value=item["amount"], key=f"card_amount_{i}")
                    pay = st.selectbox(
                        "付款方式",
                        ["未知", "現金", "刷卡"],
                        index=["未知", "現金", "刷卡"].index(item["payment_method"])
                        if item["payment_method"] in ["未知", "現金", "刷卡"] else 0,
                        key=f"card_pay_{i}",
                    )
                    location = st.text_input("店家 / 停車場", value=item["location"], key=f"card_location_{i}")
                    plate = st.text_input("車牌", value=item["plate_no"], key=f"card_plate_{i}")

                    c_save, c_delete = st.columns(2)
                    if c_save.button("💾 儲存", key=f"save_card_{i}", width="stretch"):
                        updated = dict(raw)
                        updated.update({
                            "document_type": doc_type,
                            "date": date,
                            "reference_no": ref_no,
                            "summary": summary,
                            "amount": int(amount),
                            "payment_method": pay,
                            "location": location,
                            "plate_no": plate,
                        })
                        st.session_state.items_v62[i] = updated
                        st.rerun()

                    if c_delete.button("🗑️ 刪除", key=f"delete_card_{i}", width="stretch"):
                        delete_index = i

            if delete_index is not None:
                st.session_state.items_v62.pop(delete_index)
                st.rerun()

        else:
            df = pd.DataFrame(st.session_state.items_v62)
            expected = [
                "filename", "document_type", "date", "reference_no",
                "summary", "amount", "payment_method", "location",
                "plate_no", "source", "barcode_raw",
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
                    "filename": st.column_config.TextColumn("照片", disabled=True),
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
                key="items_editor_v62",
            )
            st.session_state.items_v62 = edited.to_dict("records")

        total = sum(
            int(float(x.get("amount", 0) or 0))
            for x in st.session_state.items_v62
        )

        st.divider()
        c1, c2 = st.columns(2)
        c1.metric("目前筆數", len(st.session_state.items_v62))
        c2.metric("請款合計", f"${total:,}")

        if len(st.session_state.items_v62) > MAX_ITEMS:
            st.error(f"目前 Excel 範本最多 {MAX_ITEMS} 筆，請刪除部分資料或分成兩張請款單。")
        else:
            try:
                excel = write_excel(
                    st.session_state.items_v62,
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
            st.session_state.items_v62 = []
            st.rerun()


with st.expander("ℹ️ V6.2 手機版優化"):
    st.markdown("""
- **不會自動開啟相機**。
- **可一次批量上傳多張照片**。
- 手機建議使用 **卡片模式**：一筆單據一張卡片，不需要左右滑大型表格。
- 點「編輯」即可修改日期、編號、摘要、金額、付款方式、店家與車牌。
- 電腦仍可切換 **表格模式**，方便一次修改多筆資料。
- 電子發票支援 QR；停車單 / 代收繳費單支援 QR + 一維條碼。
- 不使用 Gemini / EasyOCR / Torch。
""")
