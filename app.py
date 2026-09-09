import hashlib
import io
import json
import os
import re
import sqlite3
from datetime import datetime
from typing import Any

import cv2
import numpy as np
import openpyxl
import pandas as pd
import streamlit as st
from PIL import Image, ImageOps

try:
    from rapidocr_onnxruntime import RapidOCR
except Exception:
    RapidOCR = None


# =========================================================
# 基本設定
# =========================================================

st.set_page_config(
    page_title="請款單自動生成系統",
    page_icon="🧾",
    layout="wide",
)

st.title("🧾 請款單自動生成工具")
st.caption("上傳發票 / 收據 / 停車繳費單 → QR Code + 本機 OCR → 人工確認 → 產出 Excel")
st.info("🔒 本版本不使用 Gemini API，也不需要 API Key；辨識在 Streamlit 執行環境內完成。")


TEMPLATE_PATH = "template.xlsx"
CACHE_DB = "receipt_cache_local.sqlite3"
OCR_ENGINE_VERSION = "rapidocr-v1"
MAX_IMAGE_SIDE = 2600


# =========================================================
# Sidebar
# =========================================================

with st.sidebar:
    st.header("⚙️ 系統設定")
    applicant_name = st.text_input("請款人姓名", value="")
    department = st.text_input("請款單位", value="")

    st.divider()
    st.write("**辨識方式**")
    st.code("QR Code + RapidOCR", language=None)
    st.caption("沒有每日 API 額度限制；較模糊、皺摺或特殊版型單據仍建議人工確認。")

    if st.button("🗑️ 清除本機辨識快取", use_container_width=True):
        try:
            if os.path.exists(CACHE_DB):
                os.remove(CACHE_DB)
            st.success("快取已清除")
        except Exception as exc:
            st.error(f"無法清除快取：{exc}")


# =========================================================
# 快取
# =========================================================

def init_cache_db() -> None:
    conn = sqlite3.connect(CACHE_DB)
    try:
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS receipt_cache (
                image_hash TEXT PRIMARY KEY,
                engine_version TEXT NOT NULL,
                result_json TEXT NOT NULL,
                created_at INTEGER NOT NULL
            )
            """
        )
        conn.commit()
    finally:
        conn.close()


def image_hash(image_bytes: bytes) -> str:
    return hashlib.sha256(image_bytes).hexdigest()


def read_cache(key: str) -> dict | None:
    init_cache_db()
    conn = sqlite3.connect(CACHE_DB)
    try:
        row = conn.execute(
            """
            SELECT result_json
            FROM receipt_cache
            WHERE image_hash = ? AND engine_version = ?
            """,
            (key, OCR_ENGINE_VERSION),
        ).fetchone()
    finally:
        conn.close()

    if not row:
        return None

    try:
        return json.loads(row[0])
    except Exception:
        return None


def write_cache(key: str, result: dict) -> None:
    init_cache_db()
    conn = sqlite3.connect(CACHE_DB)
    try:
        conn.execute(
            """
            INSERT OR REPLACE INTO receipt_cache
            (image_hash, engine_version, result_json, created_at)
            VALUES (?, ?, ?, strftime('%s','now'))
            """,
            (key, OCR_ENGINE_VERSION, json.dumps(result, ensure_ascii=False)),
        )
        conn.commit()
    finally:
        conn.close()


# =========================================================
# 圖片工具
# =========================================================

def prepare_pil(image_bytes: bytes) -> Image.Image:
    img = Image.open(io.BytesIO(image_bytes))
    img = ImageOps.exif_transpose(img)

    if img.mode != "RGB":
        img = img.convert("RGB")

    if max(img.size) > MAX_IMAGE_SIDE:
        img.thumbnail((MAX_IMAGE_SIDE, MAX_IMAGE_SIDE), Image.Resampling.LANCZOS)

    return img


def pil_to_bgr(img: Image.Image) -> np.ndarray:
    rgb = np.array(img)
    return cv2.cvtColor(rgb, cv2.COLOR_RGB2BGR)


def basic_image_check(image_bytes: bytes) -> tuple[bool, str]:
    try:
        img = Image.open(io.BytesIO(image_bytes))
        width, height = img.size
        if width < 500 or height < 500:
            return False, f"圖片解析度偏低（{width}×{height}），辨識結果可能需要人工修正。"
        return True, ""
    except Exception:
        return False, "圖片檔案無法讀取。"


# =========================================================
# QR Code
# =========================================================

def decode_qr_codes(bgr: np.ndarray) -> list[str]:
    detector = cv2.QRCodeDetector()
    found: list[str] = []

    try:
        ok, decoded_info, _, _ = detector.detectAndDecodeMulti(bgr)
        if ok and decoded_info:
            found.extend([x.strip() for x in decoded_info if x and x.strip()])
    except Exception:
        pass

    if not found:
        try:
            text, _, _ = detector.detectAndDecode(bgr)
            if text and text.strip():
                found.append(text.strip())
        except Exception:
            pass

    # 去重但保留順序
    return list(dict.fromkeys(found))


def roc_date_to_iso(value: str) -> str:
    digits = re.sub(r"\D", "", value or "")
    if len(digits) != 7:
        return ""
    try:
        roc_year = int(digits[:3])
        month = int(digits[3:5])
        day = int(digits[5:7])
        dt = datetime(roc_year + 1911, month, day)
        return dt.strftime("%Y-%m-%d")
    except Exception:
        return ""


def parse_taiwan_invoice_qr(text: str) -> dict:
    """
    台灣電子發票 QR 左碼常見固定欄位：
    發票號碼 10 碼 + 民國日期 7 碼 + 隨機碼 4 碼 +
    未稅金額 8 碼 HEX + 總金額 8 碼 HEX + ...
    """
    compact = (text or "").strip().replace("\n", "")

    if len(compact) < 37:
        return {}

    invoice_no = compact[:10].upper()
    date_roc = compact[10:17]
    total_hex = compact[29:37]

    if not re.fullmatch(r"[A-Z]{2}\d{8}", invoice_no):
        return {}

    if not re.fullmatch(r"\d{7}", date_roc):
        return {}

    try:
        amount = int(total_hex, 16)
    except Exception:
        amount = 0

    date_iso = roc_date_to_iso(date_roc)

    return {
        "invoice_no": invoice_no,
        "date": date_iso,
        "amount": amount if amount >= 0 else 0,
    }


# =========================================================
# OCR
# =========================================================

@st.cache_resource(show_spinner=False)
def get_ocr_engine():
    if RapidOCR is None:
        raise RuntimeError(
            "RapidOCR 尚未安裝。請確認 requirements.txt 已包含 rapidocr-onnxruntime。"
        )
    return RapidOCR()


def preprocess_for_ocr(bgr: np.ndarray) -> np.ndarray:
    h, w = bgr.shape[:2]

    # 小圖放大，有助於熱感紙字體辨識
    if max(h, w) < 1800:
        scale = 1800 / max(h, w)
        bgr = cv2.resize(
            bgr,
            None,
            fx=scale,
            fy=scale,
            interpolation=cv2.INTER_CUBIC,
        )

    # 輕度增強對比，不做過度二值化，避免中文字筆畫消失
    gray = cv2.cvtColor(bgr, cv2.COLOR_BGR2GRAY)
    clahe = cv2.createCLAHE(clipLimit=1.6, tileGridSize=(8, 8))
    gray = clahe.apply(gray)
    return cv2.cvtColor(gray, cv2.COLOR_GRAY2BGR)


def run_ocr(bgr: np.ndarray) -> tuple[list[str], list[float]]:
    engine = get_ocr_engine()
    img = preprocess_for_ocr(bgr)
    result, _ = engine(img)

    lines: list[str] = []
    scores: list[float] = []

    if not result:
        return lines, scores

    for item in result:
        try:
            # RapidOCR 常見格式：[box, text, score]
            text = str(item[1]).strip()
            score = float(item[2])
        except Exception:
            continue

        if text:
            lines.append(text)
            scores.append(score)

    return lines, scores


# =========================================================
# 文字解析
# =========================================================

def normalize_invoice_no(value: str) -> str:
    compact = re.sub(r"[^A-Z0-9]", "", (value or "").upper())
    match = re.search(r"[A-Z]{2}\d{8}", compact)
    return match.group(0) if match else ""


def parse_invoice_no(text: str) -> str:
    upper = text.upper()
    patterns = [
        r"\b([A-Z]{2})[\s\-]?(\d{8})\b",
        r"([A-Z]{2})\s*[-—]\s*(\d{8})",
    ]
    for pattern in patterns:
        m = re.search(pattern, upper)
        if m:
            return normalize_invoice_no("".join(m.groups()))
    return ""


def valid_iso_date(year: int, month: int, day: int) -> str:
    try:
        return datetime(year, month, day).strftime("%Y-%m-%d")
    except Exception:
        return ""


def parse_date(text: str) -> str:
    # 西元日期
    for pattern in [
        r"\b(20\d{2})[年/\-.](\d{1,2})[月/\-.](\d{1,2})日?\b",
        r"\b(20\d{2})(\d{2})(\d{2})\b",
    ]:
        m = re.search(pattern, text)
        if m:
            date = valid_iso_date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
            if date:
                return date

    # 民國日期 115/09/04、1150904
    for pattern in [
        r"(?<!\d)(1\d{2})[年/\-.](\d{1,2})[月/\-.](\d{1,2})日?(?!\d)",
        r"(?<!\d)(1\d{2})(\d{2})(\d{2})(?!\d)",
    ]:
        m = re.search(pattern, text)
        if m:
            date = valid_iso_date(int(m.group(1)) + 1911, int(m.group(2)), int(m.group(3)))
            if date:
                return date

    return ""


def parse_money_token(value: str) -> int:
    cleaned = value.replace(",", "").replace("，", "")
    cleaned = re.sub(r"[^\d]", "", cleaned)
    if not cleaned:
        return 0
    try:
        number = int(cleaned)
    except Exception:
        return 0
    return number if 0 <= number <= 10_000_000 else 0


def parse_amount(lines: list[str]) -> int:
    # 優先看明確總額關鍵字附近
    priority_keywords = [
        "應繳金額", "繳費金額", "實付", "應付", "總計", "合計",
        "TOTAL", "總額", "金額", "小計",
    ]

    candidates: list[tuple[int, int]] = []

    for idx, line in enumerate(lines):
        normalized = line.upper().replace("＄", "$")
        priority = 0
        for rank, kw in enumerate(priority_keywords):
            if kw in normalized:
                priority = 100 - rank
                break

        if priority:
            nums = re.findall(r"(?:NT\$?|TWD|\$)?\s*([0-9][0-9,]{0,10})", normalized)
            for token in nums:
                value = parse_money_token(token)
                if value > 0:
                    candidates.append((priority, value))

            # 有些熱感紙「合計」和金額分兩行
            if idx + 1 < len(lines):
                nums2 = re.findall(r"([0-9][0-9,]{0,10})", lines[idx + 1])
                for token in nums2:
                    value = parse_money_token(token)
                    if value > 0:
                        candidates.append((priority - 5, value))

    if candidates:
        candidates.sort(key=lambda x: (x[0], x[1]), reverse=True)
        return candidates[0][1]

    # 後備：只採用帶 $ / NT / 元 的數值，降低把日期或編號當金額的機率
    fallback: list[int] = []
    for line in lines:
        for token in re.findall(r"(?:NT\$?|TWD|\$)\s*([0-9][0-9,]{0,10})", line.upper()):
            value = parse_money_token(token)
            if value > 0:
                fallback.append(value)

        for token in re.findall(r"([0-9][0-9,]{0,10})\s*元", line):
            value = parse_money_token(token)
            if value > 0:
                fallback.append(value)

    return max(fallback) if fallback else 0


def parse_plate_no(text: str) -> str:
    upper = text.upper().replace("—", "-").replace("－", "-")
    # 新式 / 舊式台灣車牌的保守抓法
    patterns = [
        r"(?:車牌|車號|PLATE)\s*[:：]?\s*([A-Z0-9]{2,4}-[A-Z0-9]{2,4})",
        r"\b([A-Z]{2,4}-\d{2,4})\b",
        r"\b(\d{2,4}-[A-Z]{2,4})\b",
    ]
    for pattern in patterns:
        m = re.search(pattern, upper)
        if m:
            return m.group(1)
    return ""


def parse_parking_no(lines: list[str]) -> str:
    keywords = ["停車單號", "繳費單號", "交易編號", "繳費編號", "單號", "序號"]
    for line in lines:
        for kw in keywords:
            if kw in line:
                tail = line.split(kw, 1)[-1]
                tail = re.sub(r"^[：:\s]+", "", tail)
                m = re.search(r"[A-Z0-9\-]{5,30}", tail.upper())
                if m:
                    return m.group(0)
    return ""


def parse_location(lines: list[str], document_type: str) -> str:
    joined = " ".join(lines)

    if document_type == "停車繳費單":
        for line in lines:
            if any(k in line for k in ["停車場", "停車站", "停車", "路段"]):
                cleaned = re.sub(r"\s+", "", line)
                if 2 <= len(cleaned) <= 40:
                    return cleaned[:40]

    known_merchants = [
        ("全家", "全家便利商店"),
        ("FAMILYMART", "全家便利商店"),
        ("7-ELEVEN", "7-ELEVEN"),
        ("統一超商", "7-ELEVEN"),
        ("台灣中油", "台灣中油"),
        ("中油", "台灣中油"),
        ("CPC", "台灣中油"),
        ("全聯", "全聯"),
        ("家樂福", "家樂福"),
    ]
    upper = joined.upper()
    for keyword, name in known_merchants:
        if keyword.upper() in upper:
            return name

    # 尋找「公司 / 商店 / 加油站」類名稱
    for line in lines[:12]:
        cleaned = re.sub(r"\s+", "", line)
        if any(k in cleaned for k in ["有限公司", "股份有限公司", "商店", "加油站", "便利商店"]):
            return cleaned[:40]

    return ""


def classify_document(lines: list[str]) -> str:
    text = " ".join(lines).upper()

    if any(k in text for k in ["停車", "PARKING", "車牌", "車號", "停車費"]):
        return "停車繳費單"

    if any(k in text for k in ["電子發票證明聯", "發票號碼", "統一發票"]):
        return "發票"

    if any(k in text for k in ["收據", "繳費明細", "交易明細"]):
        return "收據"

    return "其他"


def infer_summary(lines: list[str], document_type: str, location: str) -> str:
    text = " ".join(lines).upper()

    if document_type == "停車繳費單":
        return "停車費"
    if any(k in text for k in ["95PLUS", "92無鉛", "95無鉛", "98無鉛", "汽油", "柴油", "加油", "CPC"]):
        return "加油"
    if any(k in text for k in ["餐費", "餐廳", "便當", "FOOD", "咖啡"]):
        return "餐費"
    if any(k in text for k in ["文具", "影印", "紙張"]):
        return "文具用品"
    if any(k in text for k in ["五金", "材料", "耗材"]):
        return "工程耗材"
    if any(k in text for k in ["繳費明細", "代收", "繳費"]):
        return "代收繳費"
    if location:
        return location[:20]
    if document_type == "發票":
        return "發票消費"
    if document_type == "收據":
        return "收據消費"
    return ""


def infer_payment_method(lines: list[str]) -> str:
    text = " ".join(lines).upper()
    if any(k in text for k in ["VISA", "MASTER", "JCB", "信用卡", "刷卡", "CARD", "末四碼"]):
        return "刷卡"
    if "現金" in text:
        return "現金"
    return "未知"


def parse_local_receipt(image_bytes: bytes) -> dict:
    pil_img = prepare_pil(image_bytes)
    bgr = pil_to_bgr(pil_img)

    qr_texts = decode_qr_codes(bgr)
    qr_data: dict[str, Any] = {}

    for qr in qr_texts:
        parsed = parse_taiwan_invoice_qr(qr)
        if parsed:
            qr_data.update({k: v for k, v in parsed.items() if v not in ("", 0)})
            break

    lines, scores = run_ocr(bgr)
    full_text = "\n".join(lines)

    document_type = classify_document(lines)

    # 有合法電子發票 QR 時，即使 OCR 沒看到「發票」字樣，也視為發票
    if qr_data.get("invoice_no"):
        document_type = "發票"

    location = parse_location(lines, document_type)
    invoice_no = qr_data.get("invoice_no") or parse_invoice_no(full_text)
    date = qr_data.get("date") or parse_date(full_text)
    amount = int(qr_data.get("amount") or 0) or parse_amount(lines)
    plate_no = parse_plate_no(full_text) if document_type == "停車繳費單" else ""
    parking_no = parse_parking_no(lines) if document_type == "停車繳費單" else ""
    payment_method = infer_payment_method(lines)
    summary = infer_summary(lines, document_type, location)

    confidence = round(sum(scores) / len(scores), 3) if scores else 0.0

    result = {
        "document_type": document_type,
        "date": date,
        "invoice_no": normalize_invoice_no(invoice_no),
        "parking_no": parking_no,
        "location": location,
        "plate_no": plate_no,
        "summary": summary,
        "amount": amount,
        "payment_method": payment_method,
        "ocr_confidence": confidence,
        "qr_detected": "是" if qr_texts else "否",
    }

    warnings = receipt_has_warning(result)
    result["warning"] = "、".join(warnings)
    return result


def receipt_has_warning(data: dict) -> list[str]:
    warnings = []

    if not data.get("date"):
        warnings.append("日期未辨識")

    if int(data.get("amount", 0) or 0) <= 0:
        warnings.append("金額未辨識")

    invoice_no = data.get("invoice_no", "")
    if invoice_no and not re.fullmatch(r"[A-Z]{2}\d{8}", invoice_no):
        warnings.append("發票號碼格式異常")

    if data.get("document_type") == "停車繳費單":
        if not data.get("parking_no") and not data.get("plate_no"):
            warnings.append("停車單號/車牌未辨識")

    if not data.get("summary"):
        warnings.append("摘要未辨識")

    confidence = float(data.get("ocr_confidence", 0) or 0)
    if confidence and confidence < 0.72:
        warnings.append("OCR信心偏低")

    return warnings


# =========================================================
# Excel
# =========================================================

def normalize_result(item: dict) -> dict:
    result = dict(item)
    result["invoice_no"] = normalize_invoice_no(str(result.get("invoice_no", "")))
    result["plate_no"] = str(result.get("plate_no", "") or "").upper().strip()
    result["summary"] = str(result.get("summary", "") or "").strip()[:40]
    result["location"] = str(result.get("location", "") or "").strip()[:50]
    result["parking_no"] = str(result.get("parking_no", "") or "").strip()[:50]

    try:
        result["amount"] = int(float(result.get("amount", 0) or 0))
    except Exception:
        result["amount"] = 0

    if result.get("document_type") not in ["發票", "收據", "停車繳費單", "其他"]:
        result["document_type"] = "其他"

    if result.get("payment_method") not in ["現金", "刷卡", "未知"]:
        result["payment_method"] = "未知"

    if result["document_type"] == "停車繳費單" and not result["summary"]:
        result["summary"] = "停車費"

    return result


def build_excel_summary(item: dict) -> str:
    if item.get("document_type") == "停車繳費單":
        location = item.get("location", "")
        return f"停車費－{location}"[:40] if location else "停車費"
    return item.get("summary", "") or ""


def build_excel_reference(item: dict) -> str:
    if item.get("invoice_no"):
        return item["invoice_no"]
    if item.get("document_type") == "停車繳費單" and item.get("parking_no"):
        return item["parking_no"]
    return ""


def write_to_excel_bytes(
    items: list[dict],
    template_path: str,
    dept: str,
    applicant: str,
) -> bytes:
    if not os.path.exists(template_path):
        raise FileNotFoundError(
            f"找不到 {template_path}。請將範本命名為 template.xlsx 並與 app.py 放在同一層。"
        )

    wb = openpyxl.load_workbook(template_path)
    ws = wb.active

    if dept:
        ws["B3"] = dept
    ws["F3"] = datetime.now().strftime("%Y-%m-%d")

    start_row = 5
    end_row = 12
    total_row = 13
    applicant_cell = "G16"

    max_items = end_row - start_row + 1
    if len(items) > max_items:
        raise ValueError(f"目前範本最多只能放 {max_items} 筆資料，目前共有 {len(items)} 筆。")

    for row in range(start_row, end_row + 1):
        for col in range(1, 8):
            ws.cell(row=row, column=col).value = None

    for idx, raw_item in enumerate(items):
        item = normalize_result(raw_item)
        row = start_row + idx

        ws.cell(row=row, column=1, value=item.get("date", ""))
        ws.cell(row=row, column=2, value=build_excel_reference(item))
        ws.cell(row=row, column=3, value=build_excel_summary(item))

        amount_cell = ws.cell(row=row, column=4, value=item.get("amount", 0))
        amount_cell.number_format = '$#,##0'

        ws.cell(row=row, column=5, value=item.get("payment_method", "未知"))
        ws.cell(row=row, column=6, value=item.get("document_type", ""))

        note_parts = []
        if item.get("plate_no"):
            note_parts.append(f"車牌:{item['plate_no']}")
        if item.get("document_type") == "停車繳費單" and item.get("location"):
            note_parts.append(item["location"])

        ws.cell(row=row, column=7, value=" / ".join(note_parts))

    total_cell = ws.cell(row=total_row, column=4, value=f"=SUM(D{start_row}:D{end_row})")
    total_cell.number_format = '$#,##0'

    if applicant:
        ws[applicant_cell] = applicant

    # A4 整份請款單列印
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
    ws.page_margins.header = 0.1
    ws.page_margins.footer = 0.1
    ws.print_options.horizontalCentered = True

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output.getvalue()


# =========================================================
# Session State
# =========================================================

if "parsed_results_local" not in st.session_state:
    st.session_state["parsed_results_local"] = []

if "cache_hits_local" not in st.session_state:
    st.session_state["cache_hits_local"] = 0


# =========================================================
# 上傳
# =========================================================

uploaded_files = st.file_uploader(
    "上傳發票、收據或停車繳費單照片（可多選）",
    type=["jpg", "jpeg", "png", "webp"],
    accept_multiple_files=True,
)

if uploaded_files:
    st.subheader("🖼️ 待辨識圖片")
    preview_cols = st.columns(min(4, len(uploaded_files)))

    for idx, uploaded_file in enumerate(uploaded_files):
        with preview_cols[idx % len(preview_cols)]:
            st.image(
                uploaded_file,
                caption=uploaded_file.name,
                use_container_width=True,
            )

    st.caption(f"共 {len(uploaded_files)} 張。相同圖片會優先讀取本機快取。")


# =========================================================
# 辨識
# =========================================================

if uploaded_files:
    if st.button("🚀 開始本機辨識", type="primary", use_container_width=True):
        parsed_results = []
        st.session_state["cache_hits_local"] = 0

        progress_bar = st.progress(0)
        status_text = st.empty()

        for i, uploaded_file in enumerate(uploaded_files):
            filename = uploaded_file.name
            image_bytes = uploaded_file.getvalue()

            status_text.write(f"正在處理第 {i + 1}/{len(uploaded_files)} 張：{filename}")

            ok, message = basic_image_check(image_bytes)
            if not ok:
                st.warning(f"⚠️ {filename}：{message}")

            try:
                key = image_hash(image_bytes)
                cached = read_cache(key)

                if cached is not None:
                    result = normalize_result(cached)
                    result.update({
                        "ocr_confidence": cached.get("ocr_confidence", 0),
                        "qr_detected": cached.get("qr_detected", "否"),
                        "warning": cached.get("warning", ""),
                    })
                    status = "快取"
                    st.session_state["cache_hits_local"] += 1
                else:
                    result = parse_local_receipt(image_bytes)
                    write_cache(key, result)
                    status = "本機 OCR"

                result["source_file"] = filename
                result["status"] = status
                parsed_results.append(result)

            except Exception as exc:
                st.error(f"❌ {filename} 辨識失敗：{exc}")
                parsed_results.append({
                    "document_type": "其他",
                    "date": "",
                    "invoice_no": "",
                    "parking_no": "",
                    "location": "",
                    "plate_no": "",
                    "summary": "",
                    "amount": 0,
                    "payment_method": "未知",
                    "ocr_confidence": 0,
                    "qr_detected": "否",
                    "warning": "辨識失敗，請人工輸入",
                    "source_file": filename,
                    "status": "失敗",
                })

            progress_bar.progress((i + 1) / len(uploaded_files))

        status_text.empty()
        st.session_state["parsed_results_local"] = parsed_results

        success_count = sum(1 for x in parsed_results if x.get("status") != "失敗")
        warning_count = sum(1 for x in parsed_results if x.get("warning"))
        qr_count = sum(1 for x in parsed_results if x.get("qr_detected") == "是")

        st.success("🎉 本機辨識完成")
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("總圖片數", len(uploaded_files))
        c2.metric("完成", success_count)
        c3.metric("QR Code", qr_count)
        c4.metric("需人工確認", warning_count)


# =========================================================
# 人工確認
# =========================================================

if st.session_state["parsed_results_local"]:
    st.divider()
    st.subheader("📋 辨識結果確認與微調")
    st.caption("本機 OCR 對皺摺、反光、特殊字體較敏感；有警示的欄位請人工確認後再匯出。")

    df = pd.DataFrame(st.session_state["parsed_results_local"])

    preferred_columns = [
        "document_type",
        "date",
        "invoice_no",
        "parking_no",
        "location",
        "plate_no",
        "summary",
        "amount",
        "payment_method",
        "warning",
        "qr_detected",
        "ocr_confidence",
        "source_file",
        "status",
    ]

    for col in preferred_columns:
        if col not in df.columns:
            df[col] = ""

    df = df[preferred_columns]

    edited_df = st.data_editor(
        df,
        use_container_width=True,
        hide_index=True,
        num_rows="fixed",
        column_config={
            "document_type": st.column_config.SelectboxColumn(
                "單據類型",
                options=["發票", "收據", "停車繳費單", "其他"],
                required=True,
            ),
            "date": st.column_config.TextColumn("日期"),
            "invoice_no": st.column_config.TextColumn("發票號碼"),
            "parking_no": st.column_config.TextColumn("停車/繳費單號"),
            "location": st.column_config.TextColumn("店家 / 停車場"),
            "plate_no": st.column_config.TextColumn("車牌"),
            "summary": st.column_config.TextColumn("付款內容摘要"),
            "amount": st.column_config.NumberColumn("金額", min_value=0, step=1, format="%d"),
            "payment_method": st.column_config.SelectboxColumn(
                "付款方式",
                options=["現金", "刷卡", "未知"],
                required=True,
            ),
            "warning": st.column_config.TextColumn("提醒", disabled=True),
            "qr_detected": st.column_config.TextColumn("QR", disabled=True),
            "ocr_confidence": st.column_config.NumberColumn(
                "OCR信心",
                disabled=True,
                format="%.2f",
            ),
            "source_file": st.column_config.TextColumn("來源", disabled=True),
            "status": st.column_config.TextColumn("狀態", disabled=True),
        },
    )

    export_items = edited_df.to_dict("records")

    st.divider()

    if len(export_items) > 8:
        st.error("目前 Excel 範本最多 8 筆，請減少單次匯出數量。")
    else:
        try:
            excel_bytes = write_to_excel_bytes(
                export_items,
                TEMPLATE_PATH,
                department,
                applicant_name,
            )

            st.download_button(
                "📥 下載請款單 Excel",
                data=excel_bytes,
                file_name=f"請款單_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
                use_container_width=True,
            )
        except Exception as exc:
            st.error(f"Excel 產生失敗：{exc}")


# =========================================================
# 使用說明
# =========================================================

with st.expander("ℹ️ 無 Gemini 版辨識方式"):
    st.markdown(
        """
        **辨識流程**

        1. 先嘗試讀取台灣電子發票 QR Code。
        2. 再使用 RapidOCR 做中文 / 英數字 OCR。
        3. 程式依日期、發票號碼、總額、停車、車牌等規則自動整理欄位。
        4. 最後由使用者人工確認，再輸出 Excel。

        **拍照建議**

        - 單據盡量攤平、正面拍攝。
        - 不要裁掉 QR Code、日期、發票號碼或總額。
        - 避免手指遮住文字。
        - 熱感紙反光時，稍微改變拍攝角度。
        - 皺摺嚴重或字太淡的單據，請特別檢查辨識結果。

        **與 Gemini 版不同**

        - 不需要 API Key。
        - 沒有 Gemini 每日請求額度。
        - 不會因 429 API quota 停止。
        - 特殊單據的語意判斷能力較 AI 弱，所以保留人工修正流程。
        """
    )
