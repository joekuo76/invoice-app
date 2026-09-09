import hashlib
import io
import json
import os
import random
import re
import sqlite3
import time
from datetime import datetime
from typing import Literal

import openpyxl
import pandas as pd
import streamlit as st
from google import genai
from google.genai import types
from PIL import Image, ImageOps
from pydantic import BaseModel, Field, ValidationError


# =========================================================
# 基本設定
# =========================================================

st.set_page_config(
    page_title="請款單自動生成系統",
    page_icon="🧾",
    layout="wide",
)

st.title("🧾 請款單自動生成工具")
st.caption("上傳發票 / 收據 / 停車繳費單 → AI 辨識 → 人工確認 → 產出 Excel")


# =========================================================
# 可調整參數
# =========================================================

MODEL_NAME = os.environ.get("GEMINI_MODEL", "gemini-3.8-flash")
REQUEST_INTERVAL_SECONDS = 3.0
MAX_RETRIES = 2
MAX_IMAGE_SIDE = 2400

CACHE_DB = "receipt_cache.sqlite3"
TEMPLATE_PATH = "template.xlsx"


# =========================================================
# API Key
# =========================================================

def get_api_key() -> str:
    try:
        if "GEMINI_API_KEY" in st.secrets:
            return str(st.secrets["GEMINI_API_KEY"]).strip()
    except Exception:
        pass

    return os.environ.get("GEMINI_API_KEY", "").strip()


api_key = get_api_key()


# =========================================================
# Sidebar
# =========================================================

with st.sidebar:
    st.header("⚙️ 系統設定")

    applicant_name = st.text_input("請款人姓名", value="")
    department = st.text_input("請款單位", value="")

    st.divider()

    st.write("**AI 模型**")
    st.code(MODEL_NAME, language=None)

    if api_key:
        st.success("✅ Gemini API Key 已載入")
    else:
        manual_key = st.text_input(
            "Gemini API Key",
            type="password",
            help="建議正式部署時改用 Streamlit Secrets。",
        )
        if manual_key:
            api_key = manual_key.strip()

    st.divider()

    if st.button("🗑️ 清除本機辨識快取", use_container_width=True):
        try:
            if os.path.exists(CACHE_DB):
                os.remove(CACHE_DB)
            st.success("快取已清除")
        except Exception as exc:
            st.error(f"無法清除快取：{exc}")


# =========================================================
# Structured Output Schema
# =========================================================

class ReceiptData(BaseModel):
    document_type: Literal["發票", "收據", "停車繳費單", "其他"] = "其他"

    date: str = Field(
        default="",
        description="消費、交易、停車或開立日期，格式 YYYY-MM-DD；無法辨識則空字串。",
    )

    invoice_no: str = Field(
        default="",
        description="台灣統一發票號碼，兩碼大寫英文字母加八碼數字；沒有則空字串。",
    )

    parking_no: str = Field(
        default="",
        description="停車繳費單號、停車單號、交易編號或繳費編號；沒有則空字串。",
    )

    location: str = Field(
        default="",
        description="停車場名稱、路段、店家名稱或可辨識的地點資訊。",
    )

    plate_no: str = Field(
        default="",
        description="車牌號碼；沒有則空字串。",
    )

    summary: str = Field(
        default="",
        description="簡短中文付款內容摘要，例如加油、停車費、餐費、五金耗材。",
    )

    amount: int = Field(
        default=0,
        ge=0,
        description="實際支付或應繳總額，僅整數，不含逗號與貨幣符號。",
    )

    payment_method: Literal["現金", "刷卡", "未知"] = "未知"


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
                model_name TEXT NOT NULL,
                result_json TEXT NOT NULL,
                created_at INTEGER NOT NULL
            )
            """
        )
        conn.commit()
    finally:
        conn.close()


def get_image_hash(image_bytes: bytes) -> str:
    return hashlib.sha256(image_bytes).hexdigest()


def read_cache(image_hash: str) -> dict | None:
    init_cache_db()

    conn = sqlite3.connect(CACHE_DB)
    try:
        row = conn.execute(
            """
            SELECT result_json
            FROM receipt_cache
            WHERE image_hash = ? AND model_name = ?
            """,
            (image_hash, MODEL_NAME),
        ).fetchone()
    finally:
        conn.close()

    if not row:
        return None

    try:
        return json.loads(row[0])
    except Exception:
        return None


def write_cache(image_hash: str, result: dict) -> None:
    init_cache_db()

    conn = sqlite3.connect(CACHE_DB)
    try:
        conn.execute(
            """
            INSERT OR REPLACE INTO receipt_cache
            (image_hash, model_name, result_json, created_at)
            VALUES (?, ?, ?, ?)
            """,
            (
                image_hash,
                MODEL_NAME,
                json.dumps(result, ensure_ascii=False),
                int(time.time()),
            ),
        )
        conn.commit()
    finally:
        conn.close()


# =========================================================
# 圖片處理
# =========================================================

def prepare_image(image_bytes: bytes) -> Image.Image:
    img = Image.open(io.BytesIO(image_bytes))
    img = ImageOps.exif_transpose(img)

    if img.mode != "RGB":
        img = img.convert("RGB")

    if max(img.size) > MAX_IMAGE_SIDE:
        img.thumbnail(
            (MAX_IMAGE_SIDE, MAX_IMAGE_SIDE),
            Image.Resampling.LANCZOS,
        )

    return img


def basic_image_check(image_bytes: bytes) -> tuple[bool, str]:
    try:
        img = Image.open(io.BytesIO(image_bytes))
        width, height = img.size

        if width < 400 or height < 400:
            return False, f"圖片解析度偏低（{width}×{height}），建議重新拍攝。"

        return True, ""
    except Exception:
        return False, "圖片檔案無法讀取。"


# =========================================================
# 辨識後清理 / 驗證
# =========================================================

def normalize_invoice_no(value: str) -> str:
    value = (value or "").upper()
    value = re.sub(r"[^A-Z0-9]", "", value)

    match = re.search(r"[A-Z]{2}\d{8}", value)
    return match.group(0) if match else ""


def normalize_plate_no(value: str) -> str:
    value = (value or "").upper().strip()
    value = re.sub(r"\s+", "", value)
    return value


def normalize_result(data: dict) -> dict:
    result = ReceiptData.model_validate(data).model_dump()

    result["invoice_no"] = normalize_invoice_no(result.get("invoice_no", ""))
    result["plate_no"] = normalize_plate_no(result.get("plate_no", ""))

    result["summary"] = str(result.get("summary", "")).strip()[:30]
    result["location"] = str(result.get("location", "")).strip()[:50]
    result["parking_no"] = str(result.get("parking_no", "")).strip()[:50]

    try:
        result["amount"] = int(result.get("amount", 0))
    except (TypeError, ValueError):
        result["amount"] = 0

    if result["document_type"] == "停車繳費單" and not result["summary"]:
        result["summary"] = "停車費"

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

    return warnings


# =========================================================
# Gemini
# =========================================================

OCR_PROMPT = """
你是專門處理台灣公司請款資料的發票、收據與停車繳費單辨識系統。

請先判斷單據類型：
- 發票
- 收據
- 停車繳費單
- 其他

然後抽取主要請款資訊。

【發票 / 收據】
- date：優先使用消費日期、交易日期或開立日期
- invoice_no：只填台灣統一發票號碼，格式為 2 個英文字母 + 8 個數字
- summary：例如加油、餐費、五金耗材、文具、交通費
- amount：找總計、合計、應付、實付、TOTAL
- payment_method：看到 VISA、Mastercard、JCB、信用卡、刷卡、卡號末四碼等資訊時輸出「刷卡」；
  明確看到現金則輸出「現金」，無法判斷輸出「未知」

【停車繳費單】
- document_type 必須輸出「停車繳費單」
- summary 優先輸出「停車費」
- date：優先辨識停車日期、繳費日期、交易日期
- parking_no：辨識繳費單號、停車單號、交易編號或繳費編號
- location：辨識停車場名稱、停車場站點、路段或其他地點
- plate_no：辨識車牌號碼
- invoice_no：只有真的存在正式統一發票號碼才填，否則留空
- amount：辨識應繳或實繳停車費
- payment_method：若無法判斷則輸出「未知」

【日期格式】
- 統一輸出 YYYY-MM-DD
- 民國年轉為西元年，例如 115/08/20 = 2026-08-20
- 不確定時留空，不要猜測

重要：
- 不要把統一編號 8 位數字誤認成發票號碼
- 不要把稅額、小計、折扣、找零、授權碼當成總額
- 不要猜測模糊或被遮住的文字
- 只回傳指定 schema
"""


def is_transient_error(error: Exception) -> bool:
    text = str(error).upper()
    markers = [
        "429",
        "503",
        "502",
        "500",
        "408",
        "RESOURCE_EXHAUSTED",
        "UNAVAILABLE",
        "DEADLINE_EXCEEDED",
        "TIMEOUT",
    ]
    return any(marker in text for marker in markers)


def parse_receipt(
    image_bytes: bytes,
    client: genai.Client,
) -> tuple[dict, bool]:

    image_hash = get_image_hash(image_bytes)

    cached = read_cache(image_hash)
    if cached is not None:
        return normalize_result(cached), True

    image = prepare_image(image_bytes)
    last_error = None

    for attempt in range(MAX_RETRIES + 1):
        try:
            response = client.models.generate_content(
                model=MODEL_NAME,
                contents=[OCR_PROMPT, image],
                config=types.GenerateContentConfig(
                    temperature=0,
                    response_mime_type="application/json",
                    response_schema=ReceiptData,
                ),
            )

            if not response.text:
                raise RuntimeError("Gemini 沒有回傳可解析內容。")

            parsed = ReceiptData.model_validate_json(response.text)
            result = normalize_result(parsed.model_dump())

            write_cache(image_hash, result)
            return result, False

        except (ValidationError, json.JSONDecodeError) as exc:
            last_error = exc

        except Exception as exc:
            last_error = exc

            if not is_transient_error(exc):
                raise

        if attempt < MAX_RETRIES:
            delay = (2 ** attempt) * 2 + random.uniform(0.2, 1.0)
            time.sleep(delay)

    raise RuntimeError(f"辨識失敗：{last_error}")


# =========================================================
# Excel
# =========================================================

def build_excel_summary(item: dict) -> str:
    summary = item.get("summary", "") or ""

    if item.get("document_type") == "停車繳費單":
        location = item.get("location", "")
        if location:
            return f"停車費－{location}"[:40]
        return "停車費"

    return summary


def build_excel_reference(item: dict) -> str:
    invoice_no = item.get("invoice_no", "")
    if invoice_no:
        return invoice_no

    if item.get("document_type") == "停車繳費單":
        parking_no = item.get("parking_no", "")
        if parking_no:
            return parking_no

    return ""


def write_to_excel_bytes(
    items: list[dict],
    template_path: str,
    dept: str,
    applicant: str,
) -> bytes:

    if not os.path.exists(template_path):
        raise FileNotFoundError(
            f"找不到 {template_path}。請將 Excel 範本命名為 template.xlsx "
            "並與 app.py 放在同一層。"
        )

    wb = openpyxl.load_workbook(template_path)
    ws = wb.active

    if dept:
        ws["B3"] = dept

    ws["F3"] = datetime.now().strftime("%Y-%m-%d")

    start_row = 5
    end_row = 11
    total_row = 12
    applicant_cell = "G15"

    max_items = end_row - start_row + 1

    if len(items) > max_items:
        raise ValueError(
            f"目前範本最多只能放 {max_items} 筆資料，"
            f"目前共有 {len(items)} 筆。"
        )

    for row in range(start_row, end_row + 1):
        for col in range(1, 8):
            ws.cell(row=row, column=col).value = None

    for idx, raw_item in enumerate(items):
        item = normalize_result(raw_item)
        row = start_row + idx

        ws.cell(row=row, column=1, value=item.get("date", ""))
        ws.cell(row=row, column=2, value=build_excel_reference(item))
        ws.cell(row=row, column=3, value=build_excel_summary(item))

        amount_cell = ws.cell(
            row=row,
            column=4,
            value=item.get("amount", 0),
        )
        amount_cell.number_format = '$#,##0'

        ws.cell(
            row=row,
            column=5,
            value=item.get("payment_method", "未知"),
        )

        if item.get("document_type") == "停車繳費單":
            ws.cell(row=row, column=6, value="停車繳費單")
        else:
            ws.cell(row=row, column=6, value=item.get("document_type", ""))

        note_parts = []

        if item.get("plate_no"):
            note_parts.append(f"車牌:{item['plate_no']}")

        if item.get("document_type") == "停車繳費單" and item.get("location"):
            note_parts.append(item["location"])

        ws.cell(
            row=row,
            column=7,
            value=" / ".join(note_parts),
        )

    total_cell = ws.cell(
        row=total_row,
        column=4,
        value=f"=SUM(D{start_row}:D{end_row})",
    )
    total_cell.number_format = '$#,##0'

    if applicant:
        ws[applicant_cell] = applicant

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)

    return output.getvalue()


# =========================================================
# Session State
# =========================================================

if "parsed_results" not in st.session_state:
    st.session_state["parsed_results"] = []

if "api_calls_this_run" not in st.session_state:
    st.session_state["api_calls_this_run"] = 0

if "cache_hits_this_run" not in st.session_state:
    st.session_state["cache_hits_this_run"] = 0


# =========================================================
# UI：上傳
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

    st.caption(
        f"共 {len(uploaded_files)} 張。相同圖片再次辨識時會優先讀取快取，不重複消耗 API。"
    )


# =========================================================
# UI：辨識
# =========================================================

if uploaded_files:
    if not api_key:
        st.warning("⚠️ 請先設定 Gemini API Key。")
    else:
        if st.button(
            "🚀 開始辨識單據",
            type="primary",
            use_container_width=True,
        ):
            client = genai.Client(api_key=api_key)

            parsed_results = []
            st.session_state["api_calls_this_run"] = 0
            st.session_state["cache_hits_this_run"] = 0

            progress_bar = st.progress(0)
            status_text = st.empty()

            last_real_api_call_at = 0.0

            for i, uploaded_file in enumerate(uploaded_files):
                filename = uploaded_file.name
                image_bytes = uploaded_file.getvalue()

                status_text.write(
                    f"正在處理第 {i + 1}/{len(uploaded_files)} 張：{filename}"
                )

                ok, check_message = basic_image_check(image_bytes)
                if not ok:
                    st.warning(f"⚠️ {filename}：{check_message}")

                try:
                    image_hash = get_image_hash(image_bytes)
                    cached = read_cache(image_hash)

                    if cached is not None:
                        result = normalize_result(cached)
                        from_cache = True
                        st.session_state["cache_hits_this_run"] += 1

                    else:
                        elapsed = time.time() - last_real_api_call_at

                        if elapsed < REQUEST_INTERVAL_SECONDS:
                            time.sleep(REQUEST_INTERVAL_SECONDS - elapsed)

                        result, from_cache = parse_receipt(
                            image_bytes,
                            client,
                        )

                        if not from_cache:
                            st.session_state["api_calls_this_run"] += 1
                            last_real_api_call_at = time.time()

                    result["source_file"] = filename
                    result["status"] = "快取" if from_cache else "AI 辨識"

                    warnings = receipt_has_warning(result)
                    result["warning"] = "、".join(warnings)

                    parsed_results.append(result)

                except Exception as exc:
                    st.error(f"❌ {filename} 辨識失敗")
                    st.exception(exc)

                    parsed_results.append(
                        {
                            "document_type": "其他",
                            "date": "",
                            "invoice_no": "",
                            "parking_no": "",
                            "location": "",
                            "plate_no": "",
                            "summary": "",
                            "amount": 0,
                            "payment_method": "未知",
                            "source_file": filename,
                            "status": "失敗",
                            "warning": "請人工輸入",
                        }
                    )

                progress_bar.progress((i + 1) / len(uploaded_files))

            status_text.empty()

            st.session_state["parsed_results"] = parsed_results

            st.success("🎉 辨識完成")

            col1, col2, col3 = st.columns(3)
            col1.metric(
                "本次真正 API 呼叫",
                st.session_state["api_calls_this_run"],
            )
            col2.metric(
                "快取命中",
                st.session_state["cache_hits_this_run"],
            )
            col3.metric(
                "總圖片數",
                len(uploaded_files),
            )


# =========================================================
# UI：人工確認
# =========================================================

if st.session_state["parsed_results"]:
    st.divider()
    st.subheader("📋 辨識結果確認與微調")

    df = pd.DataFrame(st.session_state["parsed_results"])

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
        "source_file",
        "status",
    ]

    for col in preferred_columns:
        if col not in df.columns:
            df[col] = ""

    df = df[preferred_columns]

    edited_df = st.data_editor(
        df,
        num_rows="dynamic",
        use_container_width=True,
        hide_index=True,
        column_config={
            "document_type": st.column_config.SelectboxColumn(
                "單據類型",
                options=["發票", "收據", "停車繳費單", "其他"],
            ),
            "date": st.column_config.TextColumn("日期"),
            "invoice_no": st.column_config.TextColumn("發票號碼"),
            "parking_no": st.column_config.TextColumn("停車單號"),
            "location": st.column_config.TextColumn("停車場 / 地點"),
            "plate_no": st.column_config.TextColumn("車牌"),
            "summary": st.column_config.TextColumn("摘要"),
            "amount": st.column_config.NumberColumn(
                "金額",
                min_value=0,
                step=1,
                format="%d",
            ),
            "payment_method": st.column_config.SelectboxColumn(
                "付款方式",
                options=["現金", "刷卡", "未知"],
            ),
            "warning": st.column_config.TextColumn(
                "檢查提醒",
                disabled=True,
            ),
            "source_file": st.column_config.TextColumn(
                "來源檔案",
                disabled=True,
            ),
            "status": st.column_config.TextColumn(
                "來源",
                disabled=True,
            ),
        },
        key="receipt_editor",
    )

    st.caption(
        "⚠️ AI 辨識結果請務必人工確認，尤其是金額、日期、發票號碼、停車單號與車牌。"
    )

    excel_records = edited_df[
        [
            "document_type",
            "date",
            "invoice_no",
            "parking_no",
            "location",
            "plate_no",
            "summary",
            "amount",
            "payment_method",
        ]
    ].to_dict("records")

    try:
        excel_bytes = write_to_excel_bytes(
            excel_records,
            TEMPLATE_PATH,
            department,
            applicant_name,
        )

        st.download_button(
            label="📥 下載完成的請款單 Excel",
            data=excel_bytes,
            file_name="請款單.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            use_container_width=True,
        )

    except FileNotFoundError as exc:
        st.warning(str(exc))

    except Exception as exc:
        st.error(f"Excel 產出失敗：{exc}")


# =========================================================
# 說明
# =========================================================

with st.expander("ℹ️ 使用與部署說明"):
    st.markdown(
        """
### 支援單據

- 台灣統一發票
- 一般收據
- 加油發票
- 停車場繳費單
- 路邊停車繳費單
- 停車收據

### 停車繳費單辨識欄位

系統會盡量辨識：

- 停車日期 / 繳費日期
- 停車單號 / 繳費編號
- 停車場名稱 / 路段
- 車牌號碼
- 停車金額
- 付款方式

若停車單沒有正式發票號碼，Excel 的「發票編號」欄會改放停車單號，方便報帳追蹤。

### Streamlit Secrets

```toml
GEMINI_API_KEY = "你的 API Key"
```

如需更換模型：

```toml
GEMINI_MODEL = "gemini-3.8-flash"
```

### 專案檔案

```text
app.py
requirements.txt
template.xlsx
```

### API 節省機制

- SHA-256 圖片快取
- 相同圖片不重複消耗 API
- 真正呼叫 API 時才限速
- 429 / 503 等暫時性錯誤才重試
- Structured Output 降低 JSON 格式錯誤
        """
    )
