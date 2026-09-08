import hashlib
import io
import json
import os
import random
import re
import sqlite3
import time
from pathlib import Path
from typing import Literal

import openpyxl
import pandas as pd
import streamlit as st
from google import genai
from google.genai import types
from openpyxl.styles import Alignment, Border, Side
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
st.caption("上傳發票 / 收據照片 → AI 辨識 → 人工確認 → 產出 Excel")


# =========================================================
# 可調整參數
# =========================================================

# 優先使用目前的 Flash 模型。
# 若日後 Google 更換模型，可只改這裡，不必修改其他程式。
MODEL_NAME = os.environ.get("GEMINI_MODEL", "gemini-3.8-flash")

# 每次真正送 API 之間至少間隔幾秒。
# 公司內部使用建議 2~4 秒，不要瞬間大量併發。
REQUEST_INTERVAL_SECONDS = 3.0

# 自訂重試次數。SDK 本身已有暫時性錯誤重試，
# 這裡只再做應用層的保護，因此不要設太高。
MAX_RETRIES = 2

# 圖片最長邊。發票文字小，不建議壓到 1200px。
MAX_IMAGE_SIDE = 2400

CACHE_DB = "receipt_cache.sqlite3"
TEMPLATE_PATH = "template.xlsx"


# =========================================================
# API Key
# =========================================================

def get_api_key() -> str:
    """依序讀取 Streamlit Secrets、環境變數。"""
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
    date: str = Field(
        default="",
        description="消費或開立日期，格式 YYYY-MM-DD；無法辨識則空字串。",
    )
    invoice_no: str = Field(
        default="",
        description="台灣發票號碼，格式兩碼大寫英文字母加八碼數字；沒有則空字串。",
    )
    summary: str = Field(
        default="",
        description="簡短中文付款內容摘要，建議 2~12 個中文字。",
    )
    amount: int = Field(
        default=0,
        ge=0,
        description="實際支付總額，僅整數，不含逗號與貨幣符號。",
    )
    payment_method: Literal["現金", "刷卡", "未知"] = Field(
        default="未知",
        description="付款方式。",
    )


# =========================================================
# 快取
# =========================================================

def init_cache_db() -> None:
    """建立 SQLite 快取資料表。"""
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
    """
    1. 自動依 EXIF 修正旋轉
    2. 轉 RGB
    3. 保留較高解析度以提高發票小字辨識率
    """
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
    """很輕量的圖片檢查，避免明顯無效圖片浪費 API。"""
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


def normalize_result(data: dict) -> dict:
    """把模型回傳內容再做一次程式端防呆。"""
    result = ReceiptData.model_validate(data).model_dump()

    result["invoice_no"] = normalize_invoice_no(result.get("invoice_no", ""))

    summary = str(result.get("summary", "")).strip()
    result["summary"] = summary[:30]

    try:
        result["amount"] = int(result.get("amount", 0))
    except (TypeError, ValueError):
        result["amount"] = 0

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

    if not data.get("summary"):
        warnings.append("摘要未辨識")

    return warnings


# =========================================================
# Gemini
# =========================================================

OCR_PROMPT = """
你是專門處理台灣公司請款資料的發票與收據辨識系統。

請仔細閱讀圖片中所有可見文字，抽取一張單據的主要請款資訊。
若圖片是台灣電子發票證明聯、傳統發票、加油發票、停車收據、
餐飲收據、五金行收據、刷卡簽單或一般收據，都使用相同欄位回傳。

請遵守以下規則：

【date】
- 優先使用「消費日期」、「交易日期」或「發票開立日期」。
- 統一輸出 YYYY-MM-DD。
- 民國年轉為西元年，例如 115/08/20 = 2026-08-20。
- 不確定時輸出空字串，不要猜測。

【invoice_no】
- 台灣統一發票號碼通常為 2 個英文字母 + 8 個數字，例如 AB12345678。
- 英文字母輸出大寫。
- 不要把「統一編號 / 統編」8 位數字當成發票號碼。
- 一般收據沒有發票號碼時輸出空字串。

【summary】
- 根據店家名稱與品項產生簡短中文摘要。
- 例如：加油、停車費、餐費、工地飲料、五金耗材、文具、交通費。
- 不要照抄整張收據，也不要加入發票號碼。
- 如果無法判斷，可以使用店家名稱或「其他費用」。

【amount】
- 找出實際應付 / 實付總額。
- 優先辨識：總計、合計、應付、實付、TOTAL、總額。
- 不要把小計、稅額、折扣、找零、信用卡授權碼當成金額。
- 新台幣只輸出整數。
- 無法確認則輸出 0。

【payment_method】
- 若看到 VISA、Mastercard、JCB、信用卡、刷卡、卡號、末四碼等刷卡資訊，輸出「刷卡」。
- 明確看到現金付款則輸出「現金」。
- 無法判斷則輸出「未知」。

重要：
- 不要猜測模糊或被遮住的文字。
- 只回傳指定 schema 內容。
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
    """
    回傳 (辨識結果, 是否來自快取)
    """

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

            # response_schema 已限制格式，仍做 Pydantic 驗證。
            parsed = ReceiptData.model_validate_json(response.text)
            result = normalize_result(parsed.model_dump())

            write_cache(image_hash, result)
            return result, False

        except (ValidationError, json.JSONDecodeError) as exc:
            # 格式錯誤通常重試一次有機會恢復
            last_error = exc

        except Exception as exc:
            last_error = exc

            # 400 / 401 / 403 / 404 等通常不應重試
            if not is_transient_error(exc):
                raise

        if attempt < MAX_RETRIES:
            # 指數退避 + jitter
            delay = (2 ** attempt) * 2 + random.uniform(0.2, 1.0)
            time.sleep(delay)

    raise RuntimeError(f"辨識失敗：{last_error}")


# =========================================================
# Excel
# =========================================================

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

    start_row = 5

    thin_border = Border(
        left=Side(style="thin"),
        right=Side(style="thin"),
        top=Side(style="thin"),
        bottom=Side(style="thin"),
    )

    center_align = Alignment(
        horizontal="center",
        vertical="center",
        wrap_text=True,
    )

    right_align = Alignment(
        horizontal="right",
        vertical="center",
    )

    for idx, raw_item in enumerate(items):
        item = normalize_result(raw_item)
        row = start_row + idx

        ws.cell(row=row, column=1, value=item.get("date", "")).alignment = center_align
        ws.cell(row=row, column=2, value=item.get("invoice_no", "")).alignment = center_align
        ws.cell(row=row, column=3, value=item.get("summary", "")).alignment = center_align
        ws.cell(row=row, column=4, value=item.get("amount", 0)).alignment = right_align
        ws.cell(row=row, column=5, value=item.get("payment_method", "未知")).alignment = center_align

        for col in range(1, 8):
            ws.cell(row=row, column=col).border = thin_border

    total_row = start_row + len(items)

    ws.cell(row=total_row, column=1, value="合    計").alignment = center_align

    if items:
        ws.cell(
            row=total_row,
            column=4,
            value=f"=SUM(D{start_row}:D{total_row - 1})",
        ).alignment = right_align
    else:
        ws.cell(row=total_row, column=4, value=0).alignment = right_align

    for col in range(1, 8):
        ws.cell(row=total_row, column=col).border = thin_border

    sign_row = total_row + 2

    roles = ["總經理", "", "出納", "", "會計", "主管", "請款人"]
    for col_idx, role in enumerate(roles, start=1):
        ws.cell(
            row=sign_row,
            column=col_idx,
            value=role,
        ).alignment = center_align

    if applicant:
        ws.cell(
            row=sign_row + 1,
            column=7,
            value=applicant,
        ).alignment = center_align

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
    "上傳發票或收據照片（可多選）",
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
                        # 只有真的要呼叫 API 時才限速。
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
                            "date": "",
                            "invoice_no": "",
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
        "date",
        "invoice_no",
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
            "date": st.column_config.TextColumn("日期"),
            "invoice_no": st.column_config.TextColumn("發票號碼"),
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

    st.caption("⚠️ AI 辨識結果請務必人工確認，尤其是金額、日期與發票號碼。")

    # Excel 只取會計欄位
    excel_records = edited_df[
        [
            "date",
            "invoice_no",
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
### Streamlit Secrets

在 Streamlit Cloud 的 Secrets 中加入：

```toml
GEMINI_API_KEY = "你的 API Key"
```

如果要指定其他模型，也可以加入：

```toml
GEMINI_MODEL = "gemini-3.8-flash"
```

### 檔案

請確定專案至少有：

```text
app.py
requirements.txt
template.xlsx
```

### API 節省機制

這個版本會：

- 使用 SHA-256 判斷同一張圖片
- 相同圖片直接使用 SQLite 快取
- 只有真正呼叫 API 時才做限速
- 429 / 503 等暫時性錯誤才重試
- 使用 Structured Output 降低 JSON 格式錯誤
- 每張正常單據原則上只需要一次 API 呼叫
        """
    )
