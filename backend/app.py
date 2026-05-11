"""
MVP: 웹 ↔ 대상 시트 직접 연동 (회의서식 시트 미사용).
- 사전 조건: 워크북이 Excel에서 이미 열려있을 것
- 헤더 매핑은 매 호출 시 시트 2행에서 동적 로드 → contracts 파일 없음
- Excel COM은 단일 인스턴스 가정 → 모든 endpoint는 EXCEL_LOCK으로 직렬화
실행: python backend/app.py  →  http://127.0.0.1:8765/docs
"""
import os
import re
import threading
from datetime import datetime
from typing import Any

import xlwings as xw
from fastapi import FastAPI, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel, Field

WORKBOOK = "demo+난수_회의서식_매크로_2026년 모니터링 엑셀서식_260428.xlsm"
TARGET = "대상"
HISTORY = "대상(2025)"
HEADER_ROW = 2
DATA_ROW = 3
NCOLS = 47
CASE_HDR = "연번"
DATE_HDR = "보고일"  # ← 실제 헤더명 다르면 알려주세요

# /review에서 쓸 수 있는 헤더는 이 3개로 제한
ALLOWED_REVIEW_KEYS = ("질환 및 발생형태", "업무관련성", "유해인자분류", "유해인자")

# Excel COM은 STA — 동시 호출 시 RPC 충돌 가능. 모든 Excel 접근은 이 락으로 직렬화.
EXCEL_LOCK = threading.Lock()

app = FastAPI(title="Excel MVP Bridge")
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],          # file:// 포함 모든 origin 허용 (로컬 단일 사용자 환경)
    allow_credentials=False,
    allow_methods=["*"],
    allow_headers=["*"],
)


def get_book() -> xw.Book:
    """이름(b.name) 또는 전체 경로(b.fullname) 어느 쪽이든 일치하면 attach."""
    target_name = os.path.basename(WORKBOOK)
    target_full = os.path.normcase(os.path.abspath(WORKBOOK)) if os.path.isabs(WORKBOOK) else None
    for a in xw.apps:
        for b in a.books:
            if b.name == target_name:
                return b
            if target_full and os.path.normcase(b.fullname) == target_full:
                return b
    raise HTTPException(503, f"'{WORKBOOK}' 미열림")


def headers_of(ws) -> dict[str, int]:
    row = ws.range((HEADER_ROW, 1), (HEADER_ROW, NCOLS)).value
    return {str(v).strip(): i + 1 for i, v in enumerate(row) if v not in (None, "")}


def find_row(ws, case_no, col) -> int | None:
    last = ws.used_range.last_cell.row
    target = str(case_no).strip()
    for r in range(DATA_ROW, last + 1):
        v = ws.cells(r, col).value
        if v is None:
            continue
        norm = str(int(v)) if isinstance(v, float) and v.is_integer() else str(v).strip()
        if norm == target:
            return r
    return None


def row_to_dict(ws, row, hdrs) -> dict:
    vals = ws.range((row, 1), (row, NCOLS)).value
    return {h: vals[c - 1] for h, c in hdrs.items()}


_DATE_FORMATS = ("%Y-%m-%d", "%Y.%m.%d", "%Y/%m/%d", "%Y%m%d", "%Y-%m-%d %H:%M:%S")


def parse_report_date(d: Any) -> datetime:
    """보고일 정렬용 — datetime이면 그대로, 문자열이면 흔한 포맷 시도, 실패 시 datetime.min."""
    if isinstance(d, datetime):
        return d
    if isinstance(d, str):
        s = d.strip()
        if not s:
            return datetime.min
        for fmt in _DATE_FORMATS:
            try:
                return datetime.strptime(s, fmt)
            except ValueError:
                pass
        # 숫자만 추출해서 YYYYMMDD 시도
        digits = re.sub(r"\D", "", s)
        if len(digits) == 8:
            try:
                return datetime.strptime(digits, "%Y%m%d")
            except ValueError:
                pass
    return datetime.min


# ─── Pydantic 모델 ────────────────────────────────────────────────────────────

class ReviewPayload(BaseModel):
    연번: str | int = Field(..., description="대상 시트 A열의 연번 (숫자/문자 모두 허용)")
    updates: dict[str, Any] = Field(
        ...,
        description=f"업데이트할 헤더명 → 값. 허용 키: {list(ALLOWED_REVIEW_KEYS)}",
        min_length=1,
    )

    model_config = {
        "json_schema_extra": {
            "examples": [
                {
                    "연번": 79,
                    "updates": {
                        "질환 및 발생형태": "석면폐",
                        "업무관련성": "Probable",
                        "유해인자분류": "석면",
                    },
                }
            ]
        }
    }


# ─── Endpoints ────────────────────────────────────────────────────────────────

# 1) 사례 1건 조회
@app.get("/case/{case_no}")
def get_case(case_no: str):
    with EXCEL_LOCK:
        ws = get_book().sheets[TARGET]
        hdrs = headers_of(ws)
        if CASE_HDR not in hdrs:
            raise HTTPException(500, f"'{CASE_HDR}' 헤더 없음")
        row = find_row(ws, case_no, hdrs[CASE_HDR])
        if row is None:
            raise HTTPException(404, f"연번 {case_no} 없음")
        return {"_row": row, **row_to_dict(ws, row, hdrs)}


# 2) 동일 연번 행의 지정 셀만 업데이트 (allowed key whitelist)
@app.post("/review")
def save_review(payload: ReviewPayload):
    bad = [k for k in payload.updates if k not in ALLOWED_REVIEW_KEYS]
    if bad:
        raise HTTPException(
            400,
            f"허용되지 않은 키: {bad}. 허용 키: {list(ALLOWED_REVIEW_KEYS)}",
        )
    with EXCEL_LOCK:
        ws = get_book().sheets[TARGET]
        hdrs = headers_of(ws)
        if CASE_HDR not in hdrs:
            raise HTTPException(500, f"'{CASE_HDR}' 헤더 없음")
        row = find_row(ws, payload.연번, hdrs[CASE_HDR])
        if row is None:
            raise HTTPException(404, f"연번 {payload.연번} 없음")
        missing_in_sheet = [k for k in payload.updates if k not in hdrs]
        if missing_in_sheet:
            raise HTTPException(
                500,
                f"시트에 헤더 없음: {missing_in_sheet} (시트={TARGET}, 헤더행={HEADER_ROW})",
            )
        written = {}
        for name, val in payload.updates.items():
            ws.cells(row, hdrs[name]).value = val
            written[name] = val
        return {"ok": True, "row": row, "updated": written}


# 3) 유사사례 검색 — 대상 + 대상(2025), 보고일 최신순, 상위 N개
@app.get("/search")
def search_similar(field: str, keyword: str, limit: int = 10):
    with EXCEL_LOCK:
        wb = get_book()
        out = []
        for sheet in (TARGET, HISTORY):
            ws = wb.sheets[sheet]
            hdrs = headers_of(ws)
            f_col = hdrs.get(field)
            if f_col is None:
                continue
            last = ws.used_range.last_cell.row
            if last < DATA_ROW:
                continue
            block = ws.range((DATA_ROW, 1), (last, NCOLS)).value
            if block is None:
                continue
            if not isinstance(block[0], list):
                block = [block]
            for vals in block:
                cell = vals[f_col - 1]
                if cell is None:
                    continue
                if keyword in str(cell):
                    rec = {h: vals[c - 1] for h, c in hdrs.items()}
                    rec["_sheet"] = sheet
                    out.append(rec)
        out.sort(key=lambda r: parse_report_date(r.get(DATE_HDR)), reverse=True)
        return out[:limit]


# 기존 hardcoded write (보존)
LEGACY_SHEET, LEGACY_CELL, LEGACY_VALUE = "회의서식", "C5", 81


@app.post("/write")
def write():
    with EXCEL_LOCK:
        wb = get_book()
        wb.sheets[LEGACY_SHEET].range(LEGACY_CELL).value = LEGACY_VALUE
        return {"ok": True, "workbook": wb.name, "sheet": LEGACY_SHEET, "cell": LEGACY_CELL, "value": LEGACY_VALUE}


if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="127.0.0.1", port=8765)
