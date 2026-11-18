"""예산 요구안 코어 — 엑셀 읽기, 검사, 요약, 엑셀 생성.

웹앱(budget_app/server.py)과 Skill 스크립트(make_budget_request.py)가 함께 쓴다.
외부 통신 없음. openpyxl만 사용.

초안 데이터(dict) 형식은 PRD.md 3-3 참고. 금액 열 이름은 `{year-1} 배정액`, `{year-1} 집행액`.
"""
from __future__ import annotations

import hashlib
import io
from datetime import datetime

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

METHODS = ("계획 산출", "올해 집행액 기준", "계획 없음", "미산정")
METHOD_FILL = {
    "계획 산출": "E2EFDA",
    "올해 집행액 기준": "FFF2CC",
    "계획 없음": "EDEDED",
    "미산정": "FCE4D6",
}
# 열 자동 인식 단어 (PRD 3-1)
AUTO_WORDS = {
    "비목": ("비목",),
    "세부사업": ("세부사업", "사업명"),
    "배정액": ("배정", "예산액"),
    "집행액": ("집행",),
}
REQUIRED = tuple(AUTO_WORDS)
MAX_HEADER_ROW = 10


# ───────────────────────── 공통 유틸 ─────────────────────────
def year_of(data: dict) -> int:
    try:
        return int(data.get("year") or 2027)
    except (TypeError, ValueError):
        return 2027


def keys(data: dict) -> tuple[str, str]:
    y = year_of(data) - 1
    return f"{y} 배정액", f"{y} 집행액"


def _num(v):
    if v is None or v == "":
        return None
    if isinstance(v, bool):
        return None
    if isinstance(v, (int, float)):
        return v
    s = str(v).replace(",", "").strip()
    try:
        f = float(s)
    except ValueError:
        return None
    return int(f) if f.is_integer() else f


def has_value(v) -> bool:
    return v is not None and v != ""


def items_of(r: dict) -> list[dict]:
    """계획 산출의 산출 줄 목록 [{항목, 단가, 수량, 횟수, 단위}]. 없으면 빈 목록(단가·수량 한 줄로 처리)."""
    items = r.get("산출")
    if r.get("산출방식") == "계획 산출" and isinstance(items, list) and items:
        return [it for it in items if isinstance(it, dict)]
    return []


def item_amount(it: dict):
    """단가 × 수량 × 횟수(비면 1). 단가·수량이 비거나 숫자가 아니면 None."""
    p, q = _num(it.get("단가")), _num(it.get("수량"))
    if p is None or q is None:
        return None
    t = _num(it.get("횟수"))
    return p * q * (1 if t is None else t)


def request_amount(r: dict):
    items = items_of(r)
    if items:
        amts = [item_amount(it) for it in items]
        return None if any(a is None for a in amts) else sum(amts)
    if has_value(r.get("단가")) and has_value(r.get("수량")):
        p, q = _num(r["단가"]), _num(r["수량"])
        return None if p is None or q is None else p * q
    return None


def threshold_of(data: dict):
    """증감 사유 점검 기준(%). 비어 있으면 None → 증감이 있는 모든 사업을 본다."""
    t = _num(data.get("reason_threshold"))
    return None if t is None or t < 0 else t


def needs_reason(r: dict, alloc, req, thr) -> bool:
    """요구 사유가 꼭 있어야 하는 사업인지 (C7). 미산정·금액 미정은 제외."""
    if r.get("산출방식") == "미산정" or req is None:
        return False
    alloc = alloc or 0
    if alloc == 0:
        return req != 0
    diff = req - alloc
    if diff == 0:
        return False
    return thr is None or abs(diff) / alloc * 100 >= thr


# ───────────────────────── 엑셀 읽기 ─────────────────────────
def parse_excel(blob: bytes, header_row: int | None = None, mapping: dict | None = None) -> dict:
    """올해 예산·집행 엑셀을 읽는다.

    반환: header_row, columns(머리글 목록), mapping(필수 열 → 열 번호, 1부터), preview(머리글+앞 5행),
          rows(세부사업 목록), errors, source_hash
    mapping이 비었거나 일부만 있으면 자동 인식으로 채운다. 못 찾은 열은 None.
    """
    out = {"header_row": None, "columns": [], "mapping": {}, "preview": [], "rows": [],
           "errors": [], "source_hash": hashlib.sha256(blob).hexdigest()}
    if not blob.startswith(b"PK"):
        out["errors"].append("엑셀(.xlsx) 파일이 아닙니다.")
        return out
    try:
        wb = load_workbook(io.BytesIO(blob), read_only=True, data_only=True)
    except Exception:
        out["errors"].append("엑셀 파일을 열 수 없습니다. 파일이 손상되었거나 암호가 걸려 있을 수 있습니다.")
        return out
    if {"요약", "산출내역", "확인 필요"} <= set(wb.sheetnames):
        wb.close()
        out["errors"].append("이 파일은 이 웹앱(또는 Skill)이 만든 '요구안 결과' 엑셀입니다. "
                             "①단계에는 올해 예산·집행 현황 엑셀(비목·세부사업·배정액·집행액이 있는 표)을 올려 주세요.")
        return out
    ws = wb.worksheets[0]
    grid = [list(r) for r in ws.iter_rows(values_only=True)]
    wb.close()
    if not grid:
        out["errors"].append("첫 번째 시트가 비어 있습니다.")
        return out

    def text(v):
        return "" if v is None else str(v).strip()

    # 머리글 행 찾기: 지정값 우선, 없으면 '세부사업'/'사업명'이 있는 첫 행
    if header_row is None:
        for i, row in enumerate(grid[:MAX_HEADER_ROW], 1):
            if any(w in text(c) for c in row for w in AUTO_WORDS["세부사업"]):
                header_row = i
                break
    if header_row is None:
        header_row = 1
        out["errors"].append("머리글 행을 자동으로 찾지 못했습니다. 머리글 행을 지정해 주세요.")
    header_row = max(1, min(int(header_row), len(grid)))
    out["header_row"] = header_row
    header = [text(c) for c in grid[header_row - 1]]
    out["columns"] = header
    out["preview"] = [[text(c) for c in r] for r in grid[header_row - 1: header_row + 5]]

    mapping = dict(mapping or {})
    for key, words in AUTO_WORDS.items():
        if mapping.get(key):
            continue
        mapping[key] = next((i + 1 for i, h in enumerate(header) if any(w in h for w in words)), None)
    out["mapping"] = mapping
    missing = [k for k in REQUIRED if not mapping.get(k)]
    if missing:
        out["errors"].append("필수 열을 찾지 못했습니다: " + ", ".join(missing) + " — 열 맞추기에서 지정해 주세요.")
        return out

    seen = set()
    for n, row in enumerate(grid[header_row:], header_row + 1):
        cell = lambda k: row[mapping[k] - 1] if mapping[k] - 1 < len(row) else None
        name = text(cell("세부사업"))
        if not name:
            if any(text(c) for c in row):
                out["errors"].append(f"{n}행: 세부사업 이름이 비어 있습니다.")
            continue
        alloc, spent = _num(cell("배정액")), _num(cell("집행액"))
        for label, raw, val in (("배정액", cell("배정액"), alloc), ("집행액", cell("집행액"), spent)):
            if has_value(raw) and val is None:
                out["errors"].append(f"{n}행 '{name}': {label}이 숫자가 아닙니다 ({raw}).")
        if name in seen:
            out["errors"].append(f"{n}행: 세부사업 '{name}'이 중복됩니다.")
        seen.add(name)
        out["rows"].append({"비목": text(cell("비목")) or "미분류", "세부사업": name,
                            "배정액": alloc or 0, "집행액": spent or 0})
    if not out["rows"]:
        out["errors"].append("세부사업 행이 없습니다.")
    return out


def original_from_parse(parsed: dict) -> list[dict]:
    return [{"세부사업": r["세부사업"], "비목": r["비목"], "배정액": r["배정액"], "집행액": r["집행액"]}
            for r in parsed["rows"]]


# ───────────────────────── 검사 ─────────────────────────
def validate(data: dict) -> list[dict]:
    """C1~C4, C7. 문제 목록 [{code, row(1부터 또는 None), msg}]"""
    probs = []
    ka, _ = keys(data)
    thr = threshold_of(data)
    seen = {}
    for i, r in enumerate(data.get("rows", []), 1):
        nm = str(r.get("세부사업") or "").strip()
        if nm and nm in seen:
            probs.append({"code": "C1", "row": i, "msg": f"'{nm}': 세부사업 이름이 {seen[nm]}번 행과 중복됩니다."})
        seen.setdefault(nm, i)
    for i, r in enumerate(data.get("rows", []), 1):
        name = r.get("세부사업") or f"{i}번 행"
        m = r.get("산출방식")
        if m not in METHODS:
            probs.append({"code": "C1", "row": i, "msg": f"'{name}': 산출방식 값이 올바르지 않습니다 ({m})."})
        raw_items = r.get("산출") if isinstance(r.get("산출"), list) else []
        if m == "미산정" and (has_value(r.get("단가")) or has_value(r.get("수량"))
                              or any(has_value(it.get("단가")) or has_value(it.get("수량"))
                                     for it in raw_items if isinstance(it, dict))):
            probs.append({"code": "C2", "row": i, "msg": f"'{name}': 미산정인데 단가·수량이 들어 있습니다."})
        if m == "계획 산출":
            items = items_of(r)
            if items:
                for k, it in enumerate(items, 1):
                    if item_amount(it) is None:
                        label = f" ({it.get('항목')})" if str(it.get("항목") or "").strip() else ""
                        probs.append({"code": "C1", "row": i,
                                      "msg": f"'{name}' 산출 {k}번째 줄{label}: 단가 또는 수량이 비어 있거나 숫자가 아닙니다."})
            elif request_amount(r) is None:
                probs.append({"code": "C1", "row": i, "msg": f"'{name}': 계획 산출인데 단가 또는 수량이 비어 있습니다."})
        if m in METHODS and needs_reason(r, _num(r.get(ka)), request_amount(r), thr) \
                and not str(r.get("요구 사유") or "").strip():
            basis = "신규 요구인데" if not (_num(r.get(ka)) or 0) else (
                "증감이 있는데" if thr is None else f"증감 {thr:g}% 이상인데")
            probs.append({"code": "C7", "row": i, "msg": f"'{name}': {basis} 요구 사유가 비어 있습니다."})
        if m in METHODS and m != "계획 산출" and not str(r.get("확인 필요") or "").strip():
            probs.append({"code": "C3", "row": i, "msg": f"'{name}': {m}인데 확인 필요가 비어 있습니다."})
        if not str(r.get("근거 자료") or "").strip():
            probs.append({"code": "C4", "row": i, "msg": f"'{name}': 근거 자료가 비어 있습니다."})
        if not str(r.get("세부사업") or "").strip():
            probs.append({"code": "C1", "row": i, "msg": f"{i}번 행: 세부사업 이름이 비어 있습니다."})
    return probs


def compare_original(data: dict, original: list[dict]) -> list[dict]:
    """C5(원본 세부사업 누락), C6(원본 금액 불일치·원본에 없는 사업에 금액)."""
    ka, ke = keys(data)
    probs = []
    rows = data.get("rows", [])
    by_name = {}
    for i, r in enumerate(rows, 1):
        by_name.setdefault(r.get("세부사업"), (i, r))
    names = set()
    for o in original:
        names.add(o["세부사업"])
        hit = by_name.get(o["세부사업"])
        if hit is None:
            probs.append({"code": "C5", "row": None, "msg": f"원본 세부사업 '{o['세부사업']}'이 요구안에 없습니다."})
            continue
        i, r = hit
        for label, k, ov in (("배정액", ka, o["배정액"]), ("집행액", ke, o["집행액"])):
            if (_num(r.get(k)) or 0) != (ov or 0):
                probs.append({"code": "C6", "row": i,
                              "msg": f"'{o['세부사업']}' {label}: 원본 {ov:,} ≠ 요구안 {_num(r.get(k)) or 0:,}"})
    for i, r in enumerate(rows, 1):
        if r.get("세부사업") not in names and ((_num(r.get(ka)) or 0) or (_num(r.get(ke)) or 0)):
            probs.append({"code": "C6", "row": i,
                          "msg": f"'{r.get('세부사업')}': 원본에 없는 사업(신규)인데 올해 배정·집행액이 있습니다."})
    return probs


def check(data: dict, original: list[dict] | None = None) -> dict:
    """전체 검사 + 요약. original을 안 주면 data['original'] 스냅샷을 쓴다."""
    if original is None:
        original = data.get("original")
    probs = validate(data)
    warnings = []
    if original is not None:
        probs += compare_original(data, original)
    else:
        warnings.append("원본 값 정보가 없어 세부사업 누락·금액 대조(C5·C6)를 하지 못했습니다.")
    if original is not None and not data.get("source_verified"):
        warnings.append("원본 파일 미대조 — 초안에 저장된 원본 값으로만 검사했습니다.")
    notes = reason_notes(data)
    if notes:
        warnings.append(f"요구 사유가 자동 문구 그대로인 사업 {len(notes)}건 — 증감 사유를 보강해 주세요.")
    return {"ok": not probs, "problems": probs, "warnings": warnings, "reason_notes": notes,
            "summary": summarize(data)}


def reason_status(data: dict, r: dict) -> str:
    """사업별 요구 사유 점검 결과: '' (대상 아님·충분) / '사유 없음' / '보강 필요'(자동 문구 그대로)."""
    ka, _ = keys(data)
    if not needs_reason(r, _num(r.get(ka)), request_amount(r), threshold_of(data)):
        return ""
    reason = str(r.get("요구 사유") or "").strip()
    if not reason:
        return "사유 없음"
    if r.get("_autoReason") and reason == str(r.get("_autoReason")).strip():
        return "보강 필요"
    return ""


def reason_notes(data: dict) -> list[dict]:
    """경고(내보내기는 가능): 증감 기준을 넘는데 요구 사유가 자동 문구 그대로인 사업."""
    return [{"row": i, "msg": f"'{r.get('세부사업')}': 요구 사유가 자동 문구 그대로입니다 — 증감 사유 보강 필요"}
            for i, r in enumerate(data.get("rows", []), 1) if reason_status(data, r) == "보강 필요"]


CHECKS = [
    ("C1", "산출방식·필수 값"),
    ("C2", "미산정 사업에 금액 추측 없음"),
    ("C3", "확인 필요 표시"),
    ("C4", "근거 자료 표시"),
    ("C5", "원본 세부사업 누락 없음"),
    ("C6", "원본 금액 일치"),
    ("C7", "증감 사유 표시"),
]


# ───────────────────────── 요약 ─────────────────────────
def summarize(data: dict) -> dict:
    ka, _ = keys(data)
    rows = data.get("rows", [])
    cats = {}
    for r in rows:
        c = cats.setdefault(r.get("비목") or "미분류",
                            {"비목": r.get("비목") or "미분류", "배정": 0, "요구": 0, "미산정": 0, "확인필요": 0})
        c["배정"] += _num(r.get(ka)) or 0
        c["요구"] += request_amount(r) or 0
        c["미산정"] += r.get("산출방식") == "미산정"
        c["확인필요"] += bool(str(r.get("확인 필요") or "").strip())
    tot_a = sum(c["배정"] for c in cats.values())
    tot_r = sum(c["요구"] for c in cats.values())
    by_project = []
    for i, r in enumerate(rows, 1):
        a, q = _num(r.get(ka)) or 0, request_amount(r)
        by_project.append({"번호": i, "비목": r.get("비목") or "미분류", "세부사업": r.get("세부사업"),
                           "산출방식": r.get("산출방식"), "배정": a, "요구": q,
                           "증감": None if q is None else q - a,
                           "증감률": None if q is None or not a else (q - a) / a,
                           "사유점검": reason_status(data, r)})
    return {
        "by_category": list(cats.values()),
        "by_project": by_project,
        "reason_threshold": threshold_of(data),
        "배정": tot_a, "요구": tot_r, "증감": tot_r - tot_a,
        "증감률": (tot_r - tot_a) / tot_a if tot_a else None,
        "미산정": sum(c["미산정"] for c in cats.values()),
        "확인필요": sum(c["확인필요"] for c in cats.values()),
        "methods": {m: sum(r.get("산출방식") == m for r in rows) for m in METHODS},
        "rows": len(rows),
    }


# ───────────────────────── 엑셀 생성 ─────────────────────────
FONT = "맑은 고딕"
HEAD_FILL = PatternFill("solid", fgColor="1F4E78")
THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
NUM = "#,##0"


def _head(ws, row, cols):
    for i, (name, width) in enumerate(cols, 1):
        c = ws.cell(row=row, column=i, value=name)
        c.font = Font(name=FONT, bold=True, color="FFFFFF")
        c.fill = HEAD_FILL
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = BORDER
        ws.column_dimensions[c.column_letter].width = width


def _cell(ws, r, c, v, num=False, center=False, bold=False):
    x = ws.cell(row=r, column=c, value=v)
    x.font = Font(name=FONT, size=10, bold=bold)
    x.border = BORDER
    x.alignment = Alignment(vertical="top", wrap_text=True,
                            horizontal="right" if num else ("center" if center else "left"))
    if num:
        x.number_format = NUM
    return x


def _safe_text(v):
    """엑셀 수식 주입 방지: 글자 칸이 = + - @ 로 시작하면 앞에 ' 를 붙인다."""
    if isinstance(v, str) and v[:1] in ("=", "+", "-", "@"):
        return "'" + v
    return v


def _fmt(v) -> str:
    n = _num(v)
    return "" if n is None else f"{n:,}"


def item_formula_text(it: dict) -> str:
    """'강사료 300,000 × 2명 × 10회 = 6,000,000' 형태의 산출식 글자."""
    unit = str(it.get("단위") or "")
    parts = [_fmt(it.get("단가")), _fmt(it.get("수량")) + unit]
    if _num(it.get("횟수")) not in (None, 1):
        parts.append(_fmt(it.get("횟수")) + "회")
    amt = item_amount(it)
    return " × ".join(parts) + ("" if amt is None else f" = {amt:,}")


def build_workbook(data: dict, result: dict | None = None) -> Workbook:
    """초안 데이터 → 요약·산출내역·확인 필요 3시트 엑셀. 금액은 수식."""
    result = result or check(data)
    rows = data["rows"]
    unit = data.get("unit", "")
    year = year_of(data)
    ka, ke = keys(data)
    req_label, a_label, e_label = f"{year} 요구액", ka, ke

    wb = Workbook()
    summ = wb.active
    summ.title = "요약"
    proj = wb.create_sheet("사업별 요약")
    det = wb.create_sheet("산출내역")
    its = wb.create_sheet("산출 세부")
    chk = wb.create_sheet("확인 필요")
    title = data.get("title") or f"{year}년 예산 요구안 초안"

    cols = [("번호", 6), ("비목", 11), ("세부사업", 22), ("산출방식", 15), ("단가", 12), ("수량", 7),
            ("단위", 7), (req_label, 14), (a_label, 14), (e_label, 14), ("산출 근거", 26),
            ("요구 사유", 30), ("근거 자료", 34), ("확인 필요", 38)]
    det["A1"] = f"{title} — 산출내역 (단위: {unit})"
    det["A1"].font = Font(name=FONT, size=13, bold=True)
    det["A2"] = (f"{data.get('base_note', '')}  |  생성: {datetime.now():%Y-%m-%d %H:%M}  |  "
                 "초안입니다. 모든 금액은 담당자 확정 전입니다.")
    det["A2"].font = Font(name=FONT, size=9, color="7F7F7F")
    # 산출 세부 시트 (계획 산출의 여러 줄): 사업번호로 산출내역과 연결
    item_lines = [(i + 1, r, it) for i, r in enumerate(rows) for it in items_of(r)]
    icols = [("사업번호", 8), ("세부사업", 22), ("항목", 20), ("단가", 12), ("수량", 8), ("횟수", 8),
             ("단위", 10), ("산출식", 30), ("금액", 14)]
    its["A1"] = f"{title} — 산출 세부 (계획 산출의 산출 줄, 단위: {unit})"
    its["A1"].font = Font(name=FONT, size=13, bold=True)
    its["A2"] = "금액 = 단가 × 수량 × 횟수(비면 1). 산출내역 시트의 요구액은 사업번호별 금액 합계입니다."
    its["A2"].font = Font(name=FONT, size=9, color="7F7F7F")
    _head(its, 4, icols)
    for k, (no, r, it) in enumerate(item_lines):
        y = 5 + k
        _cell(its, y, 1, no, center=True)
        _cell(its, y, 2, _safe_text(r.get("세부사업")))
        _cell(its, y, 3, _safe_text(it.get("항목") or ""))
        _cell(its, y, 4, _num(it.get("단가")), num=True)
        _cell(its, y, 5, _num(it.get("수량")), num=True)
        _cell(its, y, 6, _num(it.get("횟수")), num=True)
        _cell(its, y, 7, _safe_text(it.get("단위") or ""), center=True)
        _cell(its, y, 8, _safe_text(item_formula_text(it)))
        _cell(its, y, 9, f'=IF(OR(D{y}="",E{y}=""),"",D{y}*E{y}*IF(F{y}="",1,F{y}))', num=True, bold=True)
    ilast = 4 + max(len(item_lines), 1)
    its.freeze_panes = "C5"
    irng = lambda col: f"'산출 세부'!${col}$5:${col}${ilast}"

    hr = 4
    _head(det, hr, cols)
    first = hr + 1
    for i, r in enumerate(rows):
        x = first + i
        _cell(det, x, 1, i + 1, center=True)
        _cell(det, x, 2, _safe_text(r.get("비목")))
        _cell(det, x, 3, _safe_text(r.get("세부사업")))
        m = _cell(det, x, 4, r.get("산출방식"), center=True)
        m.fill = PatternFill("solid", fgColor=METHOD_FILL.get(r.get("산출방식"), "FFFFFF"))
        items = items_of(r)
        if items:
            _cell(det, x, 5, None, num=True)
            _cell(det, x, 6, None, num=True)
            _cell(det, x, 7, f"세부 {len(items)}줄", center=True)
            _cell(det, x, 8, f'=SUMIF({irng("A")},A{x},{irng("I")})', num=True, bold=True)
        else:
            _cell(det, x, 5, _num(r.get("단가")), num=True)
            _cell(det, x, 6, _num(r.get("수량")), num=True)
            _cell(det, x, 7, _safe_text(r.get("단위")), center=True)
            _cell(det, x, 8, f'=IF(OR(E{x}="",F{x}=""),"",E{x}*F{x})', num=True, bold=True)
        _cell(det, x, 9, _num(r.get(ka)) or 0, num=True)
        _cell(det, x, 10, _num(r.get(ke)) or 0, num=True)
        for c, k in zip(range(11, 15), ("산출 근거", "요구 사유", "근거 자료", "확인 필요")):
            _cell(det, x, c, _safe_text(r.get(k, "")))
    last = first + max(len(rows), 1) - 1
    det.freeze_panes = det.cell(row=first, column=4)
    det.auto_filter.ref = f"A{hr}:N{last}"

    summ["A1"] = f"{title} — 비목별 요약 (단위: {unit})"
    summ["A1"].font = Font(name=FONT, size=13, bold=True)
    summ["A2"] = "요구액은 산출내역 시트의 수식 합계입니다. 미산정 사업은 요구액에 포함되지 않습니다."
    summ["A2"].font = Font(name=FONT, size=9, color="C00000")
    scols = [("비목", 14), (a_label, 16), (f"{req_label}(초안)", 18), ("증감", 14), ("증감률", 10),
             ("미산정 건수", 12), ("확인 필요 건수", 14)]
    _head(summ, 4, scols)
    cats = list(dict.fromkeys(r.get("비목") for r in rows))
    rng = lambda col: f"산출내역!${col}${first}:${col}${last}"
    for i, cat in enumerate(cats):
        x = 5 + i
        _cell(summ, x, 1, _safe_text(cat))
        _cell(summ, x, 2, f'=SUMIF({rng("B")},A{x},{rng("I")})', num=True)
        _cell(summ, x, 3, f'=SUMIF({rng("B")},A{x},{rng("H")})', num=True, bold=True)
        _cell(summ, x, 4, f"=C{x}-B{x}", num=True)
        p = _cell(summ, x, 5, f'=IF(B{x}=0,"신규",D{x}/B{x})', num=True)
        p.number_format = "0.0%"
        _cell(summ, x, 6, f'=COUNTIFS({rng("B")},A{x},{rng("D")},"미산정")', num=True)
        _cell(summ, x, 7, f'=COUNTIFS({rng("B")},A{x},{rng("N")},"<>")', num=True)
    t = 5 + len(cats)
    _cell(summ, t, 1, "합계", bold=True)
    for c, col in zip(range(2, 8), "BCDEFG"):
        if col == "E":
            p = _cell(summ, t, c, f'=IF(B{t}=0,"",D{t}/B{t})', num=True, bold=True)
            p.number_format = "0.0%"
        else:
            _cell(summ, t, c, f"=SUM({col}5:{col}{t - 1})", num=True, bold=True)
    for c in range(1, 8):
        summ.cell(row=t, column=c).fill = PatternFill("solid", fgColor="DDEBF7")

    m0 = t + 2
    _cell(summ, m0, 1, "산출방식", bold=True)
    _cell(summ, m0, 2, "건수", bold=True)
    for i, m in enumerate(METHODS, 1):
        _cell(summ, m0 + i, 1, m).fill = PatternFill("solid", fgColor=METHOD_FILL[m])
        _cell(summ, m0 + i, 2, f'=COUNTIF({rng("D")},A{m0 + i})', num=True)

    s0 = m0 + len(METHODS) + 2
    probs = result["problems"]
    thr = threshold_of(data)
    info = [
        ("입력 자료", "\n".join(data.get("sources", []))),
        ("증감 사유 기준", "증감이 있는 모든 사업" if thr is None else f"전년 배정 대비 ±{thr:g}% 이상 (신규 사업 포함)"),
        ("자동 검사", "통과" if not probs else "실패: " + " / ".join(p["msg"] for p in probs)),
        ("경고", "\n".join(result.get("warnings", [])) or "없음"),
    ]
    if data.get("started_at"):
        info.append(("작업 시간", f"시작 {data.get('started_at')} → 내보내기 {data.get('exported_at') or '-'}"))
    for k, (label, val) in enumerate(info):
        r = s0 + k
        _cell(summ, r, 1, label, bold=True)
        c = _cell(summ, r, 2, _safe_text(val))
        if label == "경고" and val != "없음":
            c.font = Font(name=FONT, size=10, bold=True, color="C00000")
        summ.merge_cells(start_row=r, start_column=2, end_row=r, end_column=7)
        summ.row_dimensions[r].height = 15 * max(1, str(val).count("\n") + 1)

    # 사업별 요약: 산출내역 같은 행을 참조하는 수식
    pcols = [("번호", 6), ("비목", 11), ("세부사업", 22), ("산출방식", 15), (a_label, 14),
             (f"{req_label}(초안)", 16), ("증감", 14), ("증감률", 9), ("요구 사유", 40), ("사유 점검", 11)]
    proj["A1"] = f"{title} — 사업별 요약 (단위: {unit})"
    proj["A1"].font = Font(name=FONT, size=13, bold=True)
    proj["A2"] = ("사유 점검: '사유 없음' = 기준을 넘는데 요구 사유가 비어 있음 / "
                  "'보강 필요' = 자동 문구 그대로. 미산정 사업은 요구액이 비어 있습니다.")
    proj["A2"].font = Font(name=FONT, size=9, color="7F7F7F")
    _head(proj, 4, pcols)
    for i, r in enumerate(rows):
        y, x = 5 + i, first + i
        _cell(proj, y, 1, i + 1, center=True)
        _cell(proj, y, 2, _safe_text(r.get("비목")))
        _cell(proj, y, 3, _safe_text(r.get("세부사업")))
        _cell(proj, y, 4, r.get("산출방식"), center=True).fill = PatternFill(
            "solid", fgColor=METHOD_FILL.get(r.get("산출방식"), "FFFFFF"))
        _cell(proj, y, 5, f"=산출내역!I{x}", num=True)
        _cell(proj, y, 6, f"=산출내역!H{x}", num=True, bold=True)
        _cell(proj, y, 7, f'=IF(F{y}="","",F{y}-E{y})', num=True)
        p = _cell(proj, y, 8, f'=IF(F{y}="","",IF(E{y}=0,"신규",G{y}/E{y}))', num=True)
        p.number_format = "0.0%"
        _cell(proj, y, 9, _safe_text(r.get("요구 사유", "")))
        st = reason_status(data, r)
        c = _cell(proj, y, 10, st, center=True)
        if st:
            c.font = Font(name=FONT, size=10, bold=True, color="C00000" if st == "사유 없음" else "B26A00")
    t2 = 5 + len(rows)
    _cell(proj, t2, 1, "합계", bold=True)
    for c in range(2, 5):
        _cell(proj, t2, c, None)
    _cell(proj, t2, 5, f"=SUM(E5:E{t2 - 1})", num=True, bold=True)
    _cell(proj, t2, 6, f"=SUM(F5:F{t2 - 1})", num=True, bold=True)
    _cell(proj, t2, 7, f"=F{t2}-E{t2}", num=True, bold=True)
    p = _cell(proj, t2, 8, f'=IF(E{t2}=0,"",G{t2}/E{t2})', num=True, bold=True)
    p.number_format = "0.0%"
    for c in range(1, 11):
        proj.cell(row=t2, column=c).fill = PatternFill("solid", fgColor="DDEBF7")
    proj.freeze_panes = "D5"
    proj.auto_filter.ref = f"A4:J{t2 - 1}"

    ccols = [("번호", 6), ("비목", 11), ("세부사업", 22), ("산출방식", 15), ("확인 필요", 60), ("근거 자료", 34)]
    chk["A1"] = "확인 필요 항목"
    chk["A1"].font = Font(name=FONT, size=13, bold=True)
    _head(chk, 3, ccols)
    n = 0
    for i, r in enumerate(rows, 1):
        if str(r.get("확인 필요") or "").strip():
            n += 1
            for c, v in enumerate((i, r.get("비목"), r.get("세부사업"), r.get("산출방식"),
                                   r.get("확인 필요"), r.get("근거 자료")), 1):
                _cell(chk, 3 + n, c, _safe_text(v), center=c in (1, 4))
    return wb


def workbook_bytes(data: dict, result: dict | None = None) -> bytes:
    buf = io.BytesIO()
    build_workbook(data, result).save(buf)
    return buf.getvalue()
