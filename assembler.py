"""합본 xlsx 조립 및 포맷 적용."""

import re
from pathlib import Path
import openpyxl
from openpyxl.styles import (
    PatternFill, Font, Alignment, Border, Side
)
from openpyxl.utils import get_column_letter


# ── 색상 상수 ────────────────────────────────────────────────────────────────
HDR_BG = "1F3864"
HDR_FG = "FFFFFF"
DIAL_BG = "FFFDE7"
CHAR_BG = "92D050"
REC_BG  = "EA9999"
EMO_BG  = "FFE599"
ROW_ALT = "F5F5F5"
TIMING_BG = "FFE5CC"
PEAK_FG = "FF0000"
SUMMARY_BG_HDR = "1F3864"
SUMMARY_FG_HDR = "FFFFFF"
SUMMARY_TOTAL_BG = "D9E1F2"
SUMMARY_CAST_BG = "FFF2CC"

# 테두리
_THIN   = Side(style="thin",   color="BFBFBF")
_MEDIUM = Side(style="medium", color="999999")
_CELL_BORDER  = Border(left=_THIN, right=_THIN, top=_THIN, bottom=_THIN)
_HDR_BORDER   = Border(left=_MEDIUM, right=_MEDIUM, top=_MEDIUM, bottom=_MEDIUM)

# 탭 그룹 순서 (template_spec.md 기준)
GROUP_ORDER = ["메인", "전투", "PV", "스킬", "캐릭터대본", "짧은음성", "음성", "스킨"]

TYPE_GROUP = {
    "Type_메인": "메인",
    "Type_전투": "전투",
    "Type_PV": "PV",
    "Type_캐릭터음성": "음성",
    "Type_짧은음성": "짧은음성",
    "Type_무한대": "메인",
    "Type_명방캐릭터": "캐릭터대본",
}


def build_hapbon(
    processed: list[dict],
    log_entries: list[dict],
    project_title: str,
    output_path: str,
    optical_lang: str,
):
    wb = openpyxl.Workbook()
    wb.remove(wb.active)  # 기본 시트 제거

    # 그룹별 정렬
    grouped = _group_sheets(processed)

    # 데이터 시트 추가
    all_sheet_data = []  # [(sheet_name, rows)] — 개괄 집계용
    for group in GROUP_ORDER:
        for item in grouped.get(group, []):
            sname = item["sheet_name_out"][:31]
            if item["type"] == "Type_명방캐릭터":
                _add_arknights_sheet(wb, item)
            else:
                rows = item.get("rows") or []
                _add_data_sheet(wb, sname, rows, item["type"])
                all_sheet_data.append((sname, rows, item["type"]))

    # 개괄 시트 (맨 앞)
    _add_summary_sheet(wb, all_sheet_data)

    # 로그 시트 (스킵된 항목 있을 때만)
    if log_entries:
        _add_log_sheet(wb, log_entries)

    # 시트 순서 재배치: 개괄 → 로그 → 데이터
    _reorder_sheets(wb)

    wb.save(output_path)


# ── 그룹핑 ─────────────────────────────────────────────────────────────────

def _group_sheets(processed: list[dict]) -> dict:
    groups: dict[str, list] = {g: [] for g in GROUP_ORDER}
    for item in processed:
        g = TYPE_GROUP.get(item["type"], "메인")
        groups[g].append(item)
    return groups


# ── 데이터 시트 ───────────────────────────────────────────────────────────

# ALT는 항상 맨 끝에 포함 — 데이터 없으면 숨김 처리
_BASE_COLS     = ["파일명", "캐릭터명", "감정", "REC", "대사", "옵티컬(원문)", "⏱ 검수", "ALT"]
_INFINITE_COLS = ["파일명", "캐릭터명", "감정", "ADR/Wild", "타임코드", "REC", "대사", "옵티컬(원문)", "⏱ 검수", "ALT"]

# 중요 열 강조 폰트 크기
_FONT_LARGE = 14   # 캐릭터명, 대사, ALT
_FONT_SMALL = 9    # 나머지


def _get_col_list(sheet_type: str) -> list[str]:
    if sheet_type == "Type_무한대":
        return list(_INFINITE_COLS)
    return list(_BASE_COLS)


def _add_data_sheet(wb: openpyxl.Workbook, sheet_name: str, rows: list[dict], sheet_type: str):
    ws = wb.create_sheet(title=sheet_name)
    if not rows:
        return

    cols = _get_col_list(sheet_type)

    # 헤더
    _write_header(ws, cols)

    # 데이터 행
    for r_idx, row in enumerate(rows):
        if row is None:  # 전투 합산 구분선
            ws.append([""] * len(cols))
            # 구분선 행에도 테두리 적용
            for c_idx in range(1, len(cols) + 1):
                ws.cell(row=r_idx + 2, column=c_idx).border = _CELL_BORDER
            continue
        excel_row = r_idx + 2
        is_even   = (r_idx % 2 == 0)
        is_peak   = row.get("_peak", False)
        is_timing = row.get("_timing_over", False)

        vals = [row.get(c, "") for c in cols]
        ws.append(vals)

        _format_data_row(ws, excel_row, cols, is_even, is_peak, is_timing)

    _set_col_widths(ws, cols)
    ws.freeze_panes = "A2"

    # ALT 열 데이터 없으면 숨김
    alt_idx = cols.index("ALT") + 1  # 1-based
    has_alt = any(r.get("ALT") for r in rows if r is not None)
    if not has_alt:
        ws.column_dimensions[get_column_letter(alt_idx)].hidden = True


def _write_header(ws, cols: list[str]):
    ws.append(cols)
    hdr_fill  = PatternFill("solid", fgColor=HDR_BG)
    hdr_font  = Font(name="맑은 고딕", size=10, bold=True, color=HDR_FG)
    hdr_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for cell in ws[1]:
        cell.fill      = hdr_fill
        cell.font      = hdr_font
        cell.alignment = hdr_align
        cell.border    = _HDR_BORDER
    ws.row_dimensions[1].height = 25


def _format_data_row(ws, row_idx: int, cols: list[str], is_even: bool, is_peak: bool, is_timing: bool):
    base_bg = ROW_ALT if is_even else "FFFFFF"
    if is_timing:
        base_bg = TIMING_BG

    col_bg = {
        "파일명":    base_bg,
        "캐릭터명":  CHAR_BG,
        "감정":      EMO_BG,
        "REC":       REC_BG,
        "대사":      DIAL_BG,
        "ALT":       DIAL_BG,
        "옵티컬(원문)": base_bg,
        "⏱ 검수":   base_bg,
        "ADR/Wild":  EMO_BG,
        "타임코드":  base_bg,
    }

    for c_idx, col_name in enumerate(cols, 1):
        cell = ws.cell(row=row_idx, column=c_idx)
        bg   = col_bg.get(col_name, base_bg)

        cell.fill      = PatternFill("solid", fgColor=bg)
        cell.alignment = Alignment(wrap_text=True, vertical="top")
        cell.border    = _CELL_BORDER

        # 폰트 — 캐릭터명·대사·ALT는 14pt Bold
        is_char  = col_name == "캐릭터명"
        is_dial  = col_name in ("대사", "ALT")
        is_large = is_char or is_dial
        font_color = PEAK_FG if (is_peak and is_dial) else "000000"

        cell.font = Font(
            name="맑은 고딕",
            size=_FONT_LARGE if is_large else _FONT_SMALL,
            bold=is_large,
            color=font_color,
            italic=(col_name == "캐릭터명" and str(cell.value) == "내레이션"),
        )


def _set_col_widths(ws, cols: list[str]):
    widths = {
        "파일명":       40,
        "캐릭터명":     16,
        "감정":         22,
        "REC":          5,
        "대사":         55,
        "ALT":          45,
        "옵티컬(원문)": 45,
        "⏱ 검수":      18,
        "ADR/Wild":     10,
        "타임코드":     26,
    }
    for i, col in enumerate(cols, 1):
        ws.column_dimensions[get_column_letter(i)].width = widths.get(col, 15)


# ── 명방캐릭터 시트 (원본 그대로) ────────────────────────────────────────

def _add_arknights_sheet(wb: openpyxl.Workbook, item: dict):
    sname = item["sheet_name_out"][:31]
    ws = wb.create_sheet(title=sname)
    fpath = item.get("source_fpath")
    src_sheet = item.get("source_sheet")
    if not fpath or not src_sheet:
        return
    try:
        src_wb = openpyxl.load_workbook(fpath, data_only=True, read_only=True)
        src_ws = src_wb[src_sheet]
        for row in src_ws.iter_rows(values_only=True):
            ws.append([v for v in row])
        src_wb.close()
    except Exception as e:
        ws.append([f"시트 복사 오류: {e}"])


# ── 개괄 시트 ──────────────────────────────────────────────────────────────

def _add_summary_sheet(wb: openpyxl.Workbook, all_sheet_data: list[tuple]):
    ws = wb.create_sheet(title="개괄")

    # 캐릭터별 탭별 대사 수 집계
    char_tab_counts: dict[str, dict[str, int]] = {}
    tab_names = []
    for sname, rows, stype in all_sheet_data:
        if stype == "Type_명방캐릭터":
            continue
        tab_names.append(sname)
        for row in (rows or []):
            if row is None:
                continue
            char = row.get("캐릭터명", "").strip()
            if not char or char in ("캐릭터명",):
                continue
            if char not in char_tab_counts:
                char_tab_counts[char] = {}
            char_tab_counts[char][sname] = char_tab_counts[char].get(sname, 0) + 1

    # 헤더
    header = ["캐릭터명", "성우명"] + tab_names + ["합계"]
    ws.append(header)

    # 헤더 스타일
    hdr_fill = PatternFill("solid", fgColor=SUMMARY_BG_HDR)
    hdr_font = Font(name="맑은 고딕", size=10, bold=True, color=SUMMARY_FG_HDR)
    for cell in ws[1]:
        cell.fill      = hdr_fill
        cell.font      = hdr_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border    = _HDR_BORDER
    ws.row_dimensions[1].height = 28

    # 캐릭터별 행 (합계 내림차순)
    sorted_chars = sorted(
        char_tab_counts.items(),
        key=lambda x: sum(x[1].values()),
        reverse=True,
    )
    for r_idx, (char, tab_map) in enumerate(sorted_chars, 2):
        total = sum(tab_map.values())
        row_vals = [char, ""] + [tab_map.get(t, "") or "" for t in tab_names] + [total]
        ws.append(row_vals)
        for c_idx, val in enumerate(row_vals, 1):
            cell = ws.cell(row=r_idx, column=c_idx)
            col_name = header[c_idx - 1]
            cell.border = _CELL_BORDER
            if col_name == "캐릭터명":
                cell.font = Font(name="맑은 고딕", size=9, bold=True)
            elif col_name == "성우명":
                cell.fill = PatternFill("solid", fgColor=SUMMARY_CAST_BG)
                cell.font = Font(name="맑은 고딕", size=9)
            elif col_name == "합계":
                cell.fill = PatternFill("solid", fgColor=SUMMARY_TOTAL_BG)
                cell.font = Font(name="맑은 고딕", size=9, bold=True)
            else:
                cell.font = Font(name="맑은 고딕", size=9)
            cell.alignment = Alignment(horizontal="center", vertical="center")

    # 열 폭
    ws.column_dimensions["A"].width = 20
    ws.column_dimensions["B"].width = 14
    for i in range(3, len(header) + 1):
        ws.column_dimensions[get_column_letter(i)].width = max(12, len(header[i - 1]) * 1.5)

    ws.freeze_panes = "A2"


# ── 로그 시트 ──────────────────────────────────────────────────────────────

def _add_log_sheet(wb: openpyxl.Workbook, log_entries: list[dict]):
    ws = wb.create_sheet(title="로그")
    ws.append(["파일명", "시트명", "사유"])
    hdr_fill = PatternFill("solid", fgColor=HDR_BG)
    hdr_font = Font(name="맑은 고딕", size=10, bold=True, color=HDR_FG)
    for cell in ws[1]:
        cell.fill      = hdr_fill
        cell.font      = hdr_font
        cell.alignment = Alignment(horizontal="center")
        cell.border    = _HDR_BORDER
    for entry in log_entries:
        ws.append([entry.get("file", ""), entry.get("sheet", ""), entry.get("reason", "")])
    ws.column_dimensions["A"].width = 50
    ws.column_dimensions["B"].width = 30
    ws.column_dimensions["C"].width = 50


# ── 시트 순서 재배치 ─────────────────────────────────────────────────────

def _reorder_sheets(wb: openpyxl.Workbook):
    names = wb.sheetnames
    front = []
    back  = []
    for n in names:
        if n in ("개괄", "로그"):
            front.append(n)
        else:
            back.append(n)
    ordered = ["개괄"] + ([n for n in front if n != "개괄"]) + back
    # 로그는 개괄 바로 다음
    if "로그" in ordered:
        ordered.remove("로그")
        ordered.insert(1, "로그")
    for i, name in enumerate(ordered):
        if name in wb.sheetnames:
            wb.move_sheet(name, offset=wb.sheetnames.index(name) * -1 + i)
