"""타입별 시트 데이터 추출."""

import re
import openpyxl


def process_sheet(ws, cls: dict) -> list[dict]:
    """분류 결과(cls)에 따라 시트에서 행 데이터를 추출해 반환."""
    t = cls["type"]
    if t == "Type_메인":
        return _extract_main(ws, cls)
    elif t == "Type_PV":
        return _extract_pv(ws, cls)
    elif t == "Type_전투":
        return _extract_battle(ws, cls)
    elif t == "Type_캐릭터음성":
        return _extract_chara(ws, cls)
    elif t == "Type_짧은음성":
        return _extract_short(ws, cls)
    elif t == "Type_무한대":
        return _extract_infinite(ws, cls)
    else:
        return []


# ── 공통 헬퍼 ────────────────────────────────────────────────────────────────

def _cell(row_vals: list, idx: int):
    """안전하게 컬럼 값 반환. idx=-1이면 None."""
    if idx < 0 or idx >= len(row_vals):
        return None
    v = row_vals[idx]
    return str(v).strip() if v is not None else None


def _has_korean(text) -> bool:
    if not text:
        return False
    return any("가" <= c <= "힣" for c in str(text))


def _est_sec_kr(text: str) -> float:
    """한국어 글자 수로 예상 녹음 시간(초) 추정."""
    if not text:
        return 0.0
    return sum(1 for c in text if "가" <= c <= "힣") / 6.5


def _est_sec_en(text: str) -> float:
    if not text:
        return 0.0
    return len(str(text).split()) / 2.5


def _parse_duration(s) -> float | None:
    """'≤3s', '2.5초', '3-4s' 등에서 숫자(초) 파싱."""
    if s is None:
        return None
    s = str(s).replace("≤", "").replace("≦", "").replace("约", "").strip()
    # 범위 표기 (3-4s) → 최댓값 사용
    m = re.search(r"([\d.]+)\s*[-~]\s*([\d.]+)", s)
    if m:
        return float(m.group(2))
    m = re.search(r"([\d.]+)\s*[sS초]", s)
    if m:
        return float(m.group(1))
    return None


def _is_peak(text: str) -> bool:
    if not text:
        return False
    peak_words = ["비명", "기합", "으아", "절규", "소리치", "폭발", "외치", "포효",
                  "분노", "열받", "살려", "싫어", "제발", "도망", "그만", "아악", "으윽"]
    return "!!" in text or any(w in text for w in peak_words)


def _timing_flag(dialogue_kr: str, duration_raw) -> str:
    dur = _parse_duration(duration_raw)
    if dur is None:
        return ""
    est = _est_sec_kr(dialogue_kr or "")
    if est > dur * 1.1:
        return f"⚠ 예상{est:.1f}s > 허용{dur}s"
    return ""


def _ws_rows(ws, header_row: int) -> list[list]:
    """워크시트에서 header_row 이후 데이터 행들을 반환."""
    rows = []
    for i, row in enumerate(ws.iter_rows(values_only=True)):
        if i > header_row:
            rows.append(list(row))
    return rows


# ── Type_메인 ──────────────────────────────────────────────────────────────

def _extract_main(ws, cls: dict) -> list[dict]:
    header_row = cls.get("header_row", 0)
    cf = cls.get("col_filename", -1)
    ckr = cls.get("col_char_kr", -1)
    ckn = cls.get("col_char_cn", -1)
    ekr = cls.get("col_emotion_kr", -1)
    ecn = cls.get("col_emotion_cn", -1)
    dkr = cls.get("col_dialogue_kr", -1)
    dcn = cls.get("col_dialogue_cn", -1)
    opt = cls.get("col_optical", -1)
    alt = cls.get("col_alt", -1)
    step_col = cls.get("col_step_type", -1)
    skip_steps = set(cls.get("skip_step_values", []))
    narration_steps = {"旁白", "내레이션"}

    result = []
    for row_vals in _ws_rows(ws, header_row):
        filename = _cell(row_vals, cf)
        char_kr = _cell(row_vals, ckr)
        char_cn = _cell(row_vals, ckn)
        dialogue_kr = _cell(row_vals, dkr)
        dialogue_cn = _cell(row_vals, dcn)
        emotion_kr = _cell(row_vals, ekr)
        emotion_cn = _cell(row_vals, ecn)
        optical = _cell(row_vals, opt)
        alt_val = _cell(row_vals, alt)
        step = _cell(row_vals, step_col)

        # 스킵: 씬 타입이 skip_steps에 해당
        if step and step in skip_steps:
            continue

        # 스킵: 파일명도 없고 대사도 없고 캐릭터도 없는 완전 빈 행
        if not filename and not dialogue_kr and not dialogue_cn and not char_kr and not char_cn:
            continue

        # 旁白 → 내레이션
        if step and step in narration_steps:
            char_kr = "내레이션"

        # KR 없고 CN만 있는 경우
        if not dialogue_kr and dialogue_cn:
            emotion_kr = (emotion_kr or "") + ("[KR 번역 없음]" if emotion_kr else "[KR 번역 없음]")

        # 감정: KR 우선, 없으면 CN
        emotion = emotion_kr or emotion_cn or ""

        result.append({
            "파일명": filename or "",
            "캐릭터명": char_kr or char_cn or "",
            "감정": emotion,
            "REC": "",
            "대사": dialogue_kr or "",
            "ALT": alt_val or "",
            "옵티컬(원문)": optical or dialogue_cn or "",
            "⏱ 검수": "",
            "_peak": _is_peak(dialogue_kr or ""),
        })
    return result


# ── Type_PV ────────────────────────────────────────────────────────────────

def _extract_pv(ws, cls: dict) -> list[dict]:
    header_row = cls.get("header_row", 0)
    ckr = cls.get("col_char_kr", -1)
    ckn = cls.get("col_char_cn", -1)
    ekr = cls.get("col_emotion_kr", -1)
    ecn = cls.get("col_emotion_cn", -1)
    dkr = cls.get("col_dialogue_kr", -1)
    dcn = cls.get("col_dialogue_cn", -1)
    opt = cls.get("col_optical", -1)
    alt = cls.get("col_alt", -1)
    dur = cls.get("col_duration", -1)

    result = []
    for row_vals in _ws_rows(ws, header_row):
        char_kr = _cell(row_vals, ckr)
        char_cn = _cell(row_vals, ckn)
        dialogue_kr = _cell(row_vals, dkr)
        dialogue_cn = _cell(row_vals, dcn)
        emotion_kr = _cell(row_vals, ekr)
        emotion_cn = _cell(row_vals, ecn)
        optical = _cell(row_vals, opt)
        alt_val = _cell(row_vals, alt)
        duration_raw = _cell(row_vals, dur)

        # 빈 행 스킵
        if not char_kr and not char_cn and not dialogue_kr and not dialogue_cn:
            continue

        # 旁白
        if not char_kr and not char_cn and (dialogue_kr or dialogue_cn):
            char_kr = "내레이션"

        emotion = emotion_kr or emotion_cn or ""
        timing = _timing_flag(dialogue_kr or "", duration_raw) if dur >= 0 else ""

        result.append({
            "파일명": "",
            "캐릭터명": char_kr or char_cn or "",
            "감정": emotion,
            "REC": "",
            "대사": dialogue_kr or "",
            "ALT": alt_val or "",
            "옵티컬(원문)": optical or dialogue_cn or "",
            "⏱ 검수": timing,
            "_peak": _is_peak(dialogue_kr or ""),
            "_timing_over": bool(timing),
        })
    return result


# ── Type_전투 ──────────────────────────────────────────────────────────────

def _extract_battle(ws, cls: dict) -> list[dict]:
    header_row = cls.get("header_row", 0)
    cf = cls.get("col_filename", -1)
    ckr = cls.get("col_char_kr", -1)
    ckn = cls.get("col_char_cn", -1)
    ekr = cls.get("col_emotion_kr", -1)
    ecn = cls.get("col_emotion_cn", -1)
    dkr = cls.get("col_dialogue_kr", -1)
    dcn = cls.get("col_dialogue_cn", -1)
    opt = cls.get("col_optical", -1)
    alt = cls.get("col_alt", -1)
    process_col = cls.get("col_step_type", -1)  # 流程节点

    result = []
    for row_vals in _ws_rows(ws, header_row):
        filename = _cell(row_vals, cf)
        char_kr = _cell(row_vals, ckr)
        char_cn = _cell(row_vals, ckn)
        dialogue_kr = _cell(row_vals, dkr)
        dialogue_cn = _cell(row_vals, dcn)
        emotion_kr = _cell(row_vals, ekr)
        emotion_cn = _cell(row_vals, ecn)
        optical = _cell(row_vals, opt)
        alt_val = _cell(row_vals, alt)
        process_val = _cell(row_vals, process_col)

        # 파트 제목 행 스킵: 流程节点에 전투 단계가 있고 캐릭터/대사가 없는 경우
        if process_val and not char_kr and not char_cn and not dialogue_kr and not dialogue_cn:
            continue

        # 헤더 반복 행 스킵
        if char_kr in ("캐릭터", "配音对象") or dialogue_kr in ("라인", "对话文本"):
            continue

        # 완전 빈 행 스킵
        if not filename and not char_kr and not char_cn and not dialogue_kr and not dialogue_cn:
            continue

        emotion = emotion_kr or emotion_cn or ""
        result.append({
            "파일명": filename or "",
            "캐릭터명": char_kr or char_cn or "",
            "감정": emotion,
            "REC": "",
            "대사": dialogue_kr or "",
            "ALT": alt_val or "",
            "옵티컬(원문)": optical or dialogue_cn or "",
            "⏱ 검수": "",
            "_peak": _is_peak(dialogue_kr or ""),
        })
    return result


# ── Type_캐릭터음성 ────────────────────────────────────────────────────────

def _extract_chara(ws, cls: dict) -> list[dict]:
    header_row = cls.get("header_row", 0)
    cf = cls.get("col_filename", -1)
    char_name = cls.get("char_name_kr") or ""
    ekr = cls.get("col_emotion_kr", -1)
    ecn = cls.get("col_emotion_cn", -1)
    dkr = cls.get("col_dialogue_kr", -1)
    dcn = cls.get("col_dialogue_cn", -1)
    opt = cls.get("col_optical", -1)
    alt = cls.get("col_alt", -1)
    func_col = cls.get("col_functional", -1)

    result = []
    for row_vals in _ws_rows(ws, header_row):
        filename = _cell(row_vals, cf)
        dialogue_kr = _cell(row_vals, dkr)
        dialogue_cn = _cell(row_vals, dcn)
        emotion_kr = _cell(row_vals, ekr)
        emotion_cn = _cell(row_vals, ecn)
        optical = _cell(row_vals, opt)
        alt_val = _cell(row_vals, alt)
        functional = _cell(row_vals, func_col)

        # 빈 행 / 헤더 반복 스킵
        if not filename and not dialogue_kr and not dialogue_cn:
            continue
        if dialogue_kr in ("라인", "KR") or filename in ("语音命名", "台詞番号"):
            continue

        # KR 번역 없는 경우
        if not dialogue_kr and dialogue_cn:
            emotion_kr = (emotion_kr or "") + "[KR 번역 없음]"

        emotion = emotion_kr or emotion_cn or ""
        # 功能을 감정 앞에 참고 정보로 붙임
        if functional:
            emotion = f"[{functional}] {emotion}".strip()

        result.append({
            "파일명": filename or "",
            "캐릭터명": char_name,
            "감정": emotion,
            "REC": "",
            "대사": dialogue_kr or "",
            "ALT": alt_val or "",
            "옵티컬(원문)": optical or dialogue_cn or "",
            "⏱ 검수": "",
            "_peak": _is_peak(dialogue_kr or ""),
        })
    return result


# ── Type_짧은음성 ──────────────────────────────────────────────────────────

def _extract_short(ws, cls: dict) -> list[dict]:
    """병합셀 처리가 필요하므로 ws는 read_only=False로 열어야 함."""
    header_row = cls.get("header_row", 0)
    cf = cls.get("col_filename", -1)
    ckr = cls.get("col_char_kr", -1)
    ckn = cls.get("col_char_cn", -1)
    ekr = cls.get("col_emotion_kr", -1)
    ecn = cls.get("col_emotion_cn", -1)
    dkr = cls.get("col_dialogue_kr", -1)
    dcn = cls.get("col_dialogue_cn", -1)
    opt = cls.get("col_optical", -1)

    # 병합셀 캐릭터명 매핑 구성
    merged_char = {}  # row_idx → (char_cn, char_kr)
    try:
        for merge_range in ws.merged_cells.ranges:
            r_min, r_max = merge_range.min_row, merge_range.max_row
            c_min = merge_range.min_col - 1  # 0-based
            # 캐릭터명 열(ckn 또는 ckr)과 겹치는지 확인
            if c_min in (ckn, ckr):
                val = ws.cell(merge_range.min_row, merge_range.min_col).value
                for r in range(r_min, r_max + 1):
                    row_key = r - 1  # 0-based
                    if row_key not in merged_char:
                        merged_char[row_key] = {}
                    if c_min == ckn:
                        merged_char[row_key]["cn"] = str(val).strip() if val else ""
                    elif c_min == ckr:
                        merged_char[row_key]["kr"] = str(val).strip() if val else ""
    except Exception:
        pass  # read_only 모드이면 merged_cells 접근 불가 → 그냥 진행

    result = []
    all_rows = list(ws.iter_rows(values_only=True))
    for i, row_vals in enumerate(all_rows):
        if i <= header_row:
            continue
        row_vals = list(row_vals)

        filename = _cell(row_vals, cf)
        dialogue_kr = _cell(row_vals, dkr)
        dialogue_cn = _cell(row_vals, dcn)

        # 병합셀에서 캐릭터명 가져오기
        mc = merged_char.get(i, {})
        char_kr = mc.get("kr") or _cell(row_vals, ckr)
        char_cn = mc.get("cn") or _cell(row_vals, ckn)

        emotion_kr = _cell(row_vals, ekr)
        emotion_cn = _cell(row_vals, ecn)
        optical = _cell(row_vals, opt)

        # KR 대사 없는 행 스킵
        if not dialogue_kr and not dialogue_cn:
            continue
        if not filename and not dialogue_kr:
            continue

        if not dialogue_kr and dialogue_cn:
            emotion_kr = (emotion_kr or "") + "[KR 번역 없음]"

        emotion = emotion_kr or emotion_cn or ""
        result.append({
            "파일명": filename or "",
            "캐릭터명": char_kr or char_cn or "",
            "감정": emotion,
            "REC": "",
            "대사": dialogue_kr or "",
            "ALT": "",
            "옵티컬(원문)": optical or dialogue_cn or "",
            "⏱ 검수": "",
            "_peak": _is_peak(dialogue_kr or ""),
        })
    return result


# ── Type_무한대 ────────────────────────────────────────────────────────────

def _extract_infinite(ws, cls: dict) -> list[dict]:
    header_row = cls.get("header_row", 0)
    cf = cls.get("col_filename", -1)
    ckr = cls.get("col_char_kr", -1)
    ckn = cls.get("col_char_cn", -1)
    dkr = cls.get("col_dialogue_kr", -1)
    dcn = cls.get("col_dialogue_cn", -1)
    opt = cls.get("col_optical", -1)
    adr_col = cls.get("col_adr_wild", -1)
    tc_col = cls.get("col_timecode", -1)
    rec_col = -1  # 녹음여부 컬럼은 notes에서 추정

    # 녹음여부 컬럼 찾기 (Yes/No 패턴)
    if header_row >= 0:
        header_vals = []
        for i, row in enumerate(ws.iter_rows(values_only=True)):
            if i == header_row:
                header_vals = list(row)
                break
        for j, v in enumerate(header_vals):
            if v and ("收录" in str(v) or "Record" in str(v)):
                rec_col = j
                break

    result = []
    for row_vals in _ws_rows(ws, header_row):
        # 녹음여부 필터
        if rec_col >= 0:
            rec_val = _cell(row_vals, rec_col)
            if rec_val and rec_val.lower() in ("no", "否", "n"):
                continue

        filename = _cell(row_vals, cf)
        char_kr = _cell(row_vals, ckr)
        char_cn = _cell(row_vals, ckn)
        dialogue_kr = _cell(row_vals, dkr)
        dialogue_cn = _cell(row_vals, dcn)
        optical = _cell(row_vals, opt)
        adr_wild = _cell(row_vals, adr_col)
        timecode = _cell(row_vals, tc_col)

        if not dialogue_kr and not dialogue_cn:
            continue

        # 타임코드에서 duration 계산
        timing = ""
        if timecode and "-->" in str(timecode):
            dur = _parse_timecode_duration(str(timecode))
            if dur:
                timing = _timing_flag(dialogue_kr or "", f"{dur}s")

        result.append({
            "파일명": filename or "",
            "캐릭터명": char_kr or char_cn or "",
            "감정": "",
            "ADR/Wild": adr_wild or "",
            "타임코드": timecode or "",
            "REC": "",
            "대사": dialogue_kr or "",
            "ALT": "",
            "옵티컬(원문)": optical or dialogue_cn or "",
            "⏱ 검수": timing,
            "_peak": _is_peak(dialogue_kr or ""),
            "_timing_over": bool(timing),
        })
    return result


def _parse_timecode_duration(tc: str) -> float | None:
    """'00:00:05.700 --> 00:00:08.140' 형식에서 duration(초) 계산."""
    parts = tc.split("-->")
    if len(parts) != 2:
        return None
    try:
        def to_sec(s):
            s = s.strip().replace(",", ".")
            h, m, rest = s.split(":")
            return int(h) * 3600 + int(m) * 60 + float(rest)
        return to_sec(parts[1]) - to_sec(parts[0])
    except Exception:
        return None
