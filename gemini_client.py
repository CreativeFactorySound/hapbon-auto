"""Gemini API 래퍼 — 시트 분류 + 컬럼 매핑 (google-genai 사용)."""

import json
import re
from pathlib import Path
from google import genai
from google.genai import types

import sys

# 시스템 프롬프트를 외부 파일에서 로드
# 규칙 수정 시 prompts/classify_system.txt 만 편집하면 됩니다
# PyInstaller exe로 실행 시 sys._MEIPASS 경로 사용
_BASE_DIR = Path(sys._MEIPASS) if getattr(sys, "frozen", False) else Path(__file__).parent
_PROMPT_FILE = _BASE_DIR / "prompts" / "classify_system.txt"
_CLASSIFY_SYSTEM_BASE = _PROMPT_FILE.read_text(encoding="utf-8")


class GeminiClient:
    def __init__(self, api_key: str, profile: dict | None = None):
        self._client = genai.Client(api_key=api_key)
        # 프로파일 기반 classify 시스템 프롬프트 조합
        extra = (profile or {}).get("classify_extra", "").strip()
        self._classify_system = _CLASSIFY_SYSTEM_BASE + ("\n" + extra if extra else "")

    def classify_sheet(
        self,
        filename: str,
        sheetname: str,
        rows_preview: list[list],
        optical_lang: str,
        max_retries: int = 6,
    ) -> dict:
        """시트 분류 + 컬럼 매핑 JSON 반환. 503/429 에러 시 자동 재시도."""
        import time

        preview_text = _format_preview(rows_preview)
        user_msg = (
            f"파일명: {filename}\n"
            f"시트명: {sheetname}\n"
            f"옵티컬 언어 설정: {optical_lang}\n\n"
            f"시트 데이터 (rowN: [열인덱스]값 형식):\n{preview_text}\n\n"
            "위 데이터를 분석해서 JSON을 출력하세요."
        )
        last_err = None
        for attempt in range(max_retries):
            try:
                response = self._client.models.generate_content(
                    model="gemini-2.5-flash",
                    contents=user_msg,
                    config=types.GenerateContentConfig(
                        system_instruction=self._classify_system,
                        response_mime_type="application/json",
                        temperature=0.0,
                    ),
                )
                return _parse_json_response(response.text)
            except Exception as e:
                last_err = e
                err_str = str(e)
                if "503" in err_str or "429" in err_str or "UNAVAILABLE" in err_str:
                    wait = 8 * (2 ** attempt)
                    print(f"\n  [재시도 {attempt+1}/{max_retries}] {wait}초 대기...", end=" ", flush=True)
                    time.sleep(wait)
                else:
                    raise
        raise last_err


def _format_preview(rows: list[list]) -> str:
    lines = []
    for i, row in enumerate(rows):
        cells = [f"[{j}]{str(v)[:35]}" for j, v in enumerate(row) if v is not None]
        if cells:
            lines.append(f"row{i}: " + "  ".join(cells))
    return "\n".join(lines)


def _parse_json_response(text: str) -> dict:
    text = text.strip()
    # 코드블럭 제거
    text = re.sub(r"```(?:json)?", "", text).strip()
    text = re.sub(r"```", "", text).strip()
    try:
        return json.loads(text)
    except json.JSONDecodeError:
        # JSON 블럭 찾기
        m = re.search(r"\{[\s\S]+\}", text)
        if m:
            return json.loads(m.group())
        raise ValueError(f"Gemini 응답에서 JSON 파싱 실패:\n{text[:500]}")
