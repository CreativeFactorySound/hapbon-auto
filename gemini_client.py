"""Gemini API 래퍼 — 시트 분류 + 컬럼 매핑 (google-genai 사용)."""

import json
import re
from pathlib import Path
from google import genai
from google.genai import types


class GeminiClient:
    def __init__(self, api_key: str, profile: dict | None = None):
        self._client = genai.Client(api_key=api_key)
        # 프로파일별 classify_system.txt 로드 (profiles/<id>/classify_system.txt)
        from profile import get_classify_system
        self._classify_system = get_classify_system(profile or {"id": "default"})

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
            if attempt > 0:
                wait = 8 * (2 ** (attempt - 1))
                print(f"\n  [재시도 {attempt}/{max_retries}] {wait}초 대기...", end=" ", flush=True)
                time.sleep(wait)
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
                result = _parse_json_response(response.text)
                if attempt > 0:
                    time.sleep(2)   # 재시도 성공 후 다음 요청까지 여유
                else:
                    time.sleep(1)   # 정상 성공 후에도 1초 간격 유지
                return result
            except Exception as e:
                last_err = e
                err_str = str(e)
                if "503" in err_str or "429" in err_str or "UNAVAILABLE" in err_str:
                    continue
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
