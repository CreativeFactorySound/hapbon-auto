"""프로젝트 프로파일 로더 및 인터랙티브 선택.

각 프로파일은 profiles/<id>/profile.json + profiles/<id>/classify_system.txt 로 구성.
"""

import json
import sys
from pathlib import Path

def _get_base_dir() -> Path:
    """PyInstaller 패키징 여부에 따라 기준 디렉토리 반환."""
    if getattr(sys, "frozen", False):
        # exe 실행 시: sys._MEIPASS (번들 내부) 우선, 없으면 exe 옆 폴더
        meipass = Path(getattr(sys, "_MEIPASS", ""))
        if (meipass / "profiles").exists():
            return meipass
        return Path(sys.executable).parent
    return Path(__file__).parent

_PROFILES_DIR = _get_base_dir() / "profiles"

# 표시 순서
_PROFILE_ORDER = ["default", "reverse1999", "muhandae"]


def list_profiles() -> list[dict]:
    """사용 가능한 프로파일 목록을 순서대로 반환."""
    profiles = []
    seen: set[str] = set()

    for pid in _PROFILE_ORDER:
        p = _PROFILES_DIR / pid / "profile.json"
        if p.exists():
            try:
                data = json.loads(p.read_text(encoding="utf-8"))
                profiles.append(data)
                seen.add(pid)
            except Exception:
                pass

    # 순서에 없는 나머지 프로파일도 추가
    for sub in sorted(_PROFILES_DIR.iterdir()):
        if not sub.is_dir():
            continue
        pid = sub.name
        if pid in seen:
            continue
        p = sub / "profile.json"
        if p.exists():
            try:
                data = json.loads(p.read_text(encoding="utf-8"))
                profiles.append(data)
            except Exception:
                pass

    return profiles


def load_profile(profile_id: str) -> dict:
    """ID로 프로파일 로드. 없으면 default 반환."""
    p = _PROFILES_DIR / profile_id / "profile.json"
    if p.exists():
        try:
            return json.loads(p.read_text(encoding="utf-8"))
        except Exception:
            pass
    return load_profile("default")


def get_classify_system(profile: dict) -> str:
    """프로파일 전용 classify_system.txt 로드.
    프로파일 폴더에 없으면 default 폴더에서 폴백."""
    pid = profile.get("id", "default")

    # 1순위: 프로파일 전용
    profile_path = _PROFILES_DIR / pid / "classify_system.txt"
    if profile_path.exists():
        return profile_path.read_text(encoding="utf-8")

    # 2순위: default
    default_path = _PROFILES_DIR / "default" / "classify_system.txt"
    if default_path.exists():
        return default_path.read_text(encoding="utf-8")

    # 3순위: 구버전 경로 (하위 호환)
    legacy = Path(__file__).parent / "prompts" / "classify_system.txt"
    if legacy.exists():
        return legacy.read_text(encoding="utf-8")

    return ""


def select_profile_interactive() -> tuple[dict, str, str, str, str]:
    """
    인터랙티브 프로젝트 선택 메뉴.
    반환: (profile, project_name, round_str, record, optical)
    """
    profiles = list_profiles()

    print("\n" + "=" * 55)
    print("  합본 자동화 — 프로젝트 선택")
    print("=" * 55)
    for i, p in enumerate(profiles, 1):
        note = ""
        if p.get("_note"):
            note = " ⚠ (개발 예정)"
        print(f"  {i}. {p['display_name']}{note}")
    print("=" * 55)

    profile = None
    while profile is None:
        choice = input("선택 (번호): ").strip()
        if choice.isdigit():
            idx = int(choice) - 1
            if 0 <= idx < len(profiles):
                profile = profiles[idx]
            else:
                print("  잘못된 번호입니다.")
        else:
            print("  번호를 입력해주세요.")

    print(f"\n  프로파일: {profile['display_name']}")

    # 프로젝트명
    default_name = profile["display_name"].split(" (")[0]
    if profile["id"] == "default":
        name_input = input("프로젝트명 (필수): ").strip()
        project_name = name_input if name_input else "Unknown"
    else:
        name_input = input(f"프로젝트명 (Enter = {default_name!r}): ").strip()
        project_name = name_input if name_input else default_name

    # 차수
    round_input = input("차수 (예: 3.7차, Enter = 생략): ").strip()

    # 녹음 언어
    default_rec = profile["defaults"].get("record", "KR")
    rec_input = input(f"녹음 언어 KR/JP/EN (Enter = {default_rec}): ").strip().upper()
    record = rec_input if rec_input in ("KR", "JP", "EN") else default_rec

    # 옵티컬
    default_opt = profile["defaults"].get("optical", "EN")
    opt_input = input(f"옵티컬 언어 EN/CN/NONE (Enter = {default_opt}): ").strip().upper()
    optical = opt_input if opt_input in ("EN", "CN", "NONE") else default_opt

    print()
    return profile, project_name, round_input, record, optical


def build_classify_prompt(base_prompt: str, profile: dict) -> str:
    """레거시 호환용 — 현재는 get_classify_system()으로 대체됨."""
    extra = profile.get("classify_extra", "").strip()
    if extra:
        return base_prompt + "\n" + extra
    return base_prompt


def make_sheet_name_rv1999(fname: str, sname: str, cls: dict) -> str:
    """리버스1999 전용 탭명 생성.

    파일명 한국어 prefix에서 '캐릭터 (유형)' 형태 추출.
    전투 시트는 '전투'로 고정.
    """
    import re

    t = cls.get("type", "")
    char_name = cls.get("char_name_kr") or ""

    # 전투 타입 고정
    if t == "Type_전투":
        return "전투"

    stem = Path(fname).stem

    # 파일명에서 한국어 prefix 추출 (버전번호 앞까지)
    m = re.match(r"^([\w가-힣\s·]+?)\s*(?:\d+\.\d|20\d{6})", stem)
    if m:
        kr_prefix = m.group(1).strip()
    else:
        kr_prefix = re.sub(r"[一-鿿㐀-䶿]+.*$", "", stem).strip()
        kr_prefix = re.sub(r"[_\-]", " ", kr_prefix).strip()

    if not kr_prefix:
        kr_prefix = stem[:20]

    # TYPE 키워드 정규화 — 전체 순서대로 모두 적용
    TYPE_KEYWORDS = [
        (r"\b캐릭터\s*심층\s*분석\b", "(캐릭터 분석)"),
        (r"\b캐릭터\s*PV\b",           "(캐릭터 PV)"),
        (r"\b스킬\s*PV\b",             "(스킬 PV)"),
        (r"^SP\s+",                    ""),           # SP 접두어 제거 (문두)
        (r"\b음성\b",                  "(음성)"),
        (r"\b스킨\b",                  "(스킨)"),
        (r"(?<!\()\bPV\b(?!\))",       "(PV)"),       # 이미 괄호 안에 있으면 스킵
    ]

    for pattern, repl in TYPE_KEYWORDS:
        kr_prefix = re.sub(pattern, repl, kr_prefix).strip()

    # 잉여 공백 정리
    result = re.sub(r"\s+", " ", kr_prefix).strip()

    # 캐릭터명 명시적 접두어 (짧은음성 등)
    if char_name and char_name not in result:
        result = f"{char_name}_{result}"

    # Excel 시트명 금지 문자 제거 및 길이 제한
    result = re.sub(r"[:/\?*\[\]\\]", "", result)
    return result[:31] or "시트"
