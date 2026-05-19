# 합본 자동화 도구

게임 더빙 대본 엑셀 파일들을 표준 합본 형식으로 자동 변환하는 도구.  
Gemini 2.5 Flash API로 시트 타입을 분류하고, Python이 데이터를 추출·조립한다.

---

## 구조

```
hapbon_auto/
  gui.py                  # GUI 실행기 (tkinter) — 더블클릭 실행용
  main.py                 # CLI 진입점
  gemini_client.py        # Gemini API 래퍼 (시트 분류 + 컬럼 매핑)
  processor.py            # 타입별 데이터 추출
  assembler.py            # 합본 xlsx 조립 및 서식 적용
  requirements.txt        # 의존성
  prompts/
    classify_system.txt   # Gemini 시스템 프롬프트 (분류 규칙) ← 규칙 수정 시 여기만
```

---

## 처리 흐름

```
xlsx 파일들
    ↓
Gemini API  →  시트 타입 분류 + 컬럼 위치 파악
    ↓              (classify_system.txt 규칙 기반)
processor.py  →  타입별 데이터 추출
    ↓
assembler.py  →  합본 xlsx 조립 + 서식 적용
    ↓
합본.xlsx
```

### 지원 시트 타입

| 타입 | 설명 |
|------|------|
| Type_메인 | 스토리 대본 (여러 캐릭터, 파일명 열 있음) |
| Type_PV | PV/홍보 영상 대본 (时长 열 있음) |
| Type_전투 | 전투 흐름 대본 (流程节点 열 있음) |
| Type_캐릭터음성 | 캐릭터 1명 전용 음성 (ID+功能+终选语音 구조) |
| Type_짧은음성 | 병합셀 구조의 짧은 음성 대본 |
| Type_무한대 | 타임코드 포함 무한대 대본 |
| Type_명방캐릭터 | 명일방주 캐릭터 시트 (원본 그대로 복사) |

---

## 설치

```bash
python -m pip install -r requirements.txt
```

### 의존성
- `google-genai` — Gemini API
- `openpyxl` — xlsx 읽기/쓰기

---

## 실행

### GUI (권장)
```bash
python gui.py
```
더블클릭 또는 위 명령으로 실행. 입력값은 자동 저장됨 (차수 제외).

### CLI
```bash
python -X utf8 main.py \
  --source "원본폴더경로" \
  --output "합본출력.xlsx" \
  --project "프로젝트명" \
  --round "3.3차" \
  --optical EN \
  --record KR
```

| 옵션 | 필수 | 설명 |
|------|------|------|
| `--source` | ✅ | 원본 xlsx 폴더 경로 |
| `--output` | ✅ | 출력 합본 경로 (.xlsx) |
| `--project` | ✅ | 프로젝트명 |
| `--round` | | 차수 (예: 3.3차) |
| `--optical` | | 옵티컬 언어 EN/CN/NONE (기본값: EN) |
| `--record` | | 녹음 언어 KR/EN/JP (기본값: KR) |
| `--api-key` | | Gemini API 키 (환경변수 우선) |

### API 키 설정 (한 번만)
```powershell
[System.Environment]::SetEnvironmentVariable("GEMINI_API_KEY", "키값", "User")
```
이후 `--api-key` 생략 가능.

---

## exe 배포

```bash
python -m PyInstaller --onefile --windowed \
  --add-data "prompts/classify_system.txt;prompts" \
  --collect-data openpyxl \
  --name "합본자동화" gui.py
```
`dist/합본자동화.exe` 생성. Python 설치 없이 배포 가능.

---

## 분류 규칙 수정

`prompts/classify_system.txt` 파일만 편집하면 됨. Python 코드 수정 불필요.

분류 규칙 외 변경사항 (추출 방식, 출력 포맷 등)은 `processor.py` / `assembler.py` 수정 필요.  
→ **수정이 필요하면 MD 문서 들고 와서 요청.**

---

## 캐시

| 파일 | 위치 | 역할 |
|------|------|------|
| `.hapbon_cls_cache.json` | 출력 폴더 | Gemini API 재호출 방지 (같은 파일 재실행 시 비용 절약) |
| `.hapbon_gui_config.json` | 프로젝트 루트 | GUI 마지막 입력값 저장 (차수 제외) |

두 파일 모두 `.gitignore`에 포함 — git에 올라가지 않음.

---

## 출력 합본 서식

- **헤더**: 남색 배경 (#1F3864), 흰 글씨, Bold
- **캐릭터명 열**: 초록 배경 (#92D050), **14pt Bold**
- **대사/ALT 열**: 노란 배경 (#FFFDE7), **14pt Bold**
- **감정 열**: 연노랑 (#FFE599)
- **REC 열**: 연분홍 (#EA9999)
- **ALT 열**: 항상 맨 끝에 포함, 데이터 없으면 숨김 처리
- **타이밍 초과 행**: 주황 배경 (#FFE5CC) + `⚠ 예상Xs > 허용Xs`
- **피크 라인**: 대사 텍스트 빨간색 (!! 또는 비명/기합 등 키워드 포함 시)
- **개괄 시트**: 캐릭터×탭 대사 수 집계 매트릭스
- **모든 셀 테두리** 적용

---

## 업데이트 이력

자세한 내용은 [CHANGELOG.md](CHANGELOG.md) 참고.
