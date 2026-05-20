# 업데이트 이력

---

## [1.3.0] - 2026-05-20

### 다국어(JP/EN) 지원 전면 보완 — 25개 이슈 일괄 수정

#### processor.py
- **`_is_peak` 다국어화**: KR 전용이던 피크 단어 목록에 JP(叫び, 絶叫, 悲鳴 등), EN(scream, shriek 등) 추가. `record` 파라미터 추가
- **`_timing_flag` / `_est_sec` 다국어화**: JP는 `_est_sec_jp`(히라가나·가타카나·한자 모라 계산), EN은 `_est_sec_en`(단어 수 계산)으로 분기. 기존엔 항상 KR 계산(0초)
- **`_extract_*` 전체 `record` 주입**: 모든 추출 함수가 `cls["_record"]`를 읽어 피크·타이밍 등 처리
- **`[KR 번역 없음]` 레이블 제거**: `[번역 없음]`으로 수정 (JP/EN 녹음에서 오해 소지 제거)
- **`_NARRATION_STEPS` 확장**: 旁白/내레이션에 ナレーション/Narration/Narrator 추가 → 내레이션 정규화 완성
- **`_HEADER_CHAR_VALUES` / `_HEADER_DIAL_VALUES` / `_HEADER_FILE_VALUES` 전역 세트 신설**: `_extract_battle`, `_extract_chara`에서 JP/EN 헤더 반복 행도 스킵
- **`_extract_infinite` 녹음여부 컬럼 탐색 확장**: 收录/Record에 수록/録音有無/Include 추가
- **`_has_korean` 데드 코드 제거**

#### main.py
- **EN 시트 필터링 임계값 수정**: `[A-Za-z]{5,}` → `{2,}` (2글자 이상 영단어 존재 시 통과). 기존엔 4글자 이하 영단어만 있는 EN 시트가 전부 스킵됨
- **캐시 키에 `optical` + `record` 추가**: 같은 파일을 다른 언어 설정으로 재실행할 때 잘못된 캐시 재사용 방지
- **`_make_sheet_name` `strip_cn` 수정**: `぀-ヿ`(히라가나+가타카나) 범위 제거 → JP 시트명이 공백으로 치환되던 버그 수정
- **`_make_sheet_name` 유효성 검사 확장**: `[가-힣a-zA-Z]`에 `ぁ-ヿ一-鿿` 추가 → JP 시트명이 generic 라벨로 교체되던 버그 수정
- **`_record` 주입**: 분류 완료 후 모든 `cls` 딕셔너리에 `_record` 키 삽입 (processor 함수 전달)
- **중복 `[완료]` 출력 제거**

#### assembler.py
- **개괄 시트 캐릭터 헤더 제외 목록 확장**: 캐릭터명/캐릭터/配音对象/角色/キャラ/Character 추가

---

## [1.2.0] - 2026-05-20

### 변경
- **Type_짧은음성 분류 정확도 개선** (`prompts/classify_system.txt`, `hapbon/functions/classification_guide.md`)
  - 기존: "캐릭터명 병합셀 + 대사 1~3어절" 조건 → 파일명 열이 있으면 Type_메인으로 오분류
  - 변경: 미리보기 데이터에서 캐릭터명 열이 첫 데이터 행에만 있고 이후 여러 행이 비어 있으면 (병합셀 패턴) → 파일명 열/캐릭터 수/대사 길이와 무관하게 Type_짧은음성 강제 적용
  - Type_메인 조건 추가: "캐릭터명 열이 매 대사 행 채워진 구조"
  - 헤더 키워드에 `キャラ`, `台詞` 추가 (일본어 대본 지원)

- **Type_짧은음성 병합셀 전면 확장** (`processor.py`)
  - 기존: 캐릭터명 열(キャラ)의 병합셀만 확장
  - 변경: 모든 열(감정/詳細, 대사/台詞, 파일명 등)의 병합셀을 통합 맵으로 관리 → 感情·詳細 데이터가 모든 행에 올바르게 채워짐

- **JP 녹음 대사 컬럼 매핑 명확화** (`prompts/classify_system.txt`)
  - `col_dialogue_kr`: "녹음 받는 언어의 주 대사 열 — JP이면 台詞, 절대 col_optical/col_dialogue_cn에 넣지 말 것" 명시
  - `col_dialogue_cn`: JP 스크립트에 중국어 없으면 -1 명시
  - `col_optical`: 해당 언어 열 없으면 -1, 녹음 대상 언어 대사 넣지 말 것 명시

- **Type_짧은음성 이미지 보존** (`processor.py`, `assembler.py`, `main.py`, `requirements.txt`)
  - 소스 xlsx에 포함된 캐릭터 이미지(twoCellAnchor 방식)를 zipfile로 직접 추출
  - 이미지를 캐릭터 섹션에 매핑 후 합본 시트 오른쪽 열에 자동 배치
  - Pillow 의존성 추가 (`requirements.txt`)

---

## [1.1.0] - 2026-05-20

### 변경
- **녹음 언어 기반 시트 필터링** (`main.py`, `prompts/classify_system.txt`)
  - 기존: 항상 한국어(KR) 유무로 시트 스킵 여부 판단
  - 변경: `--record` 파라미터에 따라 판단 기준 언어 자동 전환
    - `KR` → 한국어(가-힣) 존재 여부
    - `JP` → 일본어 히라가나/가타카나(ぁ-ん, ァ-ン) 존재 여부
    - `EN` → 영어 단어(5자 이상 연속 알파벳) 존재 여부
- `hapbon` MD 레포 `classification_guide.md` 동일 내용 반영

---

## [1.0.0] - 2026-05-19 (초기 완성)

### 핵심 구조 확립
- Gemini 2.5 Flash API로 시트 자동 분류 + 컬럼 매핑
- 7종 시트 타입 처리 (메인/PV/전투/캐릭터음성/짧은음성/무한대/명방캐릭터)
- tkinter GUI 실행기 추가 (`gui.py`)
- PyInstaller exe 빌드 지원 (`dist/합본자동화.exe`)
- GitHub 레포 연결: https://github.com/CreativeFactorySound/hapbon-auto

### 분류·추출
- Gemini 시스템 프롬프트를 외부 파일로 분리 (`prompts/classify_system.txt`)
  - 규칙 수정 시 Python 코드 건드릴 필요 없음
- 분류 캐시 구현 (`.hapbon_cls_cache.json`) — 재실행 시 API 재호출 방지
- 503/429 에러 자동 재시도 (최대 6회, 지수 백오프)
- Type_전투: boss + 일반 전투 시트를 1개 탭으로 합산 (boss 먼저)
- Type_짧은음성: 병합셀 확장 처리
- Type_PV: 단일 캐릭터 비율 30% 이상이면 캐릭터명 접두사 자동 부여

### 서식
- 헤더: 남색 배경, 흰 Bold
- 캐릭터명·대사·ALT: **14pt Bold** 강조
- 모든 셀 테두리 적용
- ALT 열: 항상 맨 끝에 포함, 데이터 없으면 자동 숨김
- 타이밍 초과 행: 주황 배경 + 경고 메시지
- 피크 라인: 대사 빨간 텍스트
- 개괄 시트: 캐릭터×탭 대사 수 집계

### 기타
- Windows CP949 터미널 한/중 문자 출력 오류 수정
- 시트명 한글화: 중국어·날짜 숫자 제거, Sheet1 → 파일명 기반 자동 생성
- 시트명 중복 시 `_2`, `_3` 접미사 자동 처리
- GUI: 차수(round) 입력값 저장 제외 (매번 직접 입력)

---

## 업데이트 방법

변경사항 생기면 아래 형식으로 추가:

```
## [버전] - YYYY-MM-DD

### 변경 카테고리
- 내용
```

> 분류 규칙 변경은 `prompts/classify_system.txt` + 이 파일에 기록  
> 코드 변경은 해당 파일명과 함께 기록
