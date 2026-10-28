# office-automation-kr 구현계획

기준 설계: `docs/superpowers/specs/2026-10-28-office-automation-kr-design.md` (2026-10-28 승인)
작성일 2026-10-28 · 실행 환경: Windows PC, Claude Code(Code 탭), uv

## 실행 원칙

- 한 작업 = 한 커밋. 작업 순서대로 진행하고, 각 작업의 「완료 기준」이 모두 참일 때만 다음으로.
- 코드 작업은 테스트를 먼저 쓰고(실패 확인) → 구현 → 통과 → 커밋.
- 저장소 위치: 동기화 폴더 밖 `C:\dev\office-automation-kr` (git lock·동기화 충돌 회피). 완성 후 GitHub 공개 저장소로 push.
- 회사명·고객명·실제 양식은 어디에도 쓰지 않는다. 샘플은 전부 가짜(ACME 등).
- 모든 문서·주석·메시지는 한국어. 코드 식별자는 영어, 사용자에게 보이는 출력은 한국어.
- 금액은 달러 숫자 그대로(설계 5절). 원화 풀 콤마 규칙 적용 대상 아님.

## 0단계 — 저장소 뼈대

### 작업 0.1 저장소 초기화
- `C:\dev\office-automation-kr` 생성, `git init`, `uv init --no-workspace` 로 `pyproject.toml`.
- 의존성: `uv add pdfplumber openpyxl pyyaml` · dev: `uv add --dev pytest reportlab` · 선택: `pywin32`는 `[project.optional-dependencies] outlook`.
- `.gitignore`: `작업폴더.txt`, `.venv/`, `__pycache__/`, `*.xlsx`(samples 제외 없음 — 장부는 저장소에 없음), `.pytest_cache/`.
- 빈 폴더 구조: `app/ rules/ samples/ tests/ docs/superpowers/{specs,plans}`. 설계·계획 문서 복사.
- 완료 기준: `uv run python -c "import pdfplumber, openpyxl, yaml"` 성공. 첫 커밋 "저장소 뼈대".

### 작업 0.2 CLAUDE.md
- 설계 6절 다섯 항목을 그대로. 추가: "작업 폴더는 `작업폴더.txt`가 가리키는 곳. 저장소 안에 PDF·xlsx 를 만들지 않는다."
- 완료 기준: 파일 존재, 커밋.

## 1단계 — 가짜 샘플과 규칙 (테스트의 토대)

> 2026-10-28 추가: 샘플 세트가 이미 `samples/` 에 들어 있다 (`samples/README_샘플설명.md` 를 먼저 읽는다).
> 작업 1.1·1.2 는 "새로 만들기"가 아니라 **"들어 있는 파일을 검증하고 그대로 쓰기"** 다.
> - PDF 6종(Invoice 2·DM 2·서신 1·스캔 1)은 `samples/03_Invoice_DM_PDF/`, 규칙은 `samples/규칙.예시.yaml` → `rules/규칙.예시.yaml` 로 옮긴다.
> - 규칙에는 설계 5절에 없던 `행패턴`(표 행 정규식)과 "검사 순서가 의미 있음(DM 먼저)"이 들어 있다. 설계 5절을 그에 맞춰 고치고 2.3 테스트에 반영한다.
> - `samples/04_장부_엑셀/누적장부_기대결과.xlsx` 가 test_e2e 의 기대 상태다 (단, PO 행은 아래 엑셀 PO 범위 결정에 따라 제외될 수 있다).
> - `samples/01_고객오더_엑셀`·`02_회사양식_변환` 은 "고객 PO 엑셀 → 회사 SO 양식" 흐름의 예제다. 이 흐름의 1차 포함 여부는 저장소 관리자가 결정한다 — 포함이면 2단계 뒤에 **2-B단계(엑셀 PO 변환)** 를 추가하고, 미포함이면 PO 시트는 비워 두고 07장 뒤에 「08_선택_엑셀PO변환.md」 자리만 남긴다.
>
> **결정(2026-10-28)**: 엑셀 PO 변환은 **(b) 2차로 미룬다.** PO 시트는 비워 두고 08장은 자리만 둔다. 01·02 폴더 샘플은 2차 입력으로 커밋한다.
> 샘플 기준은 이 세트 하나로 통일한다. 기존 ACME 5종 PDF 는 지우고, 테스트 전용 생성기는 `tests/fixtures/build_fixtures.py` 로 옮긴다. 표 추출은 `행패턴` 한 방식만 남기고 `여러단어열` 은 없앤다.
> 이후 순서: 샘플 교체 → 검토 치명 C1~C4·중요 I1~I3 수정 → 로컬 전용 가드·Code/Cowork 설명 → docs 일반화·작성자 재작성 → 재리허설 → push(직전 확인).

### 작업 1.1 `samples/make_samples.py`
- reportlab 으로 3종 생성: `ACME_Invoice_INV-2026-0001.pdf`, `ACME_PO_PO-77001.pdf`, `ACME_DM_DM-5101.pdf`.
- 각 문서 1쪽: 상단 문서종류 제목, `Invoice No: …` / `PO Number: …` / `Debit Memo No: …`, `Date: 10/15/2026`, `Bill To: ACME Semiconductor Korea`, 품목 표 2~3줄(`Item` 헤더 ~ `Subtotal`), `Total: $12,345.67`.
- 추가 샘플 2종(오류 경로용): `Unknown_letter.pdf`(판별키워드 없음), `Scan_blank.pdf`(텍스트 없는 빈 페이지 1장).
- 생성된 PDF 5종을 커밋(사용자가 reportlab 없이도 03장 진행 가능).
- 완료 기준: `uv run python samples/make_samples.py` 로 5개 파일 생성, `pdfplumber` 로 열어 "Invoice No" 텍스트가 추출됨을 스크립트 끝에서 자가 확인 출력.

### 작업 1.2 `rules/규칙.예시.yaml`
- 설계 5절 구조대로 Invoice·PO·DM 3종. 1.1 샘플에서 모든 필드와 표가 추출되도록 패턴 작성.
- 완료 기준: 1.1 샘플 + 이 규칙으로 작업 2.3 테스트가 통과(2단계에서 확인). 커밋.

## 2단계 — 추출 (순수 함수)

### 작업 2.1 `app/rules.py`
- 테스트 `tests/test_rules.py`: ① 예시 규칙 로드 성공 ② 정규식 오류 → `RuleError`(한국어 메시지에 문서종류·필드명 포함) ③ 번호필드가 필드에 없음 → RuleError ④ 파일명 템플릿 변수가 필드에 없음 → RuleError ⑤ `형식` 값이 날짜·금액 외 → RuleError.
- 구현: `load_rules(path) -> Rules` (dataclass: 종류별 `DocType{name, keywords, number_field, fields{name: Field{pattern(compiled), fmt}}, table: Table|None, filename_tpl}`).
- 완료 기준: 테스트 5개 통과, 커밋.

### 작업 2.2 `app/extract.py` — 텍스트·판별
- 테스트 `tests/test_extract.py` (1부): ① 샘플 Invoice 1쪽 텍스트에 "Invoice No" 포함 ② `detect_type(text, rules)` 가 Invoice/PO/DM 를 맞게 반환 ③ Unknown → `None` ④ Scan_blank → 텍스트 길이 0.
- 구현: `read_text(pdf_path) -> str`(pdfplumber, 전 페이지 `\n` 결합), `detect_type(text, rules) -> DocType|None`(규칙 순서대로, 대소문자 무시).
- 완료 기준: 통과, 커밋.

### 작업 2.3 `app/extract.py` — 필드·표
- 테스트 (2부): ① Invoice 4필드 값 정확 ② 발행일이 `date` 객체 ③ 합계금액이 `Decimal('12345.67')` ④ 표 행 수와 열 값 ⑤ 번호필드 미매치 → `ExtractResult.missing_required=True` ⑥ 선택필드 미매치 → 값 `None` + `warnings` 에 기록 ⑦ 표 열 개수 불일치 줄 → 표 전체 `None` + warning.
- 구현: `extract(pdf_path, rules) -> ExtractResult{doc_type, fields{}, table rows|None, text_head(30줄), missing_required, warnings[]}`. 금액 변환: 콤마 제거 → Decimal. 날짜: `%m/%d/%Y` 우선, 실패 시 `%Y-%m-%d`.
- 완료 기준: 통과, 커밋.

## 3단계 — 장부

### 작업 3.1 `app/ledger.py`
- 테스트 `tests/test_ledger.py`: ① 없는 경로 → 시트 5개(Invoice·PO·DM·품목·처리로그) 생성, 헤더 = 고정열 + 규칙 필드 순서 ② 행 추가 후 재열기 시 값 유지, 날짜·금액 셀 서식 확인 ③ 같은 (종류, 번호) 재추가 → `is_duplicate` True, 행 수 불변 ④ 규칙에 필드 추가 후 열기 → 새 열이 끝에 추가, 기존 열 순서 유지 ⑤ `overwrite=True` → 기존 문서 행·품목 행 삭제 후 재기록 ⑥ 파일 잠김 시뮬레이션(다른 핸들로 열어 둠) → `LedgerLocked` 예외.
- 구현: `Ledger.open(path, rules)`, `.is_duplicate(doc_type, number)`, `.add(result, source_name, organized_path)`, `.remove(doc_type, number)`, `.save()` (임시 파일에 쓰고 교체). openpyxl `Table` 사용.
- 완료 기준: 통과, 커밋.

## 4단계 — 정리·로그

### 작업 4.1 `app/organize.py`
- 테스트 `tests/test_organize.py`: ① 파일명 템플릿 렌더(`{발행일:%Y-%m}` 지원, 경로에 못 쓰는 문자 `\/:*?"<>|` → `_`) ② 이동 후 원본 없음·목적지 있음 ③ 같은 이름 존재 → `_2`, 또 있으면 `_3` ④ 확인필요 이동 시 같은 이름 `.txt` 생성, 내용에 사유·text_head 포함 ⑤ 중복 → `받은문서/중복/`.
- 구현: `target_path(result, work_dir)`, `move_success(src, dst)`, `move_review(src, work_dir, reason, text_head)`, `move_duplicate(src, work_dir)`.
- 완료 기준: 통과, 커밋.

### 작업 4.2 `app/log.py`
- 테스트: CSV 한 줄 추가(utf-8-sig, 일시·파일·결과·사유), 콘솔 요약 문자열 "성공 N · 중복 N · 확인필요 N".
- 완료 기준: 통과, 커밋.

## 5단계 — 설치와 실행

### 작업 5.1 `app/config.py` + `app/init.py`
- 테스트 `tests/test_init.py`: `init(work_dir=<임시경로>, repo_root=<임시>)` 로 선택 창 우회 → ① 4개 폴더 생성 ② `규칙.yaml` 복사 ③ 기존 `규칙.yaml` 있으면 보존 ④ 저장소 하위 경로 → `InitError` ⑤ `작업폴더.txt` 내용 = 경로 ⑥ `config.work_dir()` 가 그 경로 반환, 파일 없으면 `ConfigError("설치.bat를 먼저 실행하세요")`.
- 구현: tkinter `askdirectory` 는 `work_dir` 인자가 없을 때만 호출. 취소 시 `input()` 대체. 재실행 시 현재 경로 표시 후 y/n.
- `설치.bat`: `@echo off` → `uv run python -m app init` → `pause`.
- 완료 기준: 통과, 커밋.

### 작업 5.2 `app/__main__.py` — `check`
- 테스트: `python -m app check samples/ACME_Invoice_….pdf` 출력에 필드명·값·표 행이 표 형태로 포함되고, 작업 폴더·장부에 아무 변화 없음.
- 구현: argparse. `check` 는 extract 만 호출해 표 출력(필드 / 값 / 비고(경고)).
- 완료 기준: 통과, 커밋.

### 작업 5.3 `app/__main__.py` — 기본 실행(파이프라인)
- 테스트 `tests/test_e2e.py`: 임시 작업 폴더에 샘플 5종 복사 → 실행 → ① 정리문서/ACME…/2026-10/ 에 3개 ② 확인필요/ 에 2개 + .txt 2개 ③ 받은문서/ 비었음 ④ 장부 Invoice·PO·DM 각 1행, 품목 행 수 = 샘플 표 줄 합 ⑤ 처리로그 5행 ⑥ 같은 샘플 다시 넣고 실행 → 중복 3건, 받은문서/중복/ 에 3개 ⑦ `--덮어쓰기` → 행 수 불변, 처리일시 갱신 ⑧ 장부를 열어 둔 상태 → "엑셀을 닫고 다시 실행" 출력, 파일 이동 0.
- 구현 순서(설계 8절): 받은문서/ 의 *.pdf 전부 extract → 결과 분류 → Ledger 열기(잠김이면 여기서 중단) → 추가/덮어쓰기 → save → 그 후에만 파일 이동 → 로그·요약.
- `run.bat`: `uv run python -m app %*` → `pause`.
- 완료 기준: 통과, 커밋. 전체 `uv run pytest` 녹색.

### 작업 5.4 (선택) `app/outlook.py`
- pywin32 로 Outlook `받은편지함/자동화` 폴더의 .pdf 첨부를 받은문서/ 로 저장(같은 이름 있으면 `_2`), 처리한 메일은 범주 "자동화완료" 표시. Outlook 없으면 한국어로 안내하고 종료.
- 테스트: COM 은 모킹으로 저장·중복 이름만 검증.
- `수집.bat`: `uv run python -m app outlook`.
- 완료 기준: 통과, 커밋. 설계 범위상 실패해도 1차 릴리스는 막지 않음.

## 6단계 — 문서 (설계 11절 원칙)

각 장은 「이 장이 끝나면 보이는 것」으로 시작, 명령은 블록 하나에 하나, 블록 위에 설명 한 줄, 끝에 「다음 →」.

### 작업 6.1 `README.md`
- 기존 `claude-guide-kr/README.md` 머리의 "파란 글씨 순서대로" 문구·표 양식 재사용. 읽는 순서 표(01~07, 99), 소요 시간, "이 저장소에 회사 정보는 없습니다 — 여러분 규칙은 PC에만" 박스. 요금제 안내(Pro 이상 필요, Code 탭).
### 작업 6.2 `01_구조_이해.md`
- GPT 와 다른 점(파일·폴더를 직접 다룸, 코드를 만들어 PC 에서 돌림). ASCII 흐름도(설계 3절). "Claude 가 보는 것 / 보지 않는 것" 표.
### 작업 6.3 `02_설치.md`
- 전제: 친구 PC에는 Node.js·Python·uv 가 전혀 없다. 02장 첫 줄에 "이 장은 아무것도 안 깔린 PC 기준"이라고 적는다.
- 순서: Claude Desktop 설치·로그인·**Code 탭 열기** → uv 설치 → Node.js LTS 설치 → 저장소 ZIP 다운로드·풀기(`C:\dev\` 권장) → Code 탭에서 폴더 열기 → `설치.bat` → 폴더 선택("내 PC 문서 폴더 권장, 저장소 폴더는 거부됨, OneDrive 는 동기화 잠김 주의").
- uv 설치는 두 길을 순서대로: ① PowerShell에서 `winget install --id astral-sh.uv -e` ② winget 이 없거나 회사 정책으로 막히면 uv 공식 PowerShell 설치 명령(구현 시 docs.astral.sh 에서 현재 명령을 확인해 적는다). 설치 뒤 **PowerShell 을 완전히 닫고 다시 열어야** `uv --version` 이 된다는 줄을 반드시 넣는다. uv 를 깔면 `uvx` 도 같이 생긴다.
- Node.js LTS 도 같이 깐다: `winget install --id OpenJS.NodeJS.LTS -e` (막히면 nodejs.org 의 LTS 설치 파일). 확인은 새 PowerShell 에서 `node --version`. 박스로 설명: "이 저장소는 Node.js 를 쓰지 않지만, 나중에 MCP 같은 확장을 붙일 때 반드시 필요하므로 지금 함께 깔아 둡니다. 03장은 Node.js 없이도 됩니다."
- "Python 은 따로 깔지 않습니다 — uv 가 필요할 때 알아서 내려받습니다." 를 박스로. 다른 블로그 보고 괜히 깔지 않게.
- 설치 확인 표: `uv --version` · `uvx --version` · `node --version` 세 줄이 모두 버전을 찍으면 02장 끝.
- `설치.bat` 첫 실행 때 uv 가 Python 을 내려받느라 1~3분 걸리고 검은 창에 영어가 흐르는 게 정상이라고 적는다 (03장 앞에도 한 줄).
### 작업 6.4 `03_첫_실행_10분.md`
- `samples/` 5종을 받은문서/ 에 복사 → `run.bat` → 보여야 할 콘솔 요약 "성공 3 · 중복 0 · 확인필요 2" → 장부 열어 확인 → 확인필요/ .txt 열어 읽기 → 재실행으로 중복 3 체험.
### 작업 6.5 `04_내_양식_가르치기.md`
- 샘플 가리는 법(PDF 에서 고객명·금액 바꿔 저장하거나, 그대로 쓰되 "이 PC 안에서만 처리된다"는 1장 설명 다시 보기). Claude Code 프롬프트 레시피 3개(Invoice→PO→DM) 코드 블록. `check` 결과 읽는 법, 고치는 대화 예시 3개("고객명이 비었어", "날짜가 안 잡혀", "표가 두 줄만 나와"). 완료 판정: check 표에 비고 없음.
### 작업 6.6 `05_매일_쓰기.md`
- 아침 루틴 5줄. 확인필요 루프 그림. `--덮어쓰기` 쓰는 경우. 기존 수기 장부 병행 한 달 권고.
### 작업 6.7 `06_프로젝트와_예정됨.md`
- 프로젝트 만들기, 저장소 .md 전부 지식에 추가, 지침 붙여넣기(짧은 버전 제공). 예정됨: 문서 내용 없이 돌릴 일만 — 예: 매주 월 "처리로그.csv 지난주 요약", "확인필요/ 에 쌓인 파일 수 알림". 하지 말 것: PDF 읽게 하는 예정 작업.
### 작업 6.8 `07_선택_Outlook_자동수집.md`
- 전제(Outlook 데스크톱 앱, 회사 정책 확인 문구 박스), `uv sync --extra outlook`, Outlook 에 "자동화" 폴더 만들고 규칙으로 고객 메일 이동, `수집.bat`.
### 작업 6.9 `99_막혔을_때.md`
- 증상별 표: Code 탭 안 보임(요금제, 첫 줄) / `uv`는 내부 또는 외부 명령이 아닙니다(PowerShell 재시작·PATH) / winget 없음·막힘(공식 스크립트) / 설치.bat 에서 Python 내려받다 멈춤(회사 프록시·방화벽 — IT 에 문의할 문구 제공) / 설치.bat 창이 안 뜸 / "설치.bat 를 먼저 실행하세요" / 엑셀 잠김 / 확인필요만 쌓임 / 한글 깨짐(utf-8-sig). 각 줄에 Claude Code 에 붙일 질문 문구.
- 완료 기준(6단계 공통): 모든 링크가 존재하는 파일을 가리킴(링크 검사 스크립트 `tests/test_docs_links.py`), 각 장에 「다음 →」, 커밋.

## 7단계 — 검증·배포

### 작업 7.1 새 PC 시나리오 리허설
- 다른 Windows 사용자 계정(또는 깨끗한 VM)에서 02→03→04(샘플로) 를 문서만 보고 수행. 막힌 지점을 99장과 해당 장에 반영.
### 작업 7.2 GitHub 공개 저장소
- `gh repo create office-automation-kr --public`, push. README 첫 화면에 "Code → Download ZIP" 캡처 대신 글자 경로 안내.
- 저장소 주소를 개인 메모에 기록(저장소 밖).
### 작업 7.3 친구에게 보낼 안내문
- 메신저 한 통 분량: 저장소 주소, "README 부터, 02·03 은 30분, 04 는 샘플 하나당 20분", 요금제 안내. 기존 메신저 안내 양식 참고(저장소 밖).

## 완료 정의

- `uv run pytest` 전부 통과.
- 깨끗한 PC 에서 문서만으로 03장 "성공 3 · 중복 0 · 확인필요 2" 재현.
- 저장소 전체 grep 으로 회사명·실제 고객명·실제 번호 0건.

## 추가 (2026-10-28) 한 줄 자동 설치

- 요청: 「옆에서 도와주지 않으면 이해할 수 없다. Claude Desktop 설치부터 자동으로, 설명은 빼고, 순서는 Claude Desktop → Node.js → uv·uvx」.
- 결정: `install.ps1` 을 PowerShell 한 줄(`irm … | iex`)로 실행. 저장소 ZIP 받기와 작업 폴더(`C:\사무자동화`) 생성까지 자동. 손으로 하는 방법은 99장 1절로 옮김.
- 시험: `tests/test_install_script.py` (-DryRun 으로 1~3단계 설치는 건너뛰고 4·5단계 실제 실행).
