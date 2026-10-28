# 작업지시서 · 여러 작업 · 매일 자동 실행 · 이중 확인 구현 계획

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 받은문서 하나 · run.bat 한 번의 처리 방식을, 친구가 Code 탭 대화로 만드는 「작업」 여러 개(작업마다 입력·출력 폴더, 매일 16:00부터 20분 간격 자동 실행, 바탕화면 지금 실행 · 승인, 숫자 이중 확인, 양식별 처음 10건 승인)로 바꾼다.

**Architecture:** 작업 한 개 = 작업 폴더 `작업\NN_제목\` (작업설정.yaml · 규칙.yaml · 상태 폴더). 기존 추출 엔진(extract · excel_extract · rules · ledger · outputs)은 그대로 쓰고, 그 위에 jobs(작업 모델) · verify(검산) · approval(승인) · job_run(작업 실행) · schedule/shortcuts(Windows 등록)를 얹는다. 매일 실행은 Windows 작업 스케줄러가 `.venv\Scripts\pythonw.exe -m app run NN --자동` 을 부른다. Claude 는 작업 만들기 대화에서만 쓴다.

**Tech Stack:** Python 3.13 (uv) · pdfplumber · openpyxl · PyYAML · pytest · PowerShell(Register-ScheduledTask, WScript.Shell 바로가기) · msvcrt(실행 잠금) · ctypes(MessageBoxW 알림)

**Spec:** `docs/superpowers/specs/2026-10-28-work-orders-design.md` (기존 설계 `docs/superpowers/specs/2026-10-28-office-automation-kr-design.md` 위에 덧붙임. 겹치면 새 설계가 이긴다)

## Global Constraints

- 사용자에게 보이는 모든 출력 · 문서 · 커밋 메시지는 한국어. 프로그램 메시지에 em dash(—) 쓰지 않음(콘솔 깨짐).
- 커밋 메시지 첫 줄 「W번호 제목」(예: `W1 작업 모델`). 끝에 `Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>`.
- 한 작업 = 테스트 먼저(실패 확인) → 구현 → `uv run pytest` 전체 녹색 → 커밋 1회.
- 회사명 · 실제 고객명 · 실제 양식 금지. 샘플은 가짜(ACME · Hanil). 달러 금액은 숫자 그대로.
- 금액은 Decimal 만. 반올림 · 잘라내기 금지(줄 검산 비교 때만 금액 자리로 반올림).
- 매일 실행 · 지금 실행 · 승인은 Claude 없이 돈다. 실행 중 네트워크는 `local_guard.block_network()` 로 막힌 상태(기존 `main` 경로 유지).
- 입력 · 출력 · 작업 폴더는 로컬 디스크만(`local_guard.check_work_dir` 재사용). OneDrive · Dropbox · 네트워크 · UNC 거부.
- 일정: 작업 번호 n 의 시각 = 16:00 + 20분 × (n − 1). 놓친 실행은 PC 를 켜면 실행.
- 승인: 양식(엑셀양식 이름, 없으면 문서종류 이름) 단위로 10건. 그 양식의 규칙이 바뀌면 그 양식만 0.
- 회사 양식 파일에는 시트를 더하지 않는다. 검산 기록은 `검산_{번호}.xlsx` 로 따로.
- 테스트는 실제 스케줄러 등록 · 실제 바탕화면 · 실제 메시지 창을 건드리지 않는다(실행기 주입 · 임시 폴더).

## Review Focus

1. 엑셀 원본의 부동소수 값(셀에 12.3, 0.1, 1234.5 가 float 로 저장, 표시 형식 `0.00`): 원문 대조가 정상 값을 막지 않고, 자리가 다른 값(123.4)은 막아야 한다. → W2 `test_excel_float_values_pass_and_wrong_scale_fails`
2. 오후 4시에 아직 저장 중인 파일(Outlook 저장 · 복사 중, 수정 시각이 방금): 손대지 않고 입력 폴더에 남겨 다음 실행으로. → W4 `test_file_modified_within_a_minute_is_left_for_next_run`
3. 출력 폴더나 입력 폴더가 사라짐(USB 뺌 · 이름 바꿈): 파일을 하나도 옮기지 않고 한국어로 멈춤. → W4 `test_missing_output_folder_stops_without_moving`
4. 제목에 폴더 이름으로 못 쓰는 글자(`PO/DM 정리?`) · 같은 제목 두 번: 작업 폴더가 깨지지 않고 번호로 구분. → W1 `test_title_with_forbidden_characters_and_duplicates`
5. PC 를 늦게 켜서 놓친 작업이 여러 개: 로그인 때 하나씩 차례로 돌고, 같은 작업이 두 번(자동 + 지금 실행) 겹쳐 돌지 않는다. → W4 `test_second_run_while_locked_exits_with_message`, W7 `test_catch_up_runs_missed_jobs_in_order_once`

---

## 파일 구조

| 파일 | 책임 | 작업 |
|---|---|---|
| `app/jobs.py` (새) | 작업 모델 · 작업설정.yaml 읽기/쓰기 · 번호/시각 배정 · 작업 찾기 | W1 |
| `app/rules.py` (수정) | 문서종류에 `검산:` 키(소계 · 부대비용 · 합계) | W2 |
| `app/verify.py` (새) | 원문 대조 · 줄 검산 · 합계 검산 · 소수 자리 | W2 |
| `app/outputs.py` (수정) | `write_check_file` · `read_back` | W3 |
| `app/job_run.py` (새) | 작업 한 개 실행(판별된 파일만 · 검산 · 승인 대기 · 출력 방식 · 이동 · 잠금) | W4 |
| `app/approval.py` (새) | 양식 키 · 규칙 지문 · 승인 수 · 승인 처리 | W5 |
| `app/__main__.py` (수정) | 명령 `job` · `run` · `approve` · `check --작업` · `pick-folder` · `catch-up` · 옛 받은문서 흐름 삭제 | W6 |
| `app/init.py` · `app/config.py` (수정) | 작업 폴더 = `작업\` + 고객마스터 | W6 |
| `app/shortcuts.py` (새) | 바탕화면 · 시작프로그램 바로가기 | W7 |
| `app/schedule.py` (새) | 작업 스케줄러 등록 · 확인 · 해제, 대체 방법, catch-up, 알림 | W7 |
| `작업지시서.md` (새) · `CLAUDE.md` · 장 문서 · `install.ps1` | 대화 순서 · 약속 · 안내 · 설치 | W8 |

---

### Task W1: 작업 모델

**Files:**
- Create: `app/jobs.py`
- Test: `tests/test_jobs.py`

**Interfaces:**
- Produces:
  - `JOBS_DIR = "작업"`, `JOB_FILE = "작업설정.yaml"`, `OUTPUT_MODES = ("회사양식", "누적장부", "둘다")`, `JOB_SUBFOLDERS = ("승인대기", "확인필요", "검산")` (설계 12-1: 정리문서 없음)
  - `@dataclass Job: folder: Path; number: int; title: str; input_dir: Path; output_dir: Path; output_mode: str; run_time: str; scheduler: str = ""; paused: bool = False; approvals: dict[str, int]; fingerprints: dict[str, str]; last_run: str = ""` (last_run = `YYYY-MM-DD`)
  - `Job.rules_path -> Path` (`folder / "규칙.yaml"`), `Job.name -> str` (`folder.name`)
  - `slot_time(number: int) -> str`
  - `create_job(work: Path, title: str, input_dir, output_dir, output_mode: str) -> Job`
  - `load_job(folder: Path) -> Job`, `save_job(job: Job) -> None`
  - `list_jobs(work: Path) -> list[Job]` (번호 순)
  - `find_job(work: Path, key: str) -> Job` (`"1"` · `"01"` · 제목 · 폴더 이름. 없으면 `JobError("작업 'X' 을 찾지 못했습니다. 지금 있는 작업: …")`)
  - `class JobError(Exception)`

- [ ] **Step 1: 실패하는 테스트**

```python
def test_slot_times():
    assert [slot_time(n) for n in (1, 2, 3, 4, 5)] == ["16:00", "16:20", "16:40", "17:00", "17:20"]

def test_create_job_makes_folder_and_settings(tmp_path):
    work, inp, out = tmp_path / "w", tmp_path / "in", tmp_path / "out"
    for d in (work, inp, out): d.mkdir()
    job = create_job(work, "일일 PO 정리", inp, out, "둘다")
    assert job.number == 1 and job.run_time == "16:00"
    assert job.folder == work / "작업" / "01_일일 PO 정리"
    assert all((job.folder / s).is_dir() for s in JOB_SUBFOLDERS)
    again = load_job(job.folder)
    assert (again.title, again.input_dir, again.output_mode, again.approvals) == ("일일 PO 정리", inp, "둘다", {})

def test_second_job_gets_next_slot(tmp_path): ...  # 번호 2, "16:20"

def test_title_with_forbidden_characters_and_duplicates(tmp_path):
    a = create_job(work, "PO/DM 정리?", inp, out, "누적장부")
    b = create_job(work, "PO/DM 정리?", inp, out, "누적장부")
    assert a.folder.name == "01_PO_DM 정리_" and b.folder.name == "02_PO_DM 정리_"
    assert find_job(work, "2").folder == b.folder and find_job(work, "02_PO_DM 정리_").folder == b.folder

def test_rejects_unknown_output_mode_and_missing_or_sync_folder(tmp_path):
    # output_mode "엑셀" → JobError, 없는 입력 폴더 → JobError("입력 폴더가 없습니다: …"),
    # 이름에 OneDrive 가 든 폴더 → JobError (check_work_dir 메시지)

def test_find_job_unknown_lists_existing(tmp_path):
    with pytest.raises(JobError, match="01_일일 PO 정리"):
        find_job(work, "없는작업")
```

- [ ] **Step 2: 실행해 실패 확인** — `uv run pytest tests/test_jobs.py -q` → ImportError(jobs 없음)
- [ ] **Step 3: 구현** — 폴더 이름 = `f"{번호:02d}_{제목}"`, 제목은 `organize.FORBIDDEN` 글자를 `_` 로, 앞뒤 공백 · 끝 마침표 제거, 30자 자름. 번호 = 기존 작업 최대 번호 + 1(지운 번호는 다시 쓰지 않음). yaml 은 `allow_unicode=True`, 경로는 문자열로 저장. 입력 · 출력 폴더는 `check_work_dir` 를 통과해야 하고 존재해야 함.
- [ ] **Step 4: 통과 확인** — `uv run pytest -q` 전체 녹색
- [ ] **Step 5: 커밋** — `W1 작업 모델`

---

### Task W2: 검산 (원문 대조 · 줄 · 합계 · 소수 자리)

**Files:**
- Modify: `app/rules.py` (DocType 에 `check: CheckSpec | None`, `_doc_type` 에서 `검산:` 읽기)
- Modify: `rules/규칙.예시.yaml` (Invoice 에 `검산: {소계: 소계, 부대비용: [운임], 합계: 합계금액}`, DM 에 `검산: {합계: 합계금액}`)
- Create: `app/verify.py`
- Test: `tests/test_verify.py`, `tests/test_rules.py`(검산 키 검증 추가)

**Interfaces:**
- Consumes: `ExtractResult`(fields · table · doc_type · is_excel), `Rules`
- Produces:
  - `@dataclass CheckSpec: subtotal: str | None; extras: list[str]; total: str | None` — 이름은 그 문서종류의 필드 이름이어야 함(아니면 `RuleError("검산의 '운임' 은 필드 목록에 없습니다")`)
  - `@dataclass Check: item: str; source: str; output: str; ok: bool | None` (None = 해당 없음)
  - `verify(result: ExtractResult, source_path: Path) -> list[Check]`
  - `failed(checks: list[Check]) -> list[Check]` (ok is False 만)

검산 규칙(설계 5-1 · 10절 5):
1. 원문 대조: `result.fields` 와 `result.table` 의 값 중 Decimal · int · date 와 번호필드 값은 원본에서 다시 읽은 값 모음에 있어야 함. PDF = `extract.read_text` 의 숫자 토큰(`,` `$` 제거 → Decimal, 자리 유지)과 날짜 토큰(`to_date` 로 바뀌는 것). 엑셀 = 모든 시트 셀 값(float 은 `Decimal(repr(v))`, 문자열 숫자도 포함). 비교는 Decimal 값 동등(`1234.50 == 1234.5`). 문자열 번호필드는 원본 글자 안에 그대로 있어야 함.
2. 소수 자리: PDF 는 결과 Decimal 의 소수 자리 수 == 같은 값을 준 원본 토큰의 자리 수. 엑셀은 1번 동등으로 갈음(float 은 자리 정보가 없음).
3. 줄 검산: 품목 줄에 `수량` · `단가` · `금액` 이 모두 있으면 `(수량 × 단가)` 를 금액의 소수 자리로 반올림(ROUND_HALF_UP 또는 ROUND_HALF_EVEN 중 하나라도 같으면 통과) == 금액.
4. 합계 검산: `check.subtotal` 있으면 품목 금액 합 == 소계. `check.total` 있고 `check.extras` 또는 subtotal 이 있으면 (소계 또는 품목 합) + 부대비용 합 == 합계. subtotal · extras 가 없고 total 만 있으면 품목 금액 합 == 합계. 값이 비면 해당 없음.

- [ ] **Step 1: 실패하는 테스트** (PDF 는 `tests/fixtures` 방식처럼 reportlab 으로 임시 PDF, 엑셀은 openpyxl 로 임시 파일)

```python
def test_sample_invoice_passes_all_checks():
    # samples/03_Invoice_DM_PDF 의 Invoice 1개 + 예시 규칙 → failed(verify(...)) == []

def test_wrong_digit_or_decimal_fails_number_match(tmp_path):
    result = extract(pdf, rules); result.fields["합계금액"] = Decimal("12345.00")  # 원본 1234.50
    assert [c.item for c in failed(verify(result, pdf))] == ["합계금액"]

def test_same_value_different_notation_passes():  # 원본 "1,234.50" ↔ 결과 Decimal("1234.50")
def test_decimal_places_must_match_pdf_token():   # 원본 "12.50" ↔ 결과 Decimal("12.5") → 실패 항목 "합계금액 소수 자리"
def test_line_check_allows_rounding():             # 수량 3, 단가 0.333, 금액 1.00 → 통과 / 금액 1.10 → 실패
def test_total_check_with_freight():               # 소계 100.00 + 운임 15.00 == 합계 115.00 통과, 합계 116.00 실패
def test_missing_values_are_not_applicable():      # 단가 없는 표 → 줄 검산 Check.ok is None
def test_excel_float_values_pass_and_wrong_scale_fails(tmp_path):
    # 셀 F9 = 12.3(float), number_format "0.00", 결과 Decimal("12.30") 통과 · Decimal("123.0") 실패
def test_check_spec_unknown_field_is_rule_error(tmp_path)
```

- [ ] **Step 2: 실행해 실패 확인** — `uv run pytest tests/test_verify.py tests/test_rules.py -q`
- [ ] **Step 3: 구현** — `verify.py` 는 위 1~4를 각각 작은 함수로. `Check.item` 이름: 필드 이름, `"{n}줄 금액"`, `"품목 합 = 소계"`, `"소계 + 부대비용 = 합계"`, `"{필드} 소수 자리"`.
- [ ] **Step 4: 통과 확인** — `uv run pytest -q`
- [ ] **Step 5: 커밋** — `W2 검산`

---

### Task W3: 검산 파일 · 저장 후 다시 읽기

**Files:**
- Modify: `app/outputs.py`
- Test: `tests/test_outputs.py`

**Interfaces:**
- Consumes: `Check` (W2), 기존 `company_rows`, `write_company_form`, `write_extract_result`
- Produces:
  - `write_check_file(folder: Path, result, checks: list[Check], source_name: str) -> Path` — `검산_{번호필드}.xlsx`, 시트 「검산」 열 `항목 · 원본 값 · 결과 값 · 일치`(일치 = `일치` · `불일치` · `해당 없음`), 맨 위 줄에 `원본파일`.
  - `read_back(path: Path, sheet: str, header: list[str], rows: list[list]) -> list[str]` — 저장된 파일을 다시 열어 머리글과 각 칸 값을 비교, 다른 칸 설명 목록(`"3행 Unit Price: 쓴 값 12.50, 읽은 값 12.5"`; Decimal 은 값 동등이면 같음, 날짜는 date 로 비교). 같으면 `[]`.
  - `extract_result_rows(result, source_name) -> tuple[list[str], list[list]]` (write_extract_result 가 쓰는 추출결과 시트의 머리글 · 행. read_back 용)

- [ ] **Step 1: 실패하는 테스트**

```python
def test_check_file_has_one_row_per_check(tmp_path)        # 3개 Check → 데이터 3줄, 일치 열 값 확인
def test_read_back_matches_company_form(tmp_path)           # write_company_form 직후 read_back(...) == []
def test_read_back_reports_changed_cell(tmp_path)           # 저장 뒤 openpyxl 로 한 칸 바꿔 저장 → 설명 1개
def test_company_form_from_pdf_result(tmp_path)             # PDF 추출 결과(표 있음) + form → 줄 수 = 품목 수
```

- [ ] **Step 2~4:** 실패 확인 → 구현 → `uv run pytest -q`
- [ ] **Step 5: 커밋** — `W3 검산 파일과 다시 읽기`

---

### Task W4: 작업 실행

**Files:**
- Create: `app/job_run.py`
- Test: `tests/test_job_run.py`, `tests/test_e2e_jobs.py` (기존 `test_e2e.py` · `test_e2e_excel.py` 시나리오를 작업 방식으로 옮김. 옛 파일은 W6 에서 삭제)

**Interfaces:**
- Consumes: W1 `Job`, `save_job`; W2 `verify`, `failed`; W3 `write_check_file`, `read_back`; 기존 `extract`, `read_excel`, `Ledger`, `load_master`, `lookup`, `write_company_form`, `write_extract_result`, `move_review`, `free_name`, `render_filename`, `append_log`, `summary`, `is_locked`
- Produces:
  - `run_job(job: Job, work: Path, *, now=None, source_dir: Path | None = None, gate=None, on_success=None) -> RunReport` — `source_dir` 기본 = `job.input_dir`(W5 승인은 `승인대기\` 를 넘김). 하위 폴더는 보지 않는다. `on_success(job, result)` 는 출력까지 끝난 건마다 호출(W5 승인 수 세기)
  - `@dataclass RunReport: counts: dict[str, int]; unmatched_old: int; message: str; output_files: list[Path]` — counts 키: `성공` · `중복` · `확인필요` · `승인대기` · `보류`
  - `RunReport.has_news -> bool` (성공 · 중복 · 확인필요 · 승인대기 중 하나라도 > 0 이거나 unmatched_old > 0)
  - `class JobLocked(Exception)`, `job_lock(job) -> contextmanager` (`folder/실행중.lock` 을 열어 `msvcrt.locking(LK_NBLCK)`; 못 잡으면 `JobLocked("이 작업이 이미 실행 중입니다 …")`)
  - `LEDGER_NAME = "누적장부_{제목}.xlsx"` → 출력 폴더에 둠
  - `RECORD_FILE = "처리기록.csv"` → 작업 폴더. 열: `지문,파일,문서종류,번호,상태,규칙지문,날짜` (utf-8-sig)
  - `file_fingerprint(path) -> str` (내용 sha256 hex), `load_records(job) -> list[dict]`, `rules_fingerprint(job) -> str` (규칙.yaml 바이트 sha256 앞 16자)
  - `gate(job, rules, result) -> bool` : True 면 승인대기로 보냄. 기본값은 항상 False (W5 가 승인 규칙을 넣음)

실행 순서(작업 한 개, 설계 12-1: **원본은 옮기지 않는다**):
1. 입력 · 출력 폴더가 없으면 `RunReport(message="입력 폴더가 없습니다: …")` 로 끝(아무것도 안 씀).
2. `source_dir` 의 `.pdf` · `.xlsx` (`~$` 제외, 수정 시각 60초 이내 제외, 열려 있으면 보류). `.xls` 는 판별 못 하므로 건너뜀.
3. 건너뛰기: 지문이 처리기록에 상태 성공 · 중복 · 승인대기로 있으면 건너뜀(셈에 넣지 않음). 상태 확인필요이면 날짜가 오늘이고 규칙지문이 같을 때만 건너뜀. (W5 승인 실행은 `source_dir=승인대기` 이고 승인대기 기록은 건너뛰지 않음)
4. 추출. **판별 실패(doc_type None) 파일은 기록하지 않고 그대로 둔다.** 텍스트 없음 · 번호 미추출 · 읽기 오류는 확인필요.
5. 중복: 지문은 처음인데 (문서종류, 번호)가 처리기록에 성공으로 있으면 상태 중복으로 기록.
6. `failed(verify(...))` 가 있으면 확인필요(사유에 항목 나열).
7. 회사양식 · 둘다: 고객마스터(`work / 고객마스터.xlsx`) 조회. 없으면 확인필요(기존 문구). 회사 양식은 엑셀 PO 뿐 아니라 PDF 결과에서도 만든다(설계 10절 2).
8. `gate(...)` True → 원본 **복사본**을 `승인대기\` 에(`shutil.copy2`, 이름 겹치면 `free_name`), 대조표(검산 파일)는 `승인대기\대조표\` 에. 상태 승인대기로 기록. 끝.
9. 출력: 누적장부 · 둘다 → 출력 폴더의 `누적장부_{제목}.xlsx` 에 `Ledger` 로 추가 · 저장. 회사양식 · 둘다 → `write_company_form(출력 폴더, …)`. 누적장부만이면 문서별 결과 파일 없음. 저장한 파일마다 `read_back` 이 `[]` 이 아니면: 이번 건을 장부에서 되돌리고(`Ledger.remove` 후 저장) 회사 양식 파일을 지우고 확인필요.
10. 검산 파일 → 작업 폴더 `검산\`. 상태 성공으로 기록. `on_success` 호출.
11. 확인필요는 작업 폴더 `확인필요\{원본이름}.txt` 에 원본 경로 · 사유만(원본 복사 없음).
12. `job.last_run = 오늘` 저장. `unmatched_old` = 입력 폴더 파일 중 수정 시각 24시간 넘었고, **같은 입력 폴더를 쓰는 모든 작업**의 처리기록 어디에도 지문이 없는 것의 개수. `RunReport.output_files` 에 이번에 만든 결과 파일 경로들.

- [ ] **Step 1: 실패하는 테스트** (샘플 PDF · 엑셀 · 예시 규칙 · 예시 고객마스터를 임시 작업에 복사)

```python
def test_runs_only_detected_files_and_leaves_others(tmp_path)   # 서신 · Invoice 모두 입력 폴더에 그대로, Invoice 만 처리기록에 성공
def test_second_run_skips_processed_even_if_renamed(tmp_path)   # 같은 파일 이름 바꿔 다시 넣어도 counts 전부 0, has_news False
def test_review_file_retried_next_day_or_after_rule_change(tmp_path)  # now=같은 날 → 건너뜀, 다음 날 → 다시 시도, 규칙.yaml 바꾸면 같은 날도 다시
def test_new_file_with_same_number_is_duplicate(tmp_path)       # 내용 다른 파일 · 같은 인보이스번호 → counts["중복"] == 1, 출력 없음
def test_mode_ledger_only_writes_ledger_and_no_company_file(tmp_path)
def test_mode_company_only_writes_company_file_and_no_ledger(tmp_path)
def test_mode_both_from_excel_po(tmp_path)                       # 회사양식 파일 1개 + 누적장부 PO 시트 1줄
def test_company_form_from_pdf_po_when_mode_company(tmp_path)    # PDF PO(픽스처) → 회사양식 파일
def test_verify_failure_goes_to_review_and_nothing_written(tmp_path)   # verify 를 실패하게 monkeypatch → 확인필요, 출력 폴더 비어 있음
def test_read_back_mismatch_rolls_back(tmp_path)                 # read_back 이 ["…"] 를 돌려주게 → 확인필요, 장부 줄 없음, 회사 파일 없음
def test_gate_true_copies_to_waiting_with_check_file(tmp_path)   # 원본은 입력 폴더에 그대로, 복사본 + 대조표
def test_file_modified_within_a_minute_is_left_for_next_run(tmp_path)
def test_missing_output_folder_stops_without_moving(tmp_path)
def test_second_run_while_locked_exits_with_message(tmp_path)    # job_lock 안에서 run_job → JobLocked
def test_two_jobs_sharing_input_folder_take_their_own_files(tmp_path)  # 작업1 규칙=Invoice만, 작업2=DM만
def test_unmatched_old_file_is_counted(tmp_path)                 # os.utime 으로 25시간 전, 다른 작업이 처리한 파일은 세지 않음
```

- [ ] **Step 2: 실행해 실패 확인** — `uv run pytest tests/test_job_run.py tests/test_e2e_jobs.py -q`
- [ ] **Step 3: 구현** — 기존 `__main__.run` · `_plan` · `_write_output` 의 예외 처리 습관(한 건 오류가 나머지를 막지 않음, 장부 저장이 된 뒤에만 기록 · 출력)을 따른다. 입력 폴더의 원본은 읽기만 한다(이동 · 이름 바꾸기 · 삭제 금지 — 테스트로 확인).
- [ ] **Step 4: 통과 확인** — `uv run pytest -q`
- [ ] **Step 5: 커밋** — `W4 작업 실행`

---

### Task W5: 처음 10건 승인

**Files:**
- Create: `app/approval.py`
- Test: `tests/test_approval.py`

**Interfaces:**
- Consumes: W1 `Job`, `save_job`; W4 `run_job(..., approved=, gate=)`
- Produces:
  - `REQUIRED = 10`
  - `form_key(result) -> str` (엑셀이면 `result.source_form`, 아니면 `result.doc_type.name`)
  - `form_fingerprint(rules_path: Path, key: str) -> str` — 규칙.yaml 을 yaml 로 읽어 그 양식에 해당하는 부분(엑셀양식이면 그 엑셀양식 + 연결된 문서종류, 아니면 그 문서종류) + `회사양식` 을 `yaml.safe_dump(sort_keys=True)` 한 글자의 sha256 앞 16자
  - `approval_gate(job, rules, result) -> bool` — 지문이 저장된 값과 다르면 그 양식만 승인 수 0 + 새 지문 저장. 승인 수 < 10 이면 True
  - `approve(job, work) -> RunReport` — `run_job(job, work, source_dir=job.folder/"승인대기", gate=lambda *a: False, on_success=…)`: 지금 규칙으로 다시 추출 · 검산해 출력까지 보내고, 성공한 건마다 그 양식 승인 수 +1, 그 복사본과 `승인대기\대조표\` 의 대조표 삭제(처리기록의 그 지문은 성공으로 바뀜). 승인대기가 비었으면 `message="승인할 문서가 없습니다."`
  - 친구가 틀린 문서를 승인하지 않으려면 승인 전에 `승인대기\` 의 그 복사본을 지운다(06장에 적음).

- [ ] **Step 1: 실패하는 테스트**

```python
def test_new_form_waits_for_approval(tmp_path)          # 처음 실행 → counts["승인대기"] == 1, 출력 폴더 비어 있음
def test_approve_moves_to_output_and_counts(tmp_path)   # approve → 출력 생김, job.approvals["ACME"] == 1
def test_after_ten_approvals_goes_straight_out(tmp_path)  # approvals["ACME"]=10 + 지문 같음 → 승인대기 0
def test_rule_change_resets_only_that_form(tmp_path)    # ACME 엑셀양식 머리 칸 하나 바꿈 → ACME 0, HANIL 10 그대로
def test_approve_with_empty_waiting(tmp_path)
def test_approve_rechecks_and_rejects_bad_file(tmp_path)  # 승인대기 파일이 검산 실패 → 확인필요, 승인 수 그대로
```

- [ ] **Step 2~4:** 실패 확인 → 구현 → `uv run pytest -q`
- [ ] **Step 5: 커밋** — `W5 처음 10건 승인`

---

### Task W6: 명령과 옛 흐름 정리

**Files:**
- Modify: `app/__main__.py`, `app/init.py`, `app/config.py`, `app/outlook.py`(받는 곳 = 작업 입력 폴더), `run.bat`
- Delete: `수집.bat`, `tests/test_e2e.py`, `tests/test_e2e_excel.py` (시나리오는 W4 `test_e2e_jobs.py` 로 옮겨졌는지 먼저 확인)
- Modify tests: `tests/test_check.py`, `tests/test_init.py`, `tests/test_outlook.py`, `tests/test_review_fixes.py`(받은문서 전제인 것은 작업 기준으로 고치거나, W4 에 같은 뜻의 테스트가 있으면 삭제), `tests/test_demo.py`
- Test: `tests/test_cli_jobs.py`

**Interfaces:**
- Consumes: W1~W5
- Produces (명령, 모두 `main()` 의 가드 · 네트워크 차단 안에서):
  - `job new --제목 T --입력 DIR --출력 DIR --방식 회사양식|누적장부|둘다` → 작업 생성, 작업 폴더에 예시 규칙 복사 대신 **빈 뼈대 규칙**(`rules/규칙.뼈대.yaml`: 주석으로 구조 안내 + 비어 있는 `문서종류: {}` 가 아니라 Claude 가 채울 자리. load 시 RuleError 가 나도 job new 는 성공) 을 둔다. 마지막 줄 출력 `작업 만듦: 01_제목 (매일 16:00)`.
  - `job list` → 번호 · 제목 · 시각 · 실행 방법(scheduler) · 승인 현황(`ACME 3/10`).
  - `job delete N` · `job pause N` · `job resume N` (W7 의 등록 해제 · 재등록 호출)
  - `run N [--자동]` → `run_job` + `approval_gate`. 결과를 작업 폴더 `실행기록.txt` 에 덧붙이고, `report.has_news` 이면 W7 `notify_report(job, report)` (자동 · 지금 실행 모두. 설계 12-2). W7 전에는 `notify_report` 자리에 아무것도 하지 않는 함수.
  - `approve N` → `approve` + 같은 알림
  - `check FILE --작업 N` → 그 작업 규칙으로 기존 check 표 + 검산 결과 표(`| 검산 | 원본 | 결과 | 일치 |`)
  - `pick-folder "안내 글"` → 폴더 선택 창, 고른 경로를 한 줄 출력(로컬 아니면 오류 문구, 취소면 `취소함`)
  - `catch-up` → W7
  - 인자 없음 → 작업 목록과 「바탕화면의 「지금 실행 - 작업 이름」을 쓰세요」 안내(받은문서 처리 없음)
  - `init [경로]` → 작업 폴더에 `작업\` 과 `고객마스터.xlsx`(없을 때만)만 만든다. 받은문서 · 정리문서 · 확인필요 · 출력 · 규칙.yaml 은 만들지 않는다.
  - `outlook N` → 그 작업 입력 폴더로 첨부 저장
  - `run.bat` = `uv run python -m app %*` 그대로(바로가기의 실행기). 인자 없이 누르면 안내만.

- [ ] **Step 1: 실패하는 테스트**

```python
def test_job_new_then_list(tmp_path, capsys)            # 출력에 "01_일일 PO 정리" 와 "16:00"
def test_no_args_shows_jobs_not_inbox(tmp_path, capsys) # "받은문서" 글자 없음
def test_init_makes_jobs_folder_and_master_only(tmp_path)
def test_check_with_job_shows_verify_table(tmp_path, capsys)
def test_run_auto_writes_record_and_notifies(tmp_path, monkeypatch)  # notify 를 가짜로 → 호출 1회
def test_pick_folder_rejects_sync_folder(tmp_path, monkeypatch, capsys)
def test_old_inbox_flow_is_gone()                        # config 에 INBOX 없음, 수집.bat 없음
```

- [ ] **Step 2~4:** 실패 확인 → 구현 → `uv run pytest -q`
- [ ] **Step 5: 커밋** — `W6 작업 명령과 옛 흐름 정리`

---

### Task W7: 매일 실행 등록 · 바로가기 · 놓친 실행 · 알림

**Files:**
- Create: `app/shortcuts.py`, `app/schedule.py`, `app/notify.py`
- Modify: `app/__main__.py` (job new 끝에 등록 + 바로가기, delete/pause/resume, catch-up)
- Test: `tests/test_schedule.py`, `tests/test_shortcuts.py`, `tests/test_notify.py`

**Interfaces:**
- Consumes: W1 `Job`, `list_jobs`, `save_job`; W6 명령
- Produces:
  - `shortcuts.desktop_dir(run=subprocess.run) -> Path` (PowerShell `[Environment]::GetFolderPath('Desktop')`, OneDrive 로 옮겨진 바탕화면도 따름), `startup_dir(run=...) -> Path` (`'Startup'`)
  - `shortcuts.create(path: Path, target: Path, args: str, workdir: Path, run=subprocess.run) -> None` (WScript.Shell COM 을 PowerShell 로, `-EncodedCommand` 로 한글 경로 안전)
  - `shortcuts.job_shortcuts(job, repo_root, desktop: Path, run=...) -> list[Path]` → `지금 실행 - {제목}.lnk` (`run.bat run NN`), `승인 - {제목}.lnk` (`run.bat approve NN`)
  - `schedule.TASK_PREFIX = "사무자동화"`, 작업 이름 `사무자동화\{NN}_{제목}` 
  - `schedule.register(job, repo_root, run=subprocess.run, startup: Path | None = None) -> str` — 반환 `"작업스케줄러"` · `"시작프로그램"` · `"없음"`. 1) `Register-ScheduledTask`(매일 job.run_time, `-StartWhenAvailable`, 현재 사용자 · 로그인 시에만, 동작 = `{repo}\.venv\Scripts\pythonw.exe -m app run NN --자동`, 시작 폴더 = repo) 후 `Get-ScheduledTask` 로 다시 조회해 있어야 성공. 2) 실패면 시작프로그램 폴더에 `사무자동화 놓친 작업 실행.lnk`(`pythonw -m app catch-up`) 하나(이미 있으면 그대로). 3) 그것도 실패면 `"없음"`. 결과를 `job.scheduler` 에 저장.
  - `schedule.unregister(job, run=...) -> None`, `pause` · `resume`
  - `schedule.due_jobs(jobs, now: datetime) -> list[Job]` — paused 아님 · `last_run != 오늘` · `run_time <= now` · 번호 순
  - `schedule.catch_up(work, now=None, runner=None) -> list[str]` — due_jobs 를 차례로 실행(동시에 아님). 잠긴 작업은 건너뜀
  - `notify.py` (새 파일): `report_text(job, report) -> str` (예: `"일일 PO 정리: 결과 3건, 확인필요 1건"`, 0인 항목은 뺌, 어느 작업에도 맞지 않는 파일 N개 포함), `notify_report(job, report, toast=None, box=None) -> None` — `has_news` 가 False 면 아무것도 안 함. 기본 toast = PowerShell 의 WinRT `Windows.UI.Notifications.ToastNotificationManager`(외부 모듈 없음), 클릭 = 출력 폴더 열기(결과 0건이면 작업 폴더). toast 가 실패하면 `ctypes.windll.user32.MessageBoxW`. 잠금을 푼 뒤 호출
- 작업 스케줄러 · 시작프로그램 바로가기는 사용자가 지우거나 PC 를 옮겨도 `job new` 를 다시 돌리지 않고 복구되도록 `job repair` 명령(모든 작업 다시 등록 · 바로가기 다시 만듦)도 둔다.

- [ ] **Step 1: 실패하는 테스트** (`run` 에 가짜 실행기를 넣어 PowerShell 명령 글자와 반환 코드를 흉내)

```python
def test_register_uses_task_scheduler_when_query_finds_task()   # 반환 "작업스케줄러", 명령에 "-StartWhenAvailable" 과 "16:20"
def test_register_falls_back_to_startup_when_blocked(tmp_path)  # 등록 명령 실패 → 반환 "시작프로그램", startup 폴더에 .lnk 생성 요청 1회
def test_register_none_when_everything_blocked()                 # 반환 "없음"
def test_job_shortcuts_names_and_args(tmp_path)                  # "지금 실행 - 일일 PO 정리.lnk" args "run 01"
def test_due_jobs_skips_done_paused_and_future()
def test_catch_up_runs_missed_jobs_in_order_once(tmp_path)       # 16:00·16:20 지났고 17:00 안 지남 → 1,2 차례로, 다시 부르면 0개
def test_job_new_prints_which_method_was_used(tmp_path, capsys)  # "자동 실행: Windows 작업 스케줄러 (매일 16:00)" 또는 대체 문구
def test_report_text_skips_zero_counts()                         # counts 성공 3, 확인필요 1 → "일일 PO 정리: 결과 3건, 확인필요 1건"
def test_no_notification_when_nothing_new()                      # has_news False → toast · box 둘 다 호출 0
def test_toast_failure_falls_back_to_message_box()
```

- [ ] **Step 2~4:** 실패 확인 → 구현 → `uv run pytest -q`
- [ ] **Step 5: 실제 PC 한 번 확인(테스트 아님)** — 임시 작업으로 `job new` → 작업 스케줄러에 `사무자동화\01_…` 이 보이고 바탕화면 바로가기 2개가 생기는지 본 뒤 `job delete 1` 로 지운다. 결과를 커밋 메시지 본문에 적는다.
- [ ] **Step 6: 커밋** — `W7 매일 실행 등록과 바로가기`

---

### Task W8: 작업지시서 · 약속 · 문서 · 설치

**Files:**
- Create: `작업지시서.md`, `rules/규칙.뼈대.yaml`(W6 에서 이미 만들었으면 수정만)
- Modify: `CLAUDE.md`, `README.md`, `01_구조_이해.md`, `03_프로그램_설치.md`, `04_처음_처리하기.md` → `04_작업_만들기.md`(이름 변경), `05_내_양식_가르치기.md`, `06_매일_쓰기.md`, `07_프로젝트와_예정됨.md`, `08_선택_Outlook_자동수집.md`, `99_막혔을_때.md`, `install.ps1`(끝 안내 문구 · `$LegacyFiles` 에 `04_처음_처리하기.md` · `수집.bat`), `tests/test_docs_links.py`, `docs/superpowers/specs/2026-10-28-office-automation-kr-design.md` 4절 목록
- Test: `tests/test_docs_links.py`, `tests/test_work_order_doc.py`

**작업지시서.md 내용(Claude 가 따르는 대화 순서, 설계 4절 그대로):**
1. 「무엇을 만들어 드릴까요? 제목을 입력해 주세요.」
2. 「입력할 파일(엑셀 또는 PDF)의 예를 하나 넣어 주세요.」 — 이 파일 하나만 연다. 원본이 Claude 에게 간다는 안내 한 줄을 먼저 보여 준다.
3. 「출력은 회사 양식, 누적장부, 둘 다 중 무엇으로 할까요?」 — 회사 양식이면 그 양식 파일도 받는다.
4. `uv run python -m app pick-folder "입력 폴더를 고르세요"` → 5. 같은 방식으로 출력 폴더
6. `job new …` → 그 작업 폴더의 규칙.yaml 을 채움 → `check 예시파일 --작업 NN` 을 비고 · 검산 불일치가 0이 될 때까지. 고객마스터가 필요하면 친구에게 「고객마스터.xlsx 에 이 고객 줄을 직접 넣어 주세요」(Claude 는 열지 않음).
7. 등록 결과 한 줄(어느 방법 · 몇 시) + 「바탕화면에 「지금 실행 - 제목」, 「승인 - 제목」이 생겼습니다. 처음 10건은 승인대기로 갑니다.」
8. 「다음 작업도 만드실까요, 그만하실까요?」
- 시작 문장: 친구가 Code 탭에 `작업지시서대로 새 작업 만들어 줘` 를 넣는다(CLAUDE.md 가 이 문장 = 작업지시서.md 를 따르라는 뜻이라고 적음).

**CLAUDE.md 바뀌는 점:** 작업 만들기 때 친구가 준 예시 파일(원본) 1개와 회사 양식 파일만 연다. 그 외 작업 폴더 파일 · 입력 · 출력 폴더 · 고객마스터 · 누적장부 · 승인대기는 열지 않는다. `run` · `approve` · `catch-up` 은 실행하지 않는다(친구가 바탕화면으로). 고칠 것은 그 작업의 규칙.yaml 뿐.

**문서:** 01장 데이터 경계 표에 「작업 만들 때 원본 1개가 Claude 에게 감」. 04장 = 작업 만들기(시작 문장 · 질문 순서 · 끝나면 바탕화면). 05장 = 규칙 고치기 · 「○○ 작업에 이 고객 양식 더해 줘」(새 양식은 다시 10건 승인). 06장 = 매일 4시 · 지금 실행 · 결과 알림 · 승인(틀린 건 승인대기 복사본을 지움) · 확인필요 · 원본은 입력 폴더에 그대로(이미 처리한 것은 건너뜀, 정리는 친구가 원할 때) · 「어느 작업에도 맞지 않는 파일」. 07장 = 왜 Claude 예정됨이 아니라 Windows 작업인지. 99장 = 자동 실행이 「시작프로그램」 · 「없음」으로 나올 때, IT 담당자에게 물어볼 문장, 검산 불일치 사유 읽는 법.

- [ ] **Step 1: 실패하는 테스트**

```python
def test_work_order_has_all_questions_in_order():   # 작업지시서.md 에 설계 4절 질문 1~8 문구가 이 순서로
def test_claude_md_points_to_work_order():          # "작업지시서대로 새 작업 만들어 줘" 와 "작업지시서.md"
def test_docs_mention_original_goes_to_claude():    # 01장 · 04장에 "원본" 과 "Anthropic"
# test_docs_links.py PLANNED 목록을 04_작업_만들기.md 로
```

- [ ] **Step 2~4:** 실패 확인 → 작성 → `uv run pytest -q`
- [ ] **Step 5: 커밋** — `W8 작업지시서와 문서`
- [ ] **Step 6: push 전 확인** — 대표에게 `git push origin main` 여부를 묻고, 승인 뒤에만 push.
