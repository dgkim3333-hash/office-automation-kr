# 무설치판 구현 계획

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 친구 PC 에 아무것도 설치하지 않고 Windows PowerShell 5.1 + 오피스(엑셀 · 워드)만으로, 지금 판과 같은 결과(회사 양식 · 누적장부 · 검산 · 마크업)를 내는 「무설치판」을 만든다.

**Architecture:** 세 층으로 나눈다. ① **엔진**(순수 PowerShell, 오피스 없이 돎): 규칙 JSON + 글자(PDF) 또는 칸 격자(엑셀)를 받아 추출 · 검산 · 마크업 · 회사 양식 줄을 만든다. ② **오피스 가장자리**(COM): 엑셀로 칸 격자 읽기, 워드로 PDF 글자 읽기, 엑셀로 결과 · 장부 쓰기와 되읽기. ③ **실행**: 작업 폴더 · 처리기록 · 알림 · 예약 · 준비.bat. 규칙은 대표 PC 의 지금 프로그램이 `규칙.yaml` 을 해석해 JSON 으로 내보낸다. 엔진은 pytest 가 같은 입력을 지금 판(Python)과 무설치판(PowerShell)에 넣어 결과가 같은지로 검증한다.

**Tech Stack:** Windows PowerShell 5.1(`powershell.exe`), Excel/Word COM(Office 16), 지금 판 Python 3.13 + uv + pytest(내보내기 · 묶음 · 비교 시험).

**Spec:** `docs/superpowers/specs/2026-10-29-no-install-edition-design.md`(이하 「설계」). 지금 판 동작의 기준은 `app\` 코드와 `docs/superpowers/specs/2026-10-28-work-orders-design.md`.

## Global Constraints

- 친구 PC 쪽 코드(`무설치판\`)는 **Windows PowerShell 5.1** 문법만. PowerShell 7 전용(`&&` · `||` · `?:` · `??` · `-Parallel` · `ConvertFrom-Json -AsHashtable` 등) 금지. 시험은 늘 `powershell.exe`(5.1)로 돌린다.
- `무설치판\` 의 모든 `.ps1` · `.bat` 은 UTF-8 **BOM** 으로 저장(5.1 이 한글을 읽게). `.bat` 첫 줄 `@echo off` 다음 `chcp 65001 > nul`.
- 금액 계산은 `[decimal]` 만. `[double]` 로 돈 계산 금지. 숫자 · 날짜 변환은 `[Globalization.CultureInfo]::InvariantCulture`.
- 엔진 JSON 출력에서 금액은 **문자열**(`"1234.50"` — 원본 자리 수 유지), 날짜는 `"YYYY-MM-DD"`, 정수 수량은 숫자. `ConvertTo-Json` 은 늘 `-Depth 20`.
- 친구 PC 의 설정 · 레지스트리 · 보안을 바꾸지 않는다(읽기만). 인터넷 연결 · 내려받기 코드 금지(`Invoke-WebRequest` · `Invoke-RestMethod` · `Net.WebClient` · `http://` · `https://` 가 `무설치판\` 에 없어야 함).
- 끝내는 오피스 프로세스는 **이 실행이 COM 으로 띄운 것(명령줄에 `/Automation`, 실행 전에 없던 PID)만**.
- 원본(입력 폴더 · 마크업표)은 옮기거나 고치지 않는다. 오피스로 열 때는 늘 실행마다 새로 만든 임시 폴더의 복사본을 연다.
- 지원하지 않는 규칙(`고객마스터.*`)은 내보내기에서 거절. 승인 10건 · 고객마스터 · Outlook 은 만들지 않는다(설계 2절).
- 문서 · 화면 문구 · 커밋 메시지는 한국어. 회사명 · 실제 고객 자료 금지, 샘플은 가짜(ACME 류). 저장소 안에 PDF · xlsx 를 새로 만들지 않는다(시험은 `tmp_path` 에 만든다).
- 커밋 메시지 첫 줄 = 「작업번호 제목」(예: `3 엑셀 격자 추출`). 작업마다 테스트 먼저(실패 확인) → 구현 → 통과 → 커밋 1회. 끝나면 PowerShell 에서 전체 `uv run pytest` 녹색.
- 오피스가 필요한 시험은 `@pytest.mark.office`(엑셀 · 워드 COM 이 없으면 건너뜀), `C:\PO시험` 이 필요한 시험은 `@pytest.mark.po_test_kit`(폴더 없으면 건너뜀). `powershell.exe` 가 없으면 무설치판 시험 전체 건너뜀.

## Review Focus

1. **친구 PC 지역 설정이 다름**(소수점 쉼표 · 날짜 순서가 다른 Windows): 같은 결과여야 한다 → 작업 2 · 4 에 `de-DE` 문화권으로 돌리는 시험.
2. **워드 변환이 표 줄을 합치거나 쪼갬**: 시작 · 끝키워드 사이에 행패턴과 맞지 않는 줄이 있으면 품목을 조용히 빼지 않고 확인필요 → 작업 2 시험.
3. **친구가 장부 · 결과 · 마크업표를 엑셀로 열어 둔 채 실행**: 장부 · 결과는 보류, 마크업표는 복사본으로 읽힘 → 작업 9 시험.
4. **같은 PO 가 PDF 와 엑셀로 두 번 옴(같은 번호)**: 두 번째는 중복, 장부 한 줄 → 작업 9 시험.
5. **한글 · 띄어쓰기 · 긴 파일 이름**(`메일_Order LDH-PO-2611-033.pdf` 류): 읽기 · 복사본 · 확인필요 .txt 이름이 깨지지 않음 → 작업 7 · 9 시험.

---

## 파일 구조

```
app\portable.py                 (새) 규칙 → 무설치판 JSON 내보내기, 묶음 만들기
app\__main__.py                 (고침) export-portable · bundle-portable 명령
무설치판\
  무설치판.ps1                  진입점: -Command 실행|점검|준비|엔진 … (dispatcher)
  lib\Rules.ps1                 규칙 JSON 읽기 · 정규식 컴파일
  lib\Values.ps1                금액 · 날짜 · 자리 수 · 파일명 채우기 (지금 판 to_amount · to_date · render_text · free_name)
  lib\ExtractText.ps1           PDF 글자에서 판별 · 필드 · 표
  lib\ExtractGrid.ps1           엑셀 격자에서 판별 · 머리 · 품목
  lib\Verify.ps1                원문 대조 · 자리 수 · 줄 검산 · 합계 검산
  lib\Markup.ps1                마크업표 격자 해석 · 적용 · 검산
  lib\Company.ps1               회사 양식 줄
  lib\Office.ps1                COM: 격자 읽기 · 워드 글자 읽기(정규화) · 시간 제한 · 정리
  lib\Writers.ps1               COM: 회사 양식 · 장부 · 검산 파일 쓰기 · 되읽기
  lib\Run.ps1                   작업 실행(처리기록 · 중복 · 확인필요 · 재시도 · 잠금 · 보고)
  lib\Notify.ps1                Windows 알림 · 메시지 창
  lib\Setup.ps1                 준비: 자체 시험 · 폴더 고르기 · 예약 · 바로가기
  준비.bat · 지금실행.bat(작업 번호 인자)
tests\portable\                 (새) conftest.py(마커 · ps 실행 도우미) · test_*.py
09_무설치판.md                  (새) 문서 장
```

`무설치판.ps1 -Command 엔진 -Input <json 파일> -Output <json 파일>` 은 시험 전용 입구다: 입력 JSON 의 `op` 로 엔진 함수 하나를 부르고 결과를 JSON 으로 쓴다. 이후 작업의 엔진 시험은 모두 이 입구를 쓴다.

---

### Task 1: 규칙 내보내기 (Python)

**Files:**
- Create: `app\portable.py`, `tests\portable\__init__.py`, `tests\portable\test_export.py`
- Modify: `app\__main__.py`(명령 `export-portable NN --out FILE`)

**Interfaces:**
- Produces: `rules_to_portable(rules: Rules) -> dict`, `export_rules(job_number: int, out: Path) -> Path`. JSON 모양(이후 모든 작업이 읽음):
  ```
  {"version": 1,
   "sheet_order": [...],
   "doc_types": [{"name", "keywords": [...], "number_field",
                  "fields": [{"name", "pattern", "format": null|"날짜"|"금액"}],
                  "table": null | {"start", "end", "row_pattern", "columns": [...], "amount_columns": [...], "quantity_columns": [...]},
                  "check": null | {"subtotal", "extras": [...], "total", "quantity_column", "price_column", "amount_column"},
                  "filename"}],
   "excel_forms": [{"name", "doc_type", "keywords": [...],
                    "head": [{"name", "where", "format"}],   // where = "B4" 또는 "오른쪽:글자"
                    "items": null | {"header", "columns": {"이름": "A"}, "amount_columns", "quantity_columns"}}],
   "company_form": null | {"filename", "sheet", "columns": [{"name", "value"}],
                           "markup_columns": {"part", "price", "quantity"}},
   "fingerprint": "<규칙.yaml 지문, job_run.rules_fingerprint 와 같은 값>"}
  ```
  키 이름은 `app\rules.py` 데이터클래스 속성을 따른다(설계자가 실제 속성 이름에 맞춰 고쳐도 되지만, 고친 이름을 이 블록에 남긴다).

- [ ] **Step 1: 실패하는 시험 작성** — `tests\portable\test_export.py`
  - `test_example_rules_export_round_trip`: `rules\규칙.예시.yaml` 에서 `고객마스터.*` 열을 뺀 규칙으로 `rules_to_portable` → `json.dumps` 가 되고, 문서종류 이름 순서가 `["DM", "Invoice", "PO"]`, 회사양식 열 이름 목록이 원래 순서와 같다.
  - `test_master_columns_rejected`: `고객마스터.고객코드` 가 든 규칙 → `RuleError`, 메시지에 `고객마스터` 와 `무설치판` 이 들어 있다.
  - `test_python_only_regex_rejected`: 패턴에 `(?P<x>` · `(?P=` · `\Z` 가 있으면 `RuleError`(메시지에 그 글자). `.NET` 과 다른 문법을 막는 목록이다.
  - `test_markup_columns_exported`: `마크업.단가` 열이 있으면 `company_form.markup_columns` 가 `{"part": "품번", "price": "단가", "quantity": "수량"}`(기본값).
  - `test_cli_export_writes_json(tmp_path)`: 작업 하나를 만든 작업 폴더에서 `export-portable 01 --out tmp_path/규칙.json` → 파일이 생기고 `fingerprint` 가 `job_run.rules_fingerprint(job)` 와 같다.
- [ ] **Step 2: 실패 확인** — `uv run pytest tests/portable/test_export.py -v` → `ModuleNotFoundError: app.portable`
- [ ] **Step 3: 구현** — `app\portable.py` 에 위 두 함수. 규칙 해석은 `app.rules` 를 그대로 거친 뒤 데이터클래스를 사전으로 옮긴다(정규식은 `.pattern` 문자열). `__main__` 에 명령 추가(지금 판 명령 등록 방식 그대로).
- [ ] **Step 4: 통과 확인** — 같은 명령 → PASS
- [ ] **Step 5: 커밋** — `1 규칙 내보내기`

### Task 2: PowerShell 바탕 · 엔진(글자 추출)

**Files:**
- Create: `무설치판\무설치판.ps1`, `무설치판\lib\Rules.ps1`, `무설치판\lib\Values.ps1`, `무설치판\lib\ExtractText.ps1`, `tests\portable\conftest.py`, `tests\portable\test_engine_text.py`, `tests\portable\test_ps_files.py`
- Modify: `pyproject.toml`(마커 `office`, `po_test_kit` 등록)

**Interfaces:**
- Consumes: 작업 1 JSON.
- Produces:
  - 시험 도우미 `run_engine(op: str, payload: dict, culture: str | None = None) -> dict` (`tests\portable\conftest.py`): 입력 JSON 을 `tmp_path` 에 쓰고 `powershell.exe -NoProfile -ExecutionPolicy Bypass -File 무설치판\무설치판.ps1 -Command 엔진 -Input … -Output …` 를 돌려 결과 JSON 을 돌려준다. `culture` 가 있으면 그 문화권으로 스레드를 바꾼 뒤 부른다(`-Culture` 인자).
  - PS 함수: `Import-PortableRules([string]$Path) -> object`(정규식은 `[regex]` 로 컴파일해 붙임. 컴파일 실패는 오류 문구 `규칙 정규식 오류: <이름>`), `ConvertTo-Amount([string]) -> [decimal] 또는 $null`, `ConvertTo-DateValue([string]) -> [datetime] 또는 $null`(형식 `%m/%d/%Y` · `%Y-%m-%d` 만), `Get-TextExtract($Rules, [string]$Text) -> 결과 객체`.
  - 결과 객체(엔진 공통, 이후 작업이 씀): `{doc_type, form, is_excel, fields: {이름: 값}, table: null | [ {열: 값} ], warnings: [...], review_reason: null | "사유"}`. 사유 문구는 `app\extract.py` `ExtractResult.review_reason` 과 글자까지 같다.
  - 엔진 op: `extract_text` (`{rules, text}` → 결과 객체).

- [ ] **Step 1: 실패하는 시험 작성**
  - `test_engine_text.py::test_same_as_python_on_samples`: `samples\03_Invoice_DM_PDF` 의 PDF 6개 각각, `extract.read_text` 로 얻은 글자를 Python `extract_from_text` 와 엔진 `extract_text` 에 넣으면 `doc_type` · `fields` · `table` · `review_reason` 이 같다(금액은 `str(Decimal)` 비교, 날짜는 ISO). 스캔 PDF(`Scan_signed_packing_list.pdf`)는 `review_reason == "텍스트 없음 (스캔본, OCR 범위 밖)"`.
  - `test_unmatched_row_goes_to_review`: 표 사이에 행패턴과 안 맞는 줄 하나를 끼운 글자 → `table is None`, 경고에 `표 행패턴과 맞지 않는 줄이 있어 품목 건너뜀` (Review Focus 2).
  - `test_culture_de_de_same_result`: 위 샘플 하나를 `culture="de-DE"` 로 → 기본 문화권 결과와 같다 (Review Focus 1).
  - `test_quantity_must_be_integer`: 수량 `1.5` 인 줄 → 그 줄 불일치(지금 판과 같이 표 전체 None).
  - `test_ps_files.py::test_bom_and_no_network`: `무설치판\` 의 모든 `.ps1` · `.bat` 이 BOM 으로 시작하고, `Invoke-WebRequest` · `Invoke-RestMethod` · `WebClient` · `http://` · `https://` 가 없다.
  - `test_ps_files.py::test_no_ps7_syntax`: 모든 `.ps1` 을 `powershell.exe` 의 `[System.Management.Automation.Language.Parser]::ParseFile` 로 읽어 오류 0개.
- [ ] **Step 2: 실패 확인** — `uv run pytest tests/portable -v` → 무설치판 파일 없음으로 FAIL
- [ ] **Step 3: 구현** — 동작 기준은 `app\extract.py`(`detect_type` 은 `first_page` 없이 전체 글자, 대소문자 무시 포함 검색 · `HEAD_LINES = 30` · 필드는 첫 일치의 1번 괄호 · 표는 시작키워드 줄 다음부터 끝키워드 줄 앞까지, 빈 줄 건너뜀, `row_pattern` 은 줄 처음부터 일치(`match` = .NET 에서 `^` 고정), 금액열 · 수량열 변환). 줄 나누기는 `\r?\n`.
- [ ] **Step 4: 통과 확인** — `uv run pytest tests/portable -v` → PASS
- [ ] **Step 5: 커밋** — `2 무설치판 바탕과 PDF 글자 추출`

### Task 3: 엔진(엑셀 격자 추출)

**Files:**
- Create: `무설치판\lib\ExtractGrid.ps1`, `tests\portable\grid.py`(openpyxl → 격자 JSON), `tests\portable\test_engine_grid.py`

**Interfaces:**
- Produces:
  - 격자 JSON(작업 7 의 엑셀 읽기도 같은 모양을 낸다): `{"sheets": [{"name", "cells": [[행, 열, {"t": "n"|"s"|"b"|"d", "v": 값, "f": "표시 형식"}], …]}]}` — 빈 칸은 싣지 않는다. 날짜 칸은 `t="d"`, `v="YYYY-MM-DDTHH:MM:SS"`. 숫자는 JSON 숫자(엑셀 Value2 와 같은 double).
  - `tests\portable\grid.py::grid_from_xlsx(path) -> dict`(openpyxl `data_only=True` 값 · `number_format`; `datetime` → `d`).
  - PS `Get-GridExtract($Rules, $Grid) -> 결과 객체`(`is_excel=$true`, `form` = 엑셀양식 이름). 엔진 op `extract_grid`.
- [ ] **Step 1: 실패하는 시험** — `test_same_as_python_on_samples`: `samples\01_고객오더_엑셀` 의 xlsx 2개를 `rules\규칙.예시.yaml`(고객마스터 열 뺀 것)로 → Python `read_excel` 과 엔진 결과가 같다(`form` · `fields` · `table` · `review_reason`). `test_no_form_matches`: 판별키워드 없는 격자 → 사유가 `read_excel` 과 같은 글자. `test_float_cell_places`: 단가 칸 `212.3375`(double) → 결과 문자열 `"212.3375"`(지금 판 `excel_number` 과 같은 자리 수 규칙, `EXCEL_DIGITS = 15`).
- [ ] **Step 2: 실패 확인** — `uv run pytest tests/portable/test_engine_grid.py -v` → FAIL
- [ ] **Step 3: 구현** — 기준 `app\excel_extract.py`(`DETECT_ROWS = 20`, `오른쪽:글자` 찾기, 품목은 머리글 줄 아래부터 품번 칸이 빈 줄 앞까지, `excel_number`).
- [ ] **Step 4: 통과 확인** → PASS
- [ ] **Step 5: 커밋** — `3 엑셀 격자 추출`

### Task 4: 엔진(검산)

**Files:**
- Create: `무설치판\lib\Verify.ps1`, `tests\portable\test_engine_verify.py`

**Interfaces:**
- Produces: PS `Get-Checks($Result, $Source) -> [ {item, source_value, result_value, ok} ]`. `$Source` = `{kind: "text", text}` 또는 `{kind: "grid", grid}`. 엔진 op `verify`(`{rules, result, source}` — `result` 는 작업 2 · 3 의 결과 객체). 항목 이름 · 값 글자는 `app\verify.py` `Check` 와 같다.
- [ ] **Step 1: 실패하는 시험**
  - `test_same_as_python_on_samples`: 작업 2 · 3 의 샘플 전부 → `verify(result, path)` 와 엔진 검산 목록이 같다(순서 포함, `ok` 포함). PDF 는 엔진에 pdfplumber 글자를 준다.
  - 바꾼 값 시험(각각 Python 과 엔진이 같은 `ok=False` 항목): 자리 다른 숫자(`1,234.50` → `12,345.0`), 소수점 다른 숫자, 표기만 다른 같은 값(`1,234.50` ↔ `1234.5` → 통과), 줄 금액 틀림, 소계 틀림, 부대비용 없는 합계, 엑셀 0.1 류 부동소수.
  - `test_culture_de_de_same_result` (Review Focus 1).
- [ ] **Step 2: 실패 확인** → FAIL
- [ ] **Step 3: 구현** — 기준 `app\verify.py` 전부(`NUMBER` · `DATE` 토큰, `EXCEL_MIN_PLACES = 2`, `_rounds_to`, `_total_checks`). 정규식 두 개는 .NET 에서 같은 뜻인지 시험으로 확인.
- [ ] **Step 4: 통과 확인** → PASS
- [ ] **Step 5: 커밋** — `4 검산`

### Task 5: 엔진(마크업)

**Files:**
- Create: `무설치판\lib\Markup.ps1`, `tests\portable\test_engine_markup.py`

**Interfaces:**
- Produces: PS `Read-MarkupGrid($Grid) -> @{ 열쇠 = @{part; rate; places} }`(오류는 `app\markup.py` `MarkupError` 와 같은 문구로 throw), `Get-MarkupValue([string]$Source, $Item, $Columns, $Table)`, `Get-MissingReason($Result, $Form, $Table) -> 사유 또는 $null`, `Get-MarkupChecks($Result, $RowsWritten, $Table, $Columns) -> 검산 목록`, `Get-MarkupFingerprint($Table, $Parts) -> string`(그 문서 파트들의 마크업표 줄이 바뀔 때만 바뀌는 값. 지금 판과 같은 값일 필요는 없다). 엔진 op `markup`.
- [ ] **Step 1: 실패하는 시험** — 마크업표 격자는 `tests\portable\grid.py` 로 `tmp_path` 의 openpyxl 파일에서 만든다(값 + `number_format`).
  - 같은 결과(Python `markup_from_bytes` · `markup_value` · `missing_reason` · `markup_checks` 와 엔진): `3` → 3%, 백분율 서식 `0.005` → 0.5%, 글자 `"0.5%"`, 반올림자리 2 → `ROUND_HALF_UP`(`12.345` → `12.35`), 반올림자리 빈칸 → 반올림 없음, 파트 숫자 칸 `12345.0` ↔ 품번 `"12345"`, 대소문자 · 앞뒤 공백 무시.
  - 오류 같은 문구: 같은 파트 두 줄, 마크업% 상한(`MAX_RATE = 10` 기준 지금 판 동작), 반올림자리 상한(`MAX_PLACES = 10`), 숫자 아닌 %.
  - `test_missing_part_reason`: 표에 없는 파트 → 사유에 그 파트넘버, 0% 로 넘기지 않음.
- [ ] **Step 2: 실패 확인** → FAIL
- [ ] **Step 3: 구현** — 기준 `app\markup.py` 전부.
- [ ] **Step 4: 통과 확인** → PASS
- [ ] **Step 5: 커밋** — `5 마크업`

### Task 6: 엔진(회사 양식 줄 · 파일 이름)

**Files:**
- Create: `무설치판\lib\Company.ps1`, `tests\portable\test_engine_company.py`
- Modify: `무설치판\lib\Values.ps1`(`Format-Template` · `Get-FreeName`)

**Interfaces:**
- Produces: PS `Get-CompanyRows($Result, $Form, $MarkupTable) -> @{ header = [...]; rows = [[...]] }`(값 출처: `머리.X` · `품목.X` · `줄번호`(10, 20, 30 …) · `수량x단가` · `고정:` · `빈칸` · `마크업.단가/금액/율`), `Format-Template([string]$Template, $Fields) -> string`(지금 판 `render_text`: `{필드:%Y-%m}` 날짜 서식, 금지 글자 `\/:*?"<>|` → `_`, 빈 값 `미상`), `Get-FreeName([string]$Path, $Reserved) -> string`(`_2`, `_3` …). 엔진 op `company`.
- [ ] **Step 1: 실패하는 시험** — 샘플 엑셀 2개와 `C:\PO시험` 3개(`po_test_kit`, `04_정답\참고_검증된_규칙.yaml` + `03_고객마스터_마크업\마크업표.xlsx`; 이 규칙에 고객마스터 열이 있으면 그 열을 뺀 사본으로)로 Python `company_rows` 와 엔진 결과가 칸마다 같다. `Format-Template` 은 `render_text` 와 같은 글자(날짜 서식 · 금지 글자 · 미상).
- [ ] **Step 2: 실패 확인** → FAIL
- [ ] **Step 3: 구현** — 기준 `app\outputs.py` `_value` · `company_rows`, `app\organize.py`.
- [ ] **Step 4: 통과 확인** → PASS
- [ ] **Step 5: 커밋** — `6 회사 양식 줄과 파일 이름`

### Task 7: 오피스 읽기 (엑셀 격자 · 워드 글자)

**Files:**
- Create: `무설치판\lib\Office.ps1`, `tests\portable\test_office_read.py`

**Interfaces:**
- Produces:
  - `Read-ExcelGrid([string]$Path, [int]$TimeoutSec) -> 격자`(작업 3 모양; 날짜는 `.Value` 가 `[datetime]` 인 칸, 숫자는 `.Value2`, 형식은 `.NumberFormat`).
  - `Read-WordText([string]$Path, [int]$TimeoutSec) -> @{ text; pages }`. 정규화: 문단마다 `ListFormat.ListString` 이 있으면 앞에 붙이고 한 칸, 탭 · 연속 공백 → 한 칸, 줄 앞뒤 공백 제거, 가로줄 흔적(`/` 한 글자 줄) 제거, 워드 표는 행마다 칸을 한 칸 띄어 한 줄로.
  - 둘 다: 실행마다 `%TEMP%\무설치판_<작업>_<날짜시각>_<임의>\` 에 복사본을 만들어 연다(복사본의 `Zone.Identifier` 스트림은 지움 — 원본은 그대로), 끝나면 임시 폴더를 지운다. 시간 제한을 넘으면 `OfficeTimeout` 예외(문구 `워드가 응답하지 않음` / `엑셀이 응답하지 않음`).
  - `Select-OwnOfficePids($Before, $After) -> [int[]]`: 실행 전 PID 목록과 실행 뒤 목록(각각 `{pid, commandline}`)에서 새로 생겼고 명령줄에 `/Automation` 이 있는 것만. `Stop-OwnOffice` 는 이것만 끝내고 사라질 때까지 기다린다(최대 10초).
  - `Test-PdfWarningOff() -> bool`: `HKCU:\Software\Microsoft\Office\<설치된 버전>\Word\Options\DisableConvertPdfWarning` 이 1 인지(버전은 `HKCR\Word.Application\CurVer` 에서). 읽기만.
  - 시간 제한 기본값: 파일당 120초(`$OfficeTimeoutSec`).
- [ ] **Step 1: 실패하는 시험**
  - 오피스 없이: `test_select_own_pids`: before `[1,2(/Automation)]`, after `[1,2(/Automation),3(/Automation),4]` → `[3]`.
  - `@office` `test_word_text_same_engine_result`: 샘플 PDF 5개(스캔 제외) + `po_test_kit` PDF 2개 → `Read-WordText` 글자로 낸 엔진 결과(작업 2)가 pdfplumber 글자로 낸 결과와 같다. 줄번호가 목록 번호인 PDF 에서 줄번호 열이 `1, 2, 3 …` 으로 나온다.
  - `@office` `test_excel_grid_same_as_openpyxl`: 샘플 xlsx 2개 → `Read-ExcelGrid` 와 `grid_from_xlsx` 가 칸마다 같다(숫자 · 글자 · 날짜, 형식 문자열은 비교하지 않고 엔진 결과로 비교).
  - `@office` `test_korean_name_and_original_untouched`: 이름이 `메일_Order 시험-PO 1.pdf` 인 복사본을 `tmp_path` 에 두고 읽기 → 성공, 원본 바이트 · 수정 시각 그대로, 임시 폴더 남지 않음 (Review Focus 5).
  - `@office` `test_timeout_cleans_up`: 시간 제한 0초 → `OfficeTimeout`, 시험 전에 없던 `/Automation` 워드 0개.
  - `@office` `test_reopen_after_crash`: 같은 파일을 두 번 연달아 `Read-WordText`(첫 번째는 시간 제한으로 끊음) → 두 번째가 성공(「지난번 오류」 창에 막히지 않음).
- [ ] **Step 2: 실패 확인** → FAIL
- [ ] **Step 3: 구현** — COM 호출은 별도 runspace/잡에서 돌리고 바깥에서 `Wait-Job -Timeout` 으로 제한한다. `Documents.Open(path, ConfirmConversions=$false, ReadOnly=$true, AddToRecentFiles=$false)`, `Workbooks.Open(path, UpdateLinks=0, ReadOnly=$true)`, `Visible=$false`, `DisplayAlerts` 끔. 닫기 · `Quit` · `ReleaseComObject` 뒤 `Stop-OwnOffice`.
- [ ] **Step 4: 통과 확인** — `uv run pytest tests/portable/test_office_read.py -v` → PASS(오피스 있는 이 PC 에서 건너뜀 0개여야 함)
- [ ] **Step 5: 커밋** — `7 오피스로 읽기`

### Task 8: 오피스 쓰기 (회사 양식 · 장부 · 검산 파일)

**Files:**
- Create: `무설치판\lib\Writers.ps1`, `tests\portable\test_office_write.py`

**Interfaces:**
- Consumes: 작업 4 검산 목록, 작업 6 회사 양식 줄.
- Produces: `Write-CompanyForm([string]$Folder, $Result, $Company, $Form) -> 경로`, `Add-LedgerEntry([string]$LedgerPath, $Rules, $Result, [string]$SourceName, [datetime]$When)`, `Add-LedgerLog(...)`, `Test-LedgerDuplicate($LedgerPath, $DocType, $Number) -> bool`, `Write-CheckFile([string]$Folder, $Result, $Checks, [string]$SourceName, $Company) -> 경로`, `Test-ReadBack([string]$Path, [string]$Sheet, $Header, $Rows) -> [문제 목록]`, `Test-FileLocked([string]$Path) -> bool`.
  - 장부 모양은 `app\ledger.py` 와 같다: 문서종류 시트(`DOC_HEAD = 처리일시 · 원본파일` + 필드 + `DOC_TAIL = 정리된경로`) · `품목`(`문서종류 · 문서번호` + 표 열) · `처리로그`(`일시 · 파일 · 결과 · 사유`), 시트 순서 = `sheet_order`, 표 이름 `Ledger_<시트>`, 형식 `yyyy-mm-dd hh:mm:ss` · `yyyy-mm-dd` · 금액 자리 수 서식(`amount_format`), 저장 전 `.bak.xlsx`, 사용자가 만든 시트 · 열은 건드리지 않음.
  - 회사 양식 · 검산 파일 모양은 `app\outputs.py` 와 같다(`CHECK_HEADER`, 시트 이름 `추출값` · `품목` · `회사양식`).
- [ ] **Step 1: 실패하는 시험(`@office`)**
  - `test_company_form_same_as_python`: 샘플 엑셀 하나 → Python `write_company_form` 파일과 PS 파일을 openpyxl 로 읽어 칸 값 · 숫자 형식이 같다.
  - `test_ledger_same_as_python`: 샘플 PDF 2개 + 엑셀 1개를 차례로 장부에 넣기 → Python `Ledger` 결과와 시트 이름 · 순서 · 머리 · 값 · 형식 · 표 이름이 같다. 두 번째 저장 때 `.bak.xlsx` 가 생긴다.
  - `test_user_sheet_kept`: 사용자가 장부에 만든 시트 `메모` 와 문서 시트 끝의 사용자 열이 남는다(지금 판 시험 `test_ledger.py` 의 같은 경우를 따름).
  - `test_read_back_detects_change`: 저장 뒤 칸 하나를 바꾸면 `Test-ReadBack` 이 그 칸을 알려 준다.
  - `test_locked_ledger`: openpyxl 로 연 것이 아니라 엑셀 COM 으로 장부를 연 채 `Test-FileLocked` → `$true`.
- [ ] **Step 2: 실패 확인** → FAIL
- [ ] **Step 3: 구현** — 엑셀 COM(`Visible=$false`, 작업 7 의 시간 제한 · 정리 재사용). 표는 `ListObjects`.
- [ ] **Step 4: 통과 확인** → PASS
- [ ] **Step 5: 커밋** — `8 오피스로 쓰기`

### Task 9: 작업 실행

**Files:**
- Create: `무설치판\lib\Run.ps1`, `무설치판\lib\Notify.ps1`, `무설치판\지금실행.bat`, `tests\portable\test_run.py`
- Modify: `무설치판\무설치판.ps1`(`-Command 실행 -Job NN`)

**Interfaces:**
- Consumes: 작업 1~8 전부.
- Produces: `Invoke-PortableJob([string]$Work, [int]$Number, [datetime]$Now) -> 보고 객체 {counts: {성공, 중복, 확인필요, 보류}, ledger_locked, stopped, alerts}`.
  - 친구 PC 작업 폴더 `<작업폴더>\작업\NN_제목\`: `작업설정.json`(`{number, title, input_dir, output_dir, output_mode: "회사양식"|"누적장부"|"둘다", time: "16:00", method: "작업스케줄러"|"시작프로그램"|"없음", last_run}`), `규칙.json`, `처리기록.csv`(열 = `지문 · 파일 · 문서종류 · 번호 · 상태 · 규칙지문 · 날짜 · 마크업지문`, 지문 = 파일 내용 SHA-256), `처리로그.csv`, `실행기록.txt`, `확인필요\원본이름.txt`, `검산\`, `실행중.lock`. 마크업표는 `<작업폴더>\마크업표.xlsx`. 장부는 출력 폴더 `누적장부_<제목>.xlsx`.
  - 동작 기준은 `app\job_run.py` 에서 승인 · 고객마스터를 뺀 것: 처리 순서, `SETTLE_SECONDS = 60`, 이미 성공한 지문 건너뜀, 확인필요 재시도(날짜가 바뀜 · 규칙지문이 바뀜 · 마크업지문이 바뀜), 같은 문서종류 · 번호는 `중복`, 장부 · 결과가 열려 있으면 그 실행 `보류`(문구 `LEDGER_LOCKED_REASON` · `HELD_REASON` 과 같게), 잠금 파일로 겹쳐 돌지 않음(`이 작업이 이미 실행 중입니다`), 확인필요 .txt 문구는 지금 판과 같음(`RETRY_NOTE` · `KEEP_HINT`). 알림 문구는 `app\notify.py` `report_text` 와 같은 모양.
- [ ] **Step 1: 실패하는 시험(`@office`, 모두 `tmp_path` 작업 폴더)**
  - `test_same_outputs_as_python`: 입력 폴더에 샘플 PDF 4개 + 엑셀 2개(파일 시각을 2분 전으로) → 무설치판 실행 결과(회사 양식 파일들 · 장부)가 지금 판 `run_job(job, work, gate=통과)` 결과와 칸마다 같다. 보고 숫자도 같다.
  - `@po_test_kit` `test_po_test_kit_matches_answer`: `C:\PO시험\01_입력예시` 3개를 `tmp_path` 입력 폴더로 복사 → 회사 양식 3개 · 장부가 `04_정답` 의 정답표와 같다(`03_…\마크업표.xlsx` 복사본 사용).
  - `test_rerun_skips_and_renamed_skips`: 두 번째 실행 → 성공 0, 파일 이름을 바꿔도 0.
  - `test_settling_file_waits`: 방금 쓴 파일 → 처리 안 함, 다음 실행(시각 +2분) 때 처리.
  - `test_same_po_pdf_and_excel_is_duplicate`: 같은 번호의 PDF · 엑셀 → 하나 성공 · 하나 중복, 장부 한 줄 (Review Focus 4).
  - `test_ledger_open_holds`: 장부를 엑셀 COM 으로 연 채 실행 → 보류, 장부 무변화. 마크업표를 연 채 실행 → 정상 처리 (Review Focus 3).
  - `test_bad_total_goes_to_review_and_retries_after_rule_change`: 합계가 틀린 문서 → 확인필요 .txt, 결과 · 장부 무변화. 같은 날 다시 → 그대로, `규칙.json` 의 fingerprint 를 바꾸면 다시 시도.
  - `test_missing_markup_part_retries_after_markup_change`: 표에 없는 파트 → 확인필요, 마크업표에 그 파트를 더하면 같은 날 다시 시도해 성공.
  - `test_word_timeout_only_that_doc`: 시간 제한을 아주 짧게 한 PDF 하나 + 엑셀 하나 → PDF 는 확인필요(`워드가 응답하지 않음`), 엑셀은 성공.
  - `test_lock_prevents_overlap`: `실행중.lock` 을 잡은 채 실행 → `이 작업이 이미 실행 중입니다`.
- [ ] **Step 2: 실패 확인** → FAIL
- [ ] **Step 3: 구현** — 알림은 `app\notify.py` 의 `toast_script` 와 같은 PowerShell(WinRT 알림, 꺼져 있으면 메시지 창). 시험에서는 `-NoNotify` 로 끈다.
- [ ] **Step 4: 통과 확인** → PASS
- [ ] **Step 5: 커밋** — `9 작업 실행`

### Task 10: 준비 · 예약 · 자체 시험

**Files:**
- Create: `무설치판\lib\Setup.ps1`, `무설치판\준비.bat`, `tests\portable\test_setup.py`
- Modify: `무설치판\무설치판.ps1`(`-Command 준비`, `-Command 놓친실행`, `-Command 점검 -File …`)

**Interfaces:**
- Consumes: 작업 7 `Read-WordText` · `Test-PdfWarningOff`, 작업 9 `Invoke-PortableJob`.
- Produces:
  - `Invoke-SelfTest([string]$BundleDir) -> @{ ok; pdf_ok; excel_ok; message }`: 묶음의 `자체시험\`(가짜 PDF 1 · 엑셀 1 · `정답.json`)을 읽어 엔진 결과가 `정답.json` 과 같은지. PDF 가 다르면 `pdf_ok=$false`, 메시지 `이 PC 의 워드는 PDF 를 다르게 읽습니다. PDF 작업은 켜지 않았습니다`.
  - `Register-PortableSchedule($Job, [switch]$WhatIf) -> "작업스케줄러"|"시작프로그램"|"없음"`: `Register-ScheduledTask`(현재 사용자, 로그인했을 때만, StartWhenAvailable, 시각 16:00 + 20분 × (번호 − 1)) → 등록 뒤 다시 조회로 확인 → 실패하면 시작프로그램 바로가기(`-Command 놓친실행`) → 실패하면 없음. `-WhatIf` 는 만들 명령 글자만 돌려준다.
  - `New-RunShortcut($Job, [string]$Desktop)`: 「지금 실행 - 제목」(대상 `지금실행.bat NN`).
  - `Invoke-Catchup([string]$Work, [datetime]$Now)`: 오늘 정해진 시각이 지났고 `last_run` 이 오늘이 아닌 작업만 실행.
  - `Invoke-Setup([string]$BundleDir, [string]$Work)`: 오피스 · PowerShell 버전 확인 → 자체 시험 → 작업마다(이미 있으면 규칙만 바꾸고 폴더는 묻지 않음, 처리기록 · 장부 유지) 폴더 고르기(`FolderBrowserDialog`, 로컬 디스크만 — OneDrive · 네트워크 · UNC 거부, 기준은 `app\local_guard.check_work_dir`) → 예약 · 바로가기 → `마크업표.xlsx` 가 없으면 가짜 예시 → `Test-PdfWarningOff` 가 거짓이면 「워드에서 PDF 를 한 번 열고 '이 메시지를 다시 표시 안 함'에 체크」 안내. 폴더 고르기 창 · 메시지 창은 시험에서 끼워 넣기(`-FolderPicker`, `-NoDialog`).
  - `-Command 점검 -File <경로> -Rules <규칙.json>`: 대표 PC 에서 쓰는 점검. 지금 판 `check` 와 같은 결과 표(항목 · 값 · 검산 · 비고)를 콘솔에 찍는다.
- [ ] **Step 1: 실패하는 시험**
  - `test_schedule_whatif_time`: 작업 3 → 명령 글자에 `16:40`, 현재 사용자, `StartWhenAvailable`.
  - `test_schedule_falls_back`: 예약 등록 함수를 실패하게 끼우면 `시작프로그램`, 둘 다 실패하면 `없음`.
  - `test_catchup_runs_only_due`: 15:00 → 없음, 17:00 · last_run 어제 → 실행, last_run 오늘 → 없음.
  - `test_setup_update_keeps_records`: 같은 묶음으로 준비 두 번 → 두 번째는 폴더를 묻지 않고, 처리기록 · 장부 그대로, 규칙.json 만 새 것.
  - `test_rejects_onedrive_folder`: 고른 폴더가 `OneDrive` 아래 → 거부 문구.
  - `@office` `test_self_test_passes_here`: 작업 11 이 만든 묶음으로 → `ok=$true`.
  - `@office` `test_check_command_matches_python_check`: 샘플 PDF 하나 → 콘솔 표의 값 · 검산 칸이 `uv run python -m app check` 출력과 같다.
- [ ] **Step 2: 실패 확인** → FAIL
- [ ] **Step 3: 구현** — `준비.bat` = `powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0무설치판.ps1" -Command 준비`. 바로가기는 `WScript.Shell`.
- [ ] **Step 4: 통과 확인** → PASS
- [ ] **Step 5: 커밋** — `10 준비와 예약`

### Task 11: 묶음 만들기 (Python)

**Files:**
- Modify: `app\portable.py`, `app\__main__.py`(명령 `bundle-portable --out DIR [--작업 NN …]`), `tests\portable\test_export.py`

**Interfaces:**
- Produces: `make_bundle(work: Path, numbers: list[int], out: Path) -> Path`. 묶음 = `out\무설치판\`(저장소 `무설치판\` 복사) + `out\작업\NN_제목\규칙.json` · `작업.json`(`{number, title, output_mode}`) + `out\자체시험\`(저장소 `samples` 의 가짜 PDF 하나 · 엑셀 하나 + 지금 판 엔진으로 만든 `정답.json` · 그 규칙.json). `out` 이 저장소 안이거나 OneDrive · 네트워크면 거부.
- [ ] **Step 1: 실패하는 시험** — `test_bundle_layout(tmp_path)`: 묶음 파일 목록이 위와 같고, `정답.json` 이 작업 2 · 3 엔진 결과 모양, 묶음에 `고객마스터` · `마크업표` · 입력 폴더 파일이 없다. `test_bundle_refuses_repo_and_onedrive`.
- [ ] **Step 2: 실패 확인** → FAIL
- [ ] **Step 3: 구현**
- [ ] **Step 4: 통과 확인** → PASS
- [ ] **Step 5: 커밋** — `11 묶음 만들기`

### Task 12: 문서

**Files:**
- Create: `09_무설치판.md`
- Modify: `README.md`(장 목록에 09), `99_막혔을_때.md`(무설치판 절: 스크립트 실행이 막힘 · 워드가 PDF 를 다르게 읽음 · 워드가 응답하지 않음, 「IT 담당자에게 물어볼 문장」에 `PowerShell 스크립트(-ExecutionPolicy Bypass) 실행` 추가), `작업지시서.md`(끝에 「무설치판으로 넘길 작업」: 대표 PC 기준 폴더, `export-portable` · 무설치판 점검 둘 다 통과, `bundle-portable`), `CLAUDE.md`(저장소용: 무설치판 묶음을 만들 때도 고객마스터 · 마크업표를 열지 않는다)
- Test: 기존 `tests\test_docs_links.py` · `tests\test_work_order_doc.py` 가 새 장을 포함해 통과, 새 `tests\portable\test_docs.py`

**Interfaces:** 없음.

- [ ] **Step 1: 실패하는 시험** — `test_docs.py`: `09_무설치판.md` 에 「원본 1개가 친구 회사 밖(대표)으로 나가고 Claude 에게도 한 번 간다」 문장과 「다시 표시 안 함」 안내, `준비.bat` 이 들어 있고, 「Code 탭이면 밖으로 안 나간다」 류 문장이 없다. README 장 목록에 09.
- [ ] **Step 2: 실패 확인** → FAIL
- [ ] **Step 3: 작성** — 09장 절: 언제 쓰나(설치가 막혔을 때) · 대표 PC 에서 할 일 · 친구 PC 에서 할 일(묶음 복사 → 준비.bat → 매일) · 업데이트(묶음 덮어쓰기 → 준비.bat 다시) · 지금 판과 다른 점(승인 10건 · 고객마스터 · Outlook 없음) · 막혔을 때 → 99장. 비개발자 말투, 기존 장 모양(「이 장이 끝나면 보이는 것」).
- [ ] **Step 4: 통과 확인** — `uv run pytest tests/test_docs_links.py tests/test_work_order_doc.py tests/portable/test_docs.py -v` → PASS
- [ ] **Step 5: 커밋** — `12 무설치판 문서`

### 마지막: 전체 확인

- [ ] PowerShell 에서 `uv run pytest` → 전체 녹색, `office` · `po_test_kit` 시험이 이 PC 에서 건너뜀 0개.
- [ ] `uv run python -m app bundle-portable --out C:\무설치판_시험묶음 --작업 01` 로 만든 묶음을 `C:\PO시험` 과 다른 빈 작업 폴더에 두고 `준비.bat` → 「지금 실행」 한 번(대표가 화면으로 확인).
