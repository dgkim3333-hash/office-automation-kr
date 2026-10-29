"""무설치판(PowerShell 5.1)용 규칙 내보내기. 규칙.yaml 을 지금 엔진(app.rules)으로 읽고 검사한 뒤 JSON 으로 옮긴다.

무설치판은 YAML 을 읽지 못하므로 대표 PC 에서 이 JSON 을 만들어 넘긴다. 무설치판이 지원하지 않는 규칙은 여기서 거절한다.
"""
import dataclasses
import json
import shutil
from datetime import date
from decimal import Decimal
from pathlib import Path

from app import config
from app.excel_extract import read_excel
from app.extract import extract
from app.job_run import MASTER_PREFIX, rules_fingerprint
from app.jobs import JobError, find_job
from app.rules import RuleError, Rules, load_rules

VERSION = 1
# 자체 시험(친구 PC 의 준비가 그 PC 의 워드 · 엑셀로 읽어 정답과 견줌): 묶음의 자체시험 폴더
SELF_TEST_DIR = "자체시험"
SELF_TEST_RULES = "규칙.json"
SELF_TEST_ANSWER = "정답.json"
# 저장소 samples 의 가짜 문서(make_samples.py 가 만든 것). PDF 는 워드가 「Bill To / Ship To」 를 문단 여러 개가 든 표 칸으로
# 바꾸는 문서라, 워드 표 행을 높이로 묶는 방법이 이 PC 와 다르면 고객명 등이 달라져 자체 시험이 잡는다.
SELF_TEST_PDF = Path("samples") / "03_Invoice_DM_PDF" / "INV-2026-10231.pdf"
SELF_TEST_XLSX = Path("samples") / "01_고객오더_엑셀" / "고객오더_ACME_PO-ACME-2026-0417.xlsx"
SELF_TEST_RULES_SOURCE = Path("rules") / "규칙.예시.yaml"
# 파이썬에만 있거나 .NET 과 뜻이 다른 정규식 문법(이름 붙은 묶음 · 이름 역참조 · 글자 끝 \Z)
PYTHON_ONLY_REGEX = ("(?P<", "(?P=", "\\Z")


def _check_regex(where: str, pattern: str):
    for token in PYTHON_ONLY_REGEX:
        if token in pattern:
            raise RuleError(
                f"{where} 정규식에 무설치판(.NET)과 문법이 다른 '{token}' 가 있습니다. "
                f"{token} 를 쓰지 않는 패턴으로 바꾸세요: {pattern}")


def _doc_type(doc) -> dict:
    for f in doc.fields.values():
        _check_regex(f"[{doc.name}] 필드 '{f.name}'", f.pattern.pattern)
    table = doc.table
    if table is not None:
        _check_regex(f"[{doc.name}] 표 행패턴", table.row_pattern.pattern)
    check = doc.check
    return {
        "name": doc.name,
        "keywords": list(doc.keywords),
        "number_field": doc.number_field,
        "fields": [{"name": f.name, "pattern": f.pattern.pattern, "format": f.fmt} for f in doc.fields.values()],
        "table": None if table is None else {
            "start": table.start,
            "end": table.end,
            "row_pattern": table.row_pattern.pattern,
            "columns": list(table.columns),
            "amount_columns": list(table.amount_columns),
            "quantity_columns": list(table.quantity_columns),
        },
        "check": None if check is None else {
            "subtotal": check.subtotal,
            "extras": list(check.extras),
            "total": check.total,
            "quantity_column": check.quantity,
            "price_column": check.price,
            "amount_column": check.amount,
        },
        "filename": doc.filename_tpl,
    }


def _excel_form(form) -> dict:
    items = form.items
    return {
        "name": form.name,
        "doc_type": form.doc_type,
        "keywords": list(form.keywords),
        "head": [{"name": h.name, "where": h.where, "format": h.fmt} for h in form.head.values()],
        "items": None if items is None else {
            "header": items.header_text,
            "columns": dict(items.columns),
            "amount_columns": list(items.amount_columns),
            "quantity_columns": list(items.quantity_columns),
        },
    }


def _company_form(form) -> dict | None:
    if form is None:
        return None
    for col in form.columns:
        if col.source.startswith(MASTER_PREFIX):
            raise RuleError(
                f"[회사양식] 열 '{col.name}' 의 값 '{col.source}' 는 무설치판이 지원하지 않습니다"
                f"({MASTER_PREFIX}* 는 고객마스터 파일을 읽어야 해서 무설치판에는 없습니다). "
                "그 열은 머리.이름 · 고정:글자 · 빈칸 중 하나로 바꾸세요.")
    return {
        "filename": form.filename_tpl,
        "sheet": form.sheet,
        "columns": [{"name": c.name, "value": c.source} for c in form.columns],
        "markup_columns": {"part": form.markup.part, "price": form.markup.price, "quantity": form.markup.quantity},
    }


def rules_to_portable(rules: Rules) -> dict:
    """해석된 규칙을 무설치판이 읽는 사전으로 옮긴다. 지원하지 않는 규칙이면 RuleError."""
    return {
        "version": VERSION,
        "sheet_order": list(rules.sheet_order),
        "doc_types": [_doc_type(d) for d in rules.doc_types],
        "excel_forms": [_excel_form(f) for f in rules.excel_forms],
        "company_form": _company_form(rules.company_form),
    }


def export_rules(job_number: int, out: Path, work: Path | None = None) -> Path:
    """작업 번호의 규칙.yaml 을 JSON 으로 내보낸다. 검사를 통과한 뒤에만 파일을 쓴다(work 가 없으면 작업폴더.txt 의 폴더)."""
    work = config.work_dir() if work is None else work
    job = find_job(work, str(job_number))
    data = rules_to_portable(load_rules(job.rules_path))
    data["fingerprint"] = rules_fingerprint(job)
    out = Path(out)
    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_text(json.dumps(data, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    return out


def job_number_from(key: str) -> int:
    """명령줄의 작업 번호('01')를 정수로. 숫자가 아니면 JobError."""
    if not key.strip().isdigit():
        raise JobError(f"작업 번호를 숫자로 적어 주세요(예: 01): {key}")
    return int(key)


# --- 자체 시험 자료 ---

def _plain(value):
    """엔진 JSON 표기: 금액은 자리 그대로의 글자, 날짜는 ISO, 나머지는 그대로(무설치판 ConvertTo-JsonValue 와 같음)."""
    if isinstance(value, Decimal):
        return str(value)
    if isinstance(value, date):
        return value.isoformat()
    return value


def result_view(result) -> dict:
    """추출 결과를 무설치판 엔진 결과 JSON 모양으로(doc_type · form · is_excel · fields · table · warnings · review_reason)."""
    return {
        "doc_type": result.doc_type.name if result.doc_type else None,
        "form": result.source_form,
        "is_excel": result.is_excel,
        "fields": {k: _plain(v) for k, v in result.fields.items()},
        "table": None if result.table is None else [{k: _plain(v) for k, v in row.items()} for row in result.table],
        "warnings": list(result.warnings),
        "review_reason": result.review_reason,
    }


def build_self_test(out_dir, repo_root=config.REPO_ROOT) -> Path:
    """out_dir/자체시험/ 에 가짜 PDF 1 · 가짜 엑셀 1 · 규칙.json · 정답.json 을 만들고 그 폴더를 돌려준다.
    정답은 지금 판 엔진(PDF = pdfplumber, 엑셀 = openpyxl)의 추출 결과. 규칙은 예시 규칙에서 회사양식을 뺀 것(추출만 견줌).
    정답.json = {"version": 1, "pdf": {"file", "result"}, "excel": {"file", "result"}}."""
    repo_root = Path(repo_root)
    folder = Path(out_dir) / SELF_TEST_DIR
    folder.mkdir(parents=True, exist_ok=True)
    rules = dataclasses.replace(load_rules(repo_root / SELF_TEST_RULES_SOURCE), company_form=None)
    answer = {"version": VERSION}
    for key, source, read in (("pdf", SELF_TEST_PDF, extract), ("excel", SELF_TEST_XLSX, read_excel)):
        copy = folder / source.name
        shutil.copyfile(repo_root / source, copy)
        answer[key] = {"file": source.name, "result": result_view(read(copy, rules))}
    for name, data in ((SELF_TEST_RULES, rules_to_portable(rules)), (SELF_TEST_ANSWER, answer)):
        (folder / name).write_text(json.dumps(data, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    return folder
