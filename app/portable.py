"""무설치판(PowerShell 5.1)용 규칙 내보내기. 규칙.yaml 을 지금 엔진(app.rules)으로 읽고 검사한 뒤 JSON 으로 옮긴다.

무설치판은 YAML 을 읽지 못하므로 대표 PC 에서 이 JSON 을 만들어 넘긴다. 무설치판이 지원하지 않는 규칙은 여기서 거절한다.
"""
import json
from pathlib import Path

from app import config
from app.job_run import MASTER_PREFIX, rules_fingerprint
from app.jobs import JobError, find_job
from app.rules import RuleError, Rules, load_rules

VERSION = 1
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
