"""무설치판(PowerShell 5.1)용 규칙 내보내기. 규칙.yaml 을 지금 엔진(app.rules)으로 읽고 검사한 뒤 JSON 으로 옮긴다.

무설치판은 YAML 을 읽지 못하므로 대표 PC 에서 이 JSON 을 만들어 넘긴다. 무설치판이 지원하지 않는 규칙은 여기서 거절한다.
"""
import dataclasses
import json
import os
import shutil
import sys
from datetime import date
from decimal import Decimal
from pathlib import Path

from app import config, local_guard
from app.excel_extract import read_excel
from app.extract import extract
from app.job_run import MASTER_PREFIX, rules_fingerprint
from app.jobs import JobError, _norm, find_job, list_jobs
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


def _job_rules(job) -> dict:
    """작업의 규칙.yaml → 무설치판 규칙 사전(+ fingerprint = 지금 판 rules_fingerprint). 지원하지 않으면 RuleError."""
    data = rules_to_portable(load_rules(job.rules_path))
    data["fingerprint"] = rules_fingerprint(job)
    return data


def _write_json(path: Path, data) -> None:
    """BOM 없는 utf-8 JSON. 늘 새로 쓴다(준비는 같은 번호면 규칙.json 수정 시각이 새것을 고른다)."""
    path.write_text(json.dumps(data, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")


def export_rules(job_number: int, out: Path, work: Path | None = None) -> Path:
    """작업 번호의 규칙.yaml 을 JSON 으로 내보낸다. 검사를 통과한 뒤에만 파일을 쓴다(work 가 없으면 작업폴더.txt 의 폴더)."""
    work = config.work_dir() if work is None else work
    data = _job_rules(find_job(work, str(job_number)))
    out = Path(out)
    out.parent.mkdir(parents=True, exist_ok=True)
    _write_json(out, data)
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
        _write_json(folder / name, data)
    return folder


# --- 묶음(bundle-portable): 친구 PC 로 통째로 복사할 폴더 ---

PORTABLE_DIR = "무설치판"          # 저장소의 무설치판 코드 폴더 = 묶음의 무설치판\
BUNDLE_JOBS_DIR = "작업"           # 묶음\작업\NN_제목\규칙.json · 작업.json (무설치판 준비가 읽음)
BUNDLE_JOB_RULES = "규칙.json"
BUNDLE_JOB_INFO = "작업.json"      # {number, title, output_mode}
BUNDLE_DIRS = (PORTABLE_DIR, BUNDLE_JOBS_DIR, SELF_TEST_DIR)  # make_bundle 이 만드는 폴더(이것만 묶음에 있다)
PORTABLE_CODE_SUFFIXES = (".ps1", ".bat")  # 무설치판\ 에서 옮기는 파일(코드만. 다른 것은 묻어 가지 않게)
ONEDRIVE_ENV = ("OneDrive", "OneDriveCommercial", "OneDriveConsumer")  # Windows 가 알려 주는 OneDrive 폴더
BUNDLE_HINT = "예: C:\\무설치판묶음 처럼 내 PC 로컬 디스크의 비어 있는 새 폴더를 --out 으로 주세요."


class BundleError(Exception):
    """묶음을 만들 수 없는 경우(묶음 폴더 자리 · 묶을 작업). 메시지는 사용자에게 그대로 보인다."""


def _inside(path, root) -> bool:
    """path 가 root 이거나 그 안이면 True."""
    path, root = _norm(path), _norm(root)
    return path == root or path.startswith(root.rstrip(os.sep) + os.sep)


def check_bundle_out(out, repo_root=config.REPO_ROOT) -> None:
    """묶음 폴더 자리 검사: 저장소 안 · 동기화 폴더(OneDrive 등) · 네트워크 · 지난 묶음이 있으면 BundleError.
    동기화 · 네트워크 판정은 작업 폴더 가드(local_guard)와 같은 기준에 OneDrive 환경 변수 경로를 더한 것."""
    out = Path(out)
    text = str(out)
    if text.startswith("\\\\"):
        raise BundleError(f"묶음 폴더가 네트워크 위치입니다: {text}\n{BUNDLE_HINT}")
    if _inside(out, repo_root):
        raise BundleError(f"묶음 폴더가 저장소 안입니다: {text}\n저장소에는 묶음을 만들지 않습니다. {BUNDLE_HINT}")
    for part in out.parts:
        if any(part.lower().startswith(name) for name in local_guard.SYNC_FOLDER_NAMES):
            raise BundleError(f"묶음 폴더가 동기화 폴더({part}) 안에 있습니다: {text}\n"
                              f"클라우드로 자동 업로드되므로 쓸 수 없습니다. {BUNDLE_HINT}")
    for name in ONEDRIVE_ENV:
        synced = os.environ.get(name, "").strip()
        if synced and _inside(out, synced):
            raise BundleError(f"묶음 폴더가 OneDrive 폴더({synced}) 안에 있습니다: {text}\n"
                              f"클라우드로 자동 업로드되므로 쓸 수 없습니다. {BUNDLE_HINT}")
    if sys.platform == "win32" and out.drive and local_guard._drive_type(out.drive + "\\") == local_guard.DRIVE_REMOTE:
        raise BundleError(f"묶음 폴더가 네트워크 드라이브({out.drive})에 있습니다: {text}\n{BUNDLE_HINT}")
    if out.exists() and not out.is_dir():
        raise BundleError(f"묶음 폴더 자리에 파일이 있습니다: {text}\n{BUNDLE_HINT}")
    stale = [name for name in BUNDLE_DIRS if (out / name).exists()]
    if stale:
        raise BundleError(f"묶음 폴더에 이미 {' · '.join(stale)} 가 있습니다: {text}\n"
                          f"지난 묶음이나 작업 폴더와 섞이지 않게 새 폴더를 쓰세요. {BUNDLE_HINT}")


def _portable_code(repo_root: Path) -> list[Path]:
    """저장소 무설치판\\ 의 .ps1 · .bat. 묶음에 필요한 저장소 파일(코드 · 자체 시험 샘플 · 예시 규칙)이 없으면 BundleError."""
    source = repo_root / PORTABLE_DIR
    files = [p for p in source.rglob("*") if p.is_file() and p.suffix.lower() in PORTABLE_CODE_SUFFIXES]
    if not files:
        raise BundleError(f"무설치판 코드가 없습니다: {source}")
    for needed in (SELF_TEST_PDF, SELF_TEST_XLSX, SELF_TEST_RULES_SOURCE):
        if not (repo_root / needed).is_file():
            raise BundleError(f"자체 시험 자료가 저장소에 없습니다: {repo_root / needed}")
    return files


def make_bundle(work: Path, numbers: list[int], out: Path, repo_root=config.REPO_ROOT) -> Path:
    """out 에 친구 PC 로 넘길 묶음을 만들고 out 을 돌려준다:
    무설치판\\(저장소 코드) + 작업\\NN_제목\\규칙.json · 작업.json + 자체시험\\(build_self_test).
    작업의 규칙 · 설정만 읽는다(고객마스터 · 마크업표 · 입력 · 출력 · 기록은 보지도 넣지도 않는다).
    자리 검사와 모든 작업의 규칙 검사를 통과한 뒤에만 쓴다."""
    repo_root, out = Path(repo_root), Path(out)
    check_bundle_out(out, repo_root)
    jobs = [find_job(work, str(n)) for n in dict.fromkeys(numbers)]
    if not jobs:
        raise BundleError("묶을 작업이 없습니다. 먼저 작업을 만드세요.")
    payload = []  # 하나라도 지원하지 않으면 여기서 RuleError(아무것도 쓰지 않음)
    for job in jobs:
        try:
            payload.append((job, _job_rules(job)))
        except RuleError as e:
            raise RuleError(f"작업 {job.name}: {e}") from e
    code = _portable_code(repo_root)

    out.mkdir(parents=True, exist_ok=True)
    for path in code:  # 바이트 그대로(BOM · CRLF 유지). 수정 시각은 상관없다
        copy = out / PORTABLE_DIR / path.relative_to(repo_root / PORTABLE_DIR)
        copy.parent.mkdir(parents=True, exist_ok=True)
        shutil.copyfile(path, copy)
    for job, rules in payload:
        folder = out / BUNDLE_JOBS_DIR / job.name
        folder.mkdir(parents=True)
        _write_json(folder / BUNDLE_JOB_RULES, rules)
        _write_json(folder / BUNDLE_JOB_INFO, {"number": job.number, "title": job.title, "output_mode": job.output_mode})
    build_self_test(out, repo_root)
    return out


def bundle_numbers(work: Path, keys: list[str] | None) -> list[int]:
    """명령줄 --작업 NN … 을 번호 목록으로. 없으면 멈추지 않은 모든 작업(멈춘 작업 = 만드는 중이거나 쉬는 작업.
    무설치판에는 멈춤이 없어 넘기면 매일 돌므로, 넘기려면 --작업 으로 직접 고른다)."""
    if keys:
        return [job_number_from(k) for k in keys]
    numbers = [job.number for job in list_jobs(work) if not job.paused]
    if not numbers:
        raise BundleError("묶을 작업이 없습니다(멈추지 않은 작업이 없음). 멈춘 작업을 넘기려면 --작업 NN 으로 고르세요.")
    return numbers
