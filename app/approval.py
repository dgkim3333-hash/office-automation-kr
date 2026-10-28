"""처음 10건 승인: 새 양식은 처음 10건을 사람이 승인한 뒤부터 결과를 바로 낸다.

승인 수는 양식 단위로 센다(엑셀 PO 는 엑셀양식 이름, PDF 는 문서종류 이름). 규칙.yaml 에서 그 양식에
해당하는 부분이 바뀌면 그 양식만 0 으로 되돌린다. 승인할 문서는 작업 폴더의 승인대기\\ 에 있는 원본 복사본이고,
그 대조표는 승인대기\\대조표\\ 에 있다. 승인하지 않을 문서는 친구가 그 복사본을 지운다.
"""
import csv
import hashlib
from pathlib import Path

import openpyxl
import yaml

from app.job_run import (
    INPUT_SUFFIXES, WAITING, WAITING_CHECKS, RunReport, _current_state, file_fingerprint, load_records, run_job,
)
from app.jobs import Job, save_job

REQUIRED = 10
NOTHING = "승인할 문서가 없습니다."
FINISHED = ("성공", "중복")  # 이 상태가 된 복사본은 더 승인할 것이 없으므로 치운다


def form_key(result) -> str:
    """승인 수를 세는 양식 이름. 엑셀이면 엑셀양식 이름, 아니면 문서종류 이름."""
    if result.source_form:
        return str(result.source_form)
    return result.doc_type.name


def _named(section, key: str):
    if not isinstance(section, dict):
        return None
    for name, body in section.items():
        if str(name) == key:
            return body
    return None


def form_fingerprint(rules_path: Path, key: str) -> str:
    """규칙.yaml 중 그 양식에 해당하는 부분 + 회사양식의 sha256 앞 16자.

    엑셀양식이면 그 엑셀양식과 연결된 문서종류, 아니면 그 문서종류. 다른 양식의 규칙을 고쳐도 값이 바뀌지 않는다.
    """
    data = yaml.safe_load(Path(rules_path).read_text(encoding="utf-8")) or {}
    part = {"회사양식": data.get("회사양식")}
    excel = _named(data.get("엑셀양식"), key)
    if excel is not None:
        part["엑셀양식"] = excel
        doc = excel.get("문서종류") if isinstance(excel, dict) else None
        part["문서종류"] = _named(data.get("문서종류"), str(doc))
    else:
        part["문서종류"] = _named(data.get("문서종류"), key)
    text = yaml.safe_dump(part, sort_keys=True)
    return hashlib.sha256(text.encode("utf-8")).hexdigest()[:16]


def _current_count(job: Job, key: str) -> int:
    """그 양식의 승인 수. 규칙이 저장된 지문과 다르면 그 양식만 0 으로 되돌리고 새 지문을 저장한다."""
    fingerprint = form_fingerprint(job.rules_path, key)
    if job.fingerprints.get(key) != fingerprint:
        job.approvals[key] = 0
        job.fingerprints[key] = fingerprint
        save_job(job)
    return job.approvals.get(key, 0)


def approval_gate(job: Job, rules, result) -> bool:
    """True 면 결과를 내지 않고 승인대기로 보낸다(그 양식의 승인 수가 10 미만)."""
    return _current_count(job, form_key(result)) < REQUIRED


def _documents(waiting: Path) -> list[Path]:
    if not waiting.is_dir():
        return []
    return sorted(p for p in waiting.iterdir()
                  if p.is_file() and p.suffix.lower() in INPUT_SUFFIXES and not p.name.startswith("~$"))


def approve(job: Job, work) -> RunReport:
    """승인대기의 복사본을 지금 규칙으로 다시 추출 · 검산해 결과까지 낸다.

    성공한 건마다 그 양식 승인 수 +1. 끝나면 성공 · 중복이 된 복사본과, 짝이 되는 복사본이 없는 대조표를 지운다.
    검산에 실패한 복사본은 승인대기에 남는다(확인필요 사유를 보고 고친 뒤 다시 승인하거나, 복사본을 지운다).
    """
    waiting = Path(job.folder) / WAITING
    warnings: list[str] = []
    if not _documents(waiting):
        _clear_orphan_checks(waiting, warnings)
        return RunReport(message="\n".join([NOTHING] + warnings))

    approved: list[str] = []

    def counted(job_, result):
        key = form_key(result)
        _current_count(job_, key)  # 문서가 승인대기에 있는 사이 규칙이 바뀌었으면 지금 규칙으로 센다
        job_.approvals[key] = job_.approvals.get(key, 0) + 1
        approved.append(key)
        save_job(job_)

    report = run_job(job, work, source_dir=waiting, gate=lambda *a: False, on_success=counted)
    _clear_finished(job, waiting, warnings)
    _clear_orphan_checks(waiting, warnings)

    lines = [report.message] if report.message else []
    if approved:
        counts = []
        for key in dict.fromkeys(approved):
            n = job.approvals.get(key, 0)
            counts.append(f"{key} {n}/{REQUIRED}" + (" (다음부터 바로 결과를 냅니다)" if n >= REQUIRED else ""))
        lines.append(f"승인 {len(approved)}건. 양식별 승인 수: " + ", ".join(counts))
    report.message = "\n".join(lines + warnings)
    return report


def _clear_finished(job: Job, waiting: Path, warnings: list[str]):
    """처리기록에서 성공 · 중복이 된 복사본을 지운다(처리기록은 줄을 더하기만 하므로 다시 지워도 같다)."""
    try:
        state = _current_state(load_records(job))
    except (OSError, UnicodeDecodeError, csv.Error) as e:
        warnings.append(f"경고: 처리기록을 읽지 못해 승인한 복사본을 치우지 못함 ({e}). 다음 승인 때 다시 치웁니다")
        return
    for path in _documents(waiting):
        try:
            if state.get(file_fingerprint(path), {}).get("상태") in FINISHED:
                path.unlink()
        except OSError as e:
            warnings.append(f"경고: 승인대기의 복사본을 지우지 못함 ({path.name}): {e}. 다음 승인 때 다시 치웁니다")


def _source_name(check: Path) -> str | None:
    """대조표 맨 위 줄의 원본파일(승인대기 복사본 이름). 읽지 못하면 None."""
    try:
        wb = openpyxl.load_workbook(check, read_only=True)
    except Exception:
        return None
    try:
        ws = wb.worksheets[0]
        row = next(ws.iter_rows(min_row=1, max_row=1, values_only=True), ())
        if len(row) >= 2 and row[0] == "원본파일" and row[1]:
            return str(row[1])
        return None
    except Exception:
        return None
    finally:
        wb.close()


def _clear_orphan_checks(waiting: Path, warnings: list[str]):
    """짝이 되는 복사본이 승인대기에 없는 대조표를 지운다(승인됐거나 친구가 복사본을 지운 문서)."""
    folder = waiting / WAITING_CHECKS
    if not folder.is_dir():
        return
    present = {p.name.casefold() for p in waiting.iterdir() if p.is_file()}
    for check in sorted(folder.glob("*.xlsx")):
        if check.name.startswith("~$"):
            continue
        name = _source_name(check)
        if name is None or name.casefold() in present:
            continue
        try:
            check.unlink()
        except OSError as e:
            warnings.append(f"경고: 대조표를 지우지 못함 ({check.name}): {e}. 다음 승인 때 다시 치웁니다")
