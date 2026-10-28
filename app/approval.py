"""처음 10건 승인: 새 양식은 처음 10건을 사람이 승인한 뒤부터 결과를 바로 낸다.

승인 수는 양식 단위로 센다(엑셀 PO 는 엑셀양식 이름, PDF 는 문서종류 이름). 규칙.yaml 에서 그 양식에
해당하는 부분이 바뀌면 그 양식만 0 으로 되돌린다. 승인할 문서는 작업 폴더의 승인대기\\ 에 있는 원본 복사본이고,
그 대조표는 승인대기\\대조표\\ 에 있다. 승인하지 않을 문서는 친구가 그 복사본을 지운다.

승인은 친구가 대조표에서 본 값만 내보낸다. 대조표를 만든 뒤 규칙(그 양식 지문)이나 복사본이 바뀌었거나
대조표가 없으면 내보내지 않고 대조표를 새로 만든다(job_run 의 승인 실행이 맡음).
"""
import csv
from pathlib import Path

from app.forms import form_fingerprint, form_key, rules_form_fingerprint  # noqa: F401  (바깥에서 app.approval 로 부른다)
from app.job_run import (
    COMPANY_MODES, REVIEW, WAITING, WAITING_CHECKS, RunReport, _current_state, _documents, check_file_info, file_fingerprint,
    job_lock, load_records, run_job,
)
from app.jobs import Job, save_job

REQUIRED = 10
NOTHING = "승인할 문서가 없습니다."
FINISHED = ("성공", "중복")  # 이 상태가 된 복사본은 더 승인할 것이 없으므로 치운다


def _current_count(job: Job, key: str, rules=None) -> int:
    """그 양식의 승인 수. 규칙(또는 회사 양식을 내는지)이 저장된 지문과 다르면 그 양식만 0 으로 되돌리고 새 지문을 저장한다.

    rules(이번 실행이 읽은 규칙)가 있으면 그 글자에서 지문을 계산한다(실행 중 규칙.yaml 이 바뀌어도 추출과 같은 규칙).
    """
    company = job.output_mode in COMPANY_MODES  # 출력 방식이 회사 양식을 새로 내면 지문이 바뀐다(Ruling 29)
    if rules is not None and getattr(rules, "raw", None) is not None:
        fingerprint = rules_form_fingerprint(rules, key, company)
    else:
        fingerprint = form_fingerprint(job.rules_path, key, company)
    if job.fingerprints.get(key) != fingerprint:
        job.approvals[key] = 0
        job.fingerprints[key] = fingerprint
        save_job(job)
    return job.approvals.get(key, 0)


def approval_gate(job: Job, rules, result) -> bool:
    """True 면 결과를 내지 않고 승인대기로 보낸다(그 양식의 승인 수가 10 미만)."""
    return _current_count(job, form_key(result), rules) < REQUIRED


def approve(job: Job, work) -> RunReport:
    """승인대기의 복사본을 지금 규칙으로 다시 추출 · 검산해, 대조표와 같으면 결과까지 낸다.

    성공한 건마다 그 양식 승인 수 +1. 끝나면(잠금을 쥔 채) 성공 · 중복이 된 복사본과, 짝이 되는 복사본이 없는
    대조표(와 그 복사본의 확인필요 사유)를 지운다. 검산에 실패한 복사본은 승인대기에 남는다.
    """
    waiting = Path(job.folder) / WAITING
    warnings: list[str] = []
    if not _documents(waiting):
        with job_lock(job):
            _clear_orphan_checks(job, waiting, warnings)
        return RunReport(message="\n".join([NOTHING] + warnings))

    approved: list[str] = []
    run_rules = []  # 이번 승인 실행이 읽은 규칙. 승인 실행은 출력 직전에 gate 를 부르므로 거기서 받아 둔다

    def never(job_, rules, result):
        run_rules[:] = [rules]
        return False

    def counted(job_, result):
        key = form_key(result)
        # 규칙이 바뀐 뒤 처음 승인한 건이면 그 양식은 0 에서 센다(지문은 이번 실행이 읽은 규칙에서)
        _current_count(job_, key, run_rules[0] if run_rules else None)
        job_.approvals[key] = job_.approvals.get(key, 0) + 1
        approved.append(key)
        save_job(job_)

    def cleanup(job_, report):
        _clear_finished(job_, waiting, warnings)
        _clear_orphan_checks(job_, waiting, warnings)

    report = run_job(job, work, source_dir=waiting, gate=never, on_success=counted, after=cleanup)

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


def _clear_orphan_checks(job: Job, waiting: Path, warnings: list[str]):
    """짝이 되는 복사본이 승인대기에 없는 대조표를 지운다(승인됐거나 친구가 복사본을 지운 문서).
    그 복사본이 승인 실행에서 남긴 확인필요 사유 파일도 지운다."""
    folder = waiting / WAITING_CHECKS
    if not folder.is_dir():
        return
    present = {p.name.casefold() for p in waiting.iterdir() if p.is_file()}
    for check in sorted(folder.glob("*.xlsx")):
        if check.name.startswith("~$"):
            continue
        name, _ = check_file_info(check)
        if name is None or name.casefold() in present:
            continue
        try:
            check.unlink()
        except OSError as e:
            warnings.append(f"경고: 대조표를 지우지 못함 ({check.name}): {e}. 다음 승인 때 다시 치웁니다")
        _clear_copy_note(job, waiting / name, warnings)


def _clear_copy_note(job: Job, copy: Path, warnings: list[str]):
    """확인필요\\{복사본}.txt 가 그 승인대기 복사본의 사유일 때만 지운다(같은 이름의 입력 폴더 파일 사유는 둔다)."""
    note = Path(job.folder) / REVIEW / f"{copy.name}.txt"
    try:
        if not note.is_file():
            return
        first = note.read_text(encoding="utf-8").splitlines()[:1]
        if first == [f"원본: {copy}"]:
            note.unlink()
    except (OSError, UnicodeDecodeError) as e:
        warnings.append(f"경고: 확인필요 사유 파일을 지우지 못함 ({note.name}): {e}")
