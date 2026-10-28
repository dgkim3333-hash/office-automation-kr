"""처음 10건 승인: 양식(엑셀양식 · 문서종류) 단위로 승인 수를 세고, 규칙이 바뀐 양식만 0 으로 되돌린다."""
import os
import shutil
import time
from datetime import datetime, timedelta
from pathlib import Path

import openpyxl
import pytest

import app.job_run as job_run
from app.approval import REQUIRED, approval_gate, approve, form_fingerprint, form_key
from app.job_run import load_records, run_job
from app.jobs import create_job, load_job, save_job
from app.master import write_example
from app.verify import Check

REPO = Path(__file__).resolve().parent.parent
ORDERS = REPO / "samples" / "01_고객오더_엑셀"
PDFS = REPO / "samples" / "03_Invoice_DM_PDF"
EXAMPLE_RULES = REPO / "rules" / "규칙.예시.yaml"
ACME_PO = ORDERS / "고객오더_ACME_PO-ACME-2026-0417.xlsx"
HANIL_PO = ORDERS / "고객오더_HANIL_HDC-PUR-26-0908.xlsx"
INV = PDFS / "INV-2026-10231.pdf"
TITLE = "일일 정리"
COMPANY_ACME = "회사양식_SO입력_PO-ACME-2026-0417.xlsx"
COMPANY_HANIL = "회사양식_SO입력_HDC-PUR-26-0908.xlsx"
CHECK_ACME = "검산_PO-ACME-2026-0417.xlsx"


@pytest.fixture
def env(tmp_path):
    work, inp, out = tmp_path / "work", tmp_path / "in", tmp_path / "out"
    for d in (work, inp, out):
        d.mkdir()
    write_example(work / "고객마스터.xlsx")
    job = create_job(work, TITLE, inp, out, "둘다")
    shutil.copy(EXAMPLE_RULES, job.rules_path)
    return work, inp, out, job


def _drop(src, folder, name=None, age=3600):
    dst = Path(folder) / (name or Path(src).name)
    shutil.copy(src, dst)
    _age(dst, age)
    return dst


def _age(path, age=3600):
    t = time.time() - age
    os.utime(path, (t, t))


def _files(folder):
    folder = Path(folder)
    if not folder.is_dir():
        return []
    return sorted(p.name for p in folder.iterdir() if p.is_file() and not p.name.endswith(".bak.xlsx"))


def _approved(job, key, count):
    """그 양식이 지금 규칙으로 이미 count 건 승인된 상태로 둔다."""
    job.approvals[key] = count
    job.fingerprints[key] = form_fingerprint(job.rules_path, key)


def test_form_key_and_fingerprint(env):
    work, inp, out, job = env
    from app.excel_extract import read_excel
    from app.rules import load_rules

    rules = load_rules(job.rules_path)
    assert form_key(read_excel(ACME_PO, rules)) == "ACME"
    acme = form_fingerprint(job.rules_path, "ACME")
    assert len(acme) == 16 and acme != form_fingerprint(job.rules_path, "HANIL")
    assert form_fingerprint(job.rules_path, "Invoice") != form_fingerprint(job.rules_path, "DM")


def test_new_form_waits_for_approval(env):
    work, inp, out, job = env
    _drop(ACME_PO, inp)

    report = run_job(job, work, gate=approval_gate)

    assert report.counts["승인대기"] == 1 and report.counts["성공"] == 0
    assert _files(out) == []
    assert _files(job.folder / "승인대기") == [ACME_PO.name]
    assert _files(job.folder / "승인대기" / "대조표") == [CHECK_ACME]
    saved = load_job(job.folder)
    assert saved.approvals == {"ACME": 0}
    assert saved.fingerprints == {"ACME": form_fingerprint(job.rules_path, "ACME")}


def test_approve_moves_to_output_and_counts(env):
    work, inp, out, job = env
    _drop(ACME_PO, inp)
    run_job(job, work, gate=approval_gate)
    job.last_run = ""  # 자동 실행 전 아침에 승인하는 경우
    save_job(job)

    report = approve(job, work)

    assert report.counts["성공"] == 1
    assert COMPANY_ACME in _files(out) and f"누적장부_{TITLE}.xlsx" in _files(out)
    assert job.approvals["ACME"] == 1 and load_job(job.folder).approvals["ACME"] == 1
    assert _files(job.folder / "승인대기") == []
    assert _files(job.folder / "승인대기" / "대조표") == []
    assert load_records(job)[-1]["상태"] == "성공"
    assert "ACME 1/10" in report.message
    assert load_job(job.folder).last_run == ""  # 승인은 그날의 자동 실행을 대신하지 않는다(놓친 실행 판단)
    # 입력 폴더의 원본은 이미 처리됨(같은 지문): 다음 자동 실행이 다시 내보내거나 승인대기로 보내지 않는다
    again = run_job(job, work, gate=approval_gate)
    assert again.counts["성공"] == 0 and again.counts["승인대기"] == 0


def test_after_ten_approvals_goes_straight_out(env):
    work, inp, out, job = env
    _approved(job, "ACME", REQUIRED)
    _drop(ACME_PO, inp)

    report = run_job(job, work, gate=approval_gate)

    assert report.counts["승인대기"] == 0 and report.counts["성공"] == 1
    assert COMPANY_ACME in _files(out)
    assert job.approvals["ACME"] == REQUIRED  # 자동으로 나간 건은 세지 않는다


def test_rule_change_resets_only_that_form(env):
    work, inp, out, job = env
    _approved(job, "ACME", REQUIRED)
    _approved(job, "HANIL", REQUIRED)
    text = job.rules_path.read_text(encoding="utf-8")
    changed = text.replace('납품처:     "오른쪽:Deliver To:"', '납품처:     "B7"')  # ACME 머리 칸 하나(같은 값)
    assert changed != text
    job.rules_path.write_text(changed, encoding="utf-8")
    _drop(ACME_PO, inp)
    _drop(HANIL_PO, inp)

    report = run_job(job, work, gate=approval_gate)

    assert report.counts["승인대기"] == 1 and report.counts["성공"] == 1
    assert job.approvals == {"ACME": 0, "HANIL": REQUIRED}
    assert job.fingerprints["ACME"] == form_fingerprint(job.rules_path, "ACME")
    assert _files(job.folder / "승인대기") == [ACME_PO.name]
    assert COMPANY_HANIL in _files(out) and COMPANY_ACME not in _files(out)


def test_approve_with_empty_waiting(env):
    work, inp, out, job = env

    report = approve(job, work)

    assert report.message == "승인할 문서가 없습니다."
    assert report.counts == dict.fromkeys(report.counts, 0) and not report.has_news


def test_approve_rechecks_and_rejects_bad_file(env):
    """승인대기의 복사본을 지금 규칙으로 다시 검산한다. 틀린 파일은 확인필요, 승인 수 그대로, 결과 없음."""
    work, inp, out, job = env
    _drop(ACME_PO, inp)
    run_job(job, work, gate=approval_gate)
    copy = job.folder / "승인대기" / ACME_PO.name
    wb = openpyxl.load_workbook(copy)
    ws = wb.active
    ws["G15"] = 1  # 둘째 품목 금액을 틀리게(수량 x 단가 와 다름)
    wb.save(copy)
    _age(copy)

    report = approve(job, work)

    assert report.counts["확인필요"] == 1 and report.counts["성공"] == 0
    assert job.approvals["ACME"] == 0
    assert _files(out) == []
    assert copy.is_file()  # 고쳐서 다시 승인하거나, 승인하지 않으려면 지우면 된다
    assert _files(job.folder / "승인대기" / "대조표") == [CHECK_ACME]
    note = (job.folder / "확인필요" / f"{ACME_PO.name}.txt").read_text(encoding="utf-8")
    assert "검산 불일치" in note and "다시 승인" in note


# --- W4 리뷰에서 넘어온 것: 승인 실행의 확인필요와 자동 실행 ---

def test_approval_review_keeps_waiting_so_next_day_run_makes_no_second_copy(env, monkeypatch):
    """승인 실행에서 확인필요(고객마스터에 없음)가 돼도 승인대기 상태를 지킨다. 다음 날 자동 실행이 _2 복사본을 만들지 않는다."""
    work, inp, out, job = env
    _drop(ACME_PO, inp)
    run_job(job, work, gate=approval_gate)
    monkeypatch.setattr(job_run, "lookup", lambda master, customer: None)

    first = approve(job, work)

    assert first.counts["확인필요"] == 1 and first.counts["성공"] == 0
    assert load_records(job)[-1]["상태"] == "승인대기"
    monkeypatch.undo()

    tomorrow = datetime.now() + timedelta(days=1)
    scheduled = run_job(job, work, now=tomorrow, gate=approval_gate)

    assert scheduled.counts["승인대기"] == 0 and scheduled.counts["확인필요"] == 0
    assert _files(job.folder / "승인대기") == [ACME_PO.name]
    assert _files(job.folder / "승인대기" / "대조표") == [CHECK_ACME]

    second = approve(job, work)  # 고객마스터를 고친 뒤 같은 날 다시 승인하면 다시 확인한다

    assert second.counts["성공"] == 1 and job.approvals["ACME"] == 1
    assert _files(job.folder / "승인대기") == []
    assert not (job.folder / "확인필요" / f"{ACME_PO.name}.txt").exists()


def test_approval_verify_failure_then_scheduled_run_same_and_next_day(env, monkeypatch):
    work, inp, out, job = env
    _drop(INV, inp)
    run_job(job, work, gate=approval_gate)
    monkeypatch.setattr(job_run, "verify", lambda result, path: [Check("합계금액", "100.00", "1000.00", False)])

    assert approve(job, work).counts["확인필요"] == 1
    monkeypatch.undo()

    for day in (0, 1, 2):
        report = run_job(job, work, now=datetime.now() + timedelta(days=day), gate=approval_gate)
        assert report.counts["승인대기"] == 0
    assert _files(job.folder / "승인대기") == [INV.name]


def test_approve_clears_duplicate_copy_and_orphan_check_files(env):
    """같은 번호의 다른 파일 두 개가 승인대기에 있으면 하나는 성공, 하나는 중복. 둘 다 승인대기에서 치운다.
    친구가 복사본을 지워 승인하지 않은 문서의 대조표도 승인 때 치운다."""
    work, inp, out, job = env
    _drop(ACME_PO, inp)
    other = inp / "고객오더_ACME_다시보냄.xlsx"
    wb = openpyxl.load_workbook(ACME_PO)
    wb.active["J1"] = "resent"
    wb.save(other)
    _age(other)
    _drop(INV, inp)
    run_job(job, work, gate=approval_gate)
    waiting = job.folder / "승인대기"
    assert len(_files(waiting)) == 3 and len(_files(waiting / "대조표")) == 3
    (waiting / INV.name).unlink()  # 승인하지 않음

    report = approve(job, work)

    assert report.counts["성공"] == 1 and report.counts["중복"] == 1
    assert job.approvals["ACME"] == 1
    assert _files(waiting) == []
    assert _files(waiting / "대조표") == []


def test_check_file_is_paired_with_copy_name_not_original_name(env, monkeypatch):
    """입력 폴더에서 같은 이름으로 들어온 다른 문서는 복사본이 _2 가 된다. 대조표는 복사본 이름으로 짝지어
    한쪽을 승인해도 아직 기다리는 쪽의 대조표를 지우지 않는다."""
    work, inp, out, job = env
    _drop(ACME_PO, inp, name="PO.xlsx")
    run_job(job, work, gate=approval_gate)
    (inp / "PO.xlsx").unlink()
    _drop(HANIL_PO, inp, name="PO.xlsx")
    run_job(job, work, gate=approval_gate)
    waiting = job.folder / "승인대기"
    assert _files(waiting) == ["PO.xlsx", "PO_2.xlsx"]
    real = job_run.lookup
    monkeypatch.setattr(job_run, "lookup", lambda master, name: None if "Hanil" in str(name) else real(master, name))

    report = approve(job, work)

    assert report.counts["성공"] == 1 and report.counts["확인필요"] == 1
    assert _files(waiting) == ["PO_2.xlsx"]
    assert _files(waiting / "대조표") == ["검산_HDC-PUR-26-0908.xlsx"]


def _edit_rules(job, old, new):
    text = job.rules_path.read_text(encoding="utf-8")
    assert text.count(old) == 1
    job.rules_path.write_text(text.replace(old, new), encoding="utf-8")


def test_fingerprint_covers_linked_doc_type_and_company_form(env):
    work, inp, out, job = env
    keys = ("ACME", "HANIL", "PO", "Invoice", "DM")
    before = {k: form_fingerprint(job.rules_path, k) for k in keys}

    _edit_rules(job, '통화:       {패턴: "Currency', '통화:       {패턴: "(?i)Currency')  # 문서종류 PO
    after_po = {k: form_fingerprint(job.rules_path, k) for k in keys}
    assert [k for k in keys if after_po[k] != before[k]] == ["ACME", "HANIL", "PO"]

    _edit_rules(job, '"고정:KR-Branch / D. Lee"', '"고정:KR-Branch / D. Kim"')  # 회사양식
    after_form = {k: form_fingerprint(job.rules_path, k) for k in keys}
    assert all(after_form[k] != after_po[k] for k in keys)


def test_rule_change_while_waiting_counts_under_new_rules(env):
    """승인대기에 있는 사이 그 양식의 규칙이 바뀌면, 승인은 새 규칙으로 다시 검산하고 새 규칙의 첫 승인으로 센다."""
    work, inp, out, job = env
    _approved(job, "ACME", 5)
    _drop(ACME_PO, inp)
    run_job(job, work, gate=approval_gate)
    _edit_rules(job, '납품처:     "오른쪽:Deliver To:"', '납품처:     "B7"')

    report = approve(job, work)

    assert report.counts["성공"] == 1
    assert job.approvals["ACME"] == 1
    assert job.fingerprints["ACME"] == form_fingerprint(job.rules_path, "ACME")


def test_approval_rechecks_copy_even_if_same_day_review_record_exists(env, monkeypatch):
    """승인은 사람이 누른 실행이므로 같은 날 · 같은 규칙의 확인필요 기록이 있어도 건너뛰지 않는다."""
    work, inp, out, job = env
    _drop(ACME_PO, inp)
    monkeypatch.setattr(job_run, "lookup", lambda master, customer: None)
    assert run_job(job, work, gate=approval_gate).counts["확인필요"] == 1
    monkeypatch.undo()
    shutil.copy2(inp / ACME_PO.name, job.folder / "승인대기" / ACME_PO.name)  # 친구가 고객마스터를 고친 뒤 직접 넣음

    report = approve(job, work)

    assert report.counts["성공"] == 1 and job.approvals["ACME"] == 1
    assert _files(job.folder / "승인대기") == []
