"""결과 알림: Windows 알림(PowerShell WinRT), 안 되면 메시지 창. 새 것이 없으면 알리지 않는다.
실제 알림 · 메시지 창은 띄우지 않는다(가짜 실행기 · 가짜 창)."""
import os
import shutil
import time
from datetime import datetime
from pathlib import Path

import pytest

import app.notify as notify
from app.job_run import COUNT_KEYS, RunReport, run_job, unmatched_files
from app.jobs import create_job

REPO = Path(__file__).resolve().parent.parent
EXAMPLE_RULES = REPO / "rules" / "규칙.예시.yaml"
LETTER = REPO / "samples" / "03_Invoice_DM_PDF" / "Letter_warranty_notice.pdf"


def _report(**counts) -> RunReport:
    report = RunReport()
    report.counts = {**dict.fromkeys(COUNT_KEYS, 0), **counts}
    return report


@pytest.fixture
def job(tmp_path):
    inp, out = tmp_path / "in", tmp_path / "out"
    inp.mkdir()
    out.mkdir()
    return create_job(tmp_path / "work", "일일 PO 정리", inp, out, "누적장부")


def test_report_text_skips_zero_counts(job):
    assert notify.report_text(job, _report(성공=3, 확인필요=1)) == "일일 PO 정리: 결과 3건, 확인필요 1건"


def test_report_text_lists_every_kind(job):
    report = _report(중복=1, 승인대기=2)
    report.unmatched_old = 4
    text = notify.report_text(job, report)
    assert text == "일일 PO 정리: 중복 1건, 승인대기 2건, 어느 작업에도 맞지 않는 파일 4개"


def test_report_text_counts_old_excel_separately(job):
    report = _report()
    report.unmatched_old, report.unmatched_xls = 3, 2
    text = notify.report_text(job, report)
    assert "어느 작업에도 맞지 않는 파일 1개" in text
    assert "옛 엑셀(.xls) 2개: 엑셀에서 열어 .xlsx 로 다시 저장하세요" in text


def test_report_text_only_xls(job):
    report = _report()
    report.unmatched_old = report.unmatched_xls = 1
    text = notify.report_text(job, report)
    assert "어느 작업에도" not in text and "옛 엑셀(.xls) 1개" in text


def test_report_text_stopped_and_ledger_locked(job):
    report = _report(보류=2)
    report.stopped = report.ledger_locked = True
    text = notify.report_text(job, report)
    assert "멈췄습니다" in text and "실행기록.txt" in text and "누적장부" in text


def test_no_notification_when_nothing_new(job):
    calls = []
    notify.notify_report(job, _report(보류=1), toast=lambda *a: calls.append("toast"), box=lambda *a: calls.append("box"))
    assert calls == []


def test_toast_shows_report_and_opens_output_folder(job):
    calls = []
    notify.notify_report(job, _report(성공=1), toast=lambda *a: calls.append(a), box=lambda *a: pytest.fail("창 안 됨"))
    ((title, body, folder),) = calls
    assert title == "사무자동화" and body == "일일 PO 정리: 결과 1건" and folder == job.output_dir


def test_click_opens_job_folder_when_no_results(job):
    calls = []
    notify.notify_report(job, _report(승인대기=1), toast=lambda *a: calls.append(a[2]), box=lambda *a: None)
    assert calls == [job.folder]


def test_toast_failure_falls_back_to_message_box(job):
    def broken(*a):
        raise notify.NotifyError("알림 막힘")

    boxes = []
    notify.notify_report(job, _report(확인필요=2), toast=broken, box=lambda title, body: boxes.append((title, body)))
    assert boxes == [("사무자동화", "일일 PO 정리: 확인필요 2건")]


def test_message_box_failure_is_swallowed(job):
    def broken(*a):
        raise OSError("창도 막힘")

    notify.notify_report(job, _report(성공=1), toast=broken, box=broken)  # 예외가 밖으로 나오지 않는다


def test_default_toast_uses_powershell_winrt(job, fake_ps):
    boxes = []
    notify.notify_report(job, _report(성공=1), box=lambda *a: boxes.append(a))
    (script,) = fake_ps.having("ToastNotificationManager")
    assert "Windows.UI.Notifications" in script and "일일 PO 정리: 결과 1건" in script
    assert job.output_dir.as_uri() in script and 'activationType="protocol"' in script
    assert boxes == []


def test_default_toast_failure_uses_box(job, fake_ps):
    fake_ps.fail.append("ToastNotificationManager")
    boxes = []
    notify.notify_report(job, _report(성공=1), box=lambda *a: boxes.append(a))
    assert len(boxes) == 1


def test_toast_text_is_xml_escaped():
    script = notify.toast_script("사무자동화", "A&B <C>", Path("C:/작업/폴더 1"))
    assert "A&amp;B &lt;C&gt;" in script and "A&B <C>" not in script


def test_notify_failure_shows_what_to_fix():
    calls = []
    notify.notify_failure("작업 1 자동 실행", "설치.bat를 먼저 실행하세요", toast=lambda *a: calls.append(a),
                          box=lambda *a: pytest.fail("창 안 됨"))
    ((title, body, folder),) = calls
    assert "실패" in title and "작업 1 자동 실행" in body and "설치.bat" in body


# --- 옛 엑셀(.xls) 따로 세기 ---

def _old(path, hours=30):
    t = time.time() - hours * 3600
    os.utime(path, (t, t))


def test_unmatched_xls_counted_separately_and_hinted(job):
    shutil.copy(EXAMPLE_RULES, job.rules_path)
    (job.input_dir / "옛양식.xls").write_bytes(b"old")
    shutil.copy(LETTER, job.input_dir / "모르는.pdf")  # 어느 작업에도 안 맞는 PDF
    for p in job.input_dir.iterdir():
        _old(p)
    files = unmatched_files(job, job.folder.parent.parent, datetime.now())
    assert sorted(p.name for p in files) == ["모르는.pdf", "옛양식.xls"]

    report = run_job(job, job.folder.parent.parent)
    assert report.unmatched_old == 2 and report.unmatched_xls == 1
    assert "옛 엑셀(.xls) 1개" in report.message and ".xlsx 로 다시 저장" in report.message


def test_toast_checks_notifications_are_enabled():
    """사용자 · 그룹 정책이 알림을 꺼 두면 Show 는 성공해도 아무것도 안 보이므로, 설정을 보고 꺼져 있으면 실패로."""
    script = notify.toast_script("사무자동화", "본문", None)
    assert "$n.Setting" in script and "'Enabled'" in script and "throw" in script
    assert script.index("$n.Setting") < script.index("$n.Show(")


def test_notifications_turned_off_falls_back_to_box(job, fake_ps):
    fake_ps.fail.append("$n.Setting")  # 알림 꺼짐 → 스크립트가 throw → 반환 코드 1
    boxes = []
    notify.notify_report(job, _report(성공=1), box=lambda *a: boxes.append(a))
    assert len(fake_ps.having("ToastNotificationManager")) == 1
    assert boxes == [("사무자동화", "일일 PO 정리: 결과 1건")]


def test_notify_failure_click_opens_log_folder(tmp_path):
    calls = []
    notify.notify_failure("작업 1 자동 실행", "고칠 것", toast=lambda *a: calls.append(a), folder=tmp_path)
    assert calls[0][2] == tmp_path
