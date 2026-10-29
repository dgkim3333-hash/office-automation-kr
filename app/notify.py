"""결과 알림(설계 12-2). PC 안에서만 뜬다.

- 기본: PowerShell 의 Windows 알림(WinRT ToastNotificationManager, 외부 모듈 없음). 화면 오른쪽 아래.
  누르면 출력 폴더가 열린다(결과가 0건이면 작업 폴더).
- 알림이 실패하면 메시지 창(user32 MessageBoxW)으로 대신한다.
- 새 것이 없으면(has_news 가 False) 아무것도 하지 않는다. 실행 잠금을 푼 뒤에 부른다.
"""
from pathlib import Path
from xml.sax.saxutils import escape

from app import powershell

TITLE = "사무자동화"
FAILURE_TITLE = "사무자동화: 자동 실행 실패"
# Windows PowerShell 의 앱 ID. 따로 앱을 등록하지 않아도 이 이름으로 알림을 띄울 수 있다
APP_ID = "{1AC14E77-02E7-4E5D-B744-2EB1AE5198B7}\\WindowsPowerShell\\v1.0\\powershell.exe"
LABELS = (("성공", "결과"), ("중복", "중복"), ("확인필요", "확인필요"), ("승인대기", "승인대기"))
XLS_HINT = "옛 엑셀(.xls) {n}개: 엑셀에서 열어 .xlsx 로 다시 저장하세요"
MB_ICONINFORMATION = 0x40
MB_ICONWARNING = 0x30
MB_TOPMOST = 0x40000


class NotifyError(Exception):
    """Windows 알림을 띄우지 못함(메시지 창으로 대신한다)."""


def report_text(job, report) -> str:
    """예: 「일일 PO 정리: 결과 3건, 확인필요 1건」. 0인 항목은 뺀다."""
    parts = []
    if report.stopped:
        parts.append("실행이 중간에 멈췄습니다(작업 폴더의 실행기록.txt 를 보세요)")
    for key, label in LABELS:
        if report.counts.get(key, 0):
            parts.append(f"{label} {report.counts[key]}건")
    if report.ledger_locked:
        parts.append("누적장부가 엑셀에서 열려 있어 보류(엑셀을 닫으세요)")
    xls = getattr(report, "unmatched_xls", 0)
    others = report.unmatched_old - xls
    if others > 0:
        parts.append(f"어느 작업에도 맞지 않는 파일 {others}개")
    if xls > 0:
        parts.append(XLS_HINT.format(n=xls))
    text = f"{job.title}: " + ", ".join(parts)
    alerts = getattr(report, "alerts", [])
    return "\n".join([text] + [f"경고: {a}" for a in alerts])


def toast_script(title: str, body: str, folder: Path | None) -> str:
    launch = ""
    if folder:  # 누르면 그 폴더를 연다(file:/// 주소, 한글은 %인코딩)
        uri = escape(Path(folder).as_uri(), {'"': "&quot;"})
        launch = f' activationType="protocol" launch="{uri}"'
    xml = (f"<toast{launch}><visual><binding template=\"ToastGeneric\">"
           f"<text>{escape(title)}</text><text>{escape(body)}</text></binding></visual></toast>")
    q = powershell.quote
    return "\n".join([
        "[Windows.UI.Notifications.ToastNotificationManager, Windows.UI.Notifications, ContentType = WindowsRuntime] | Out-Null",
        "[Windows.Data.Xml.Dom.XmlDocument, Windows.Data.Xml.Dom.XmlDocument, ContentType = WindowsRuntime] | Out-Null",
        "$xml = New-Object Windows.Data.Xml.Dom.XmlDocument",
        f"$xml.LoadXml({q(xml)})",
        "$toast = New-Object Windows.UI.Notifications.ToastNotification $xml",
        f"$n = [Windows.UI.Notifications.ToastNotificationManager]::CreateToastNotifier({q(APP_ID)})",
        # 사용자나 회사 정책이 알림을 꺼 두면 Show 는 성공해도 아무것도 안 보인다. 그러면 실패로 돌려 메시지 창으로
        "if ($n.Setting -ne 'Enabled') { throw \"Windows 알림이 꺼져 있음($($n.Setting))\" }",
        "$n.Show($toast)",
    ])


def _toast(title: str, body: str, folder: Path | None) -> None:
    result = powershell.run_script(toast_script(title, body, folder), timeout=60)
    if not result.ok:
        raise NotifyError(result.err or "Windows 알림을 띄우지 못함")


def _message_box(title: str, body: str, icon: int = MB_ICONINFORMATION) -> None:
    import ctypes

    ctypes.windll.user32.MessageBoxW(None, body, title, icon | MB_TOPMOST)


def _show(title, body, folder, toast, box) -> None:
    """알림을 띄운다. 실패하면 메시지 창. 둘 다 실패해도 예외를 밖으로 내지 않는다(실행 결과는 이미 기록됨)."""
    try:
        (toast or _toast)(title, body, folder)
        return
    except Exception:
        pass
    try:
        (box or _message_box)(title, body)
    except Exception:
        pass


def notify_report(job, report, toast=None, box=None) -> None:
    """실행 결과 알림. 알릴 것이 없으면 아무것도 하지 않는다."""
    if not report.has_news:
        return
    folder = Path(job.output_dir) if report.counts.get("성공", 0) > 0 else Path(job.folder)
    _show(TITLE, report_text(job, report), folder, toast, box)


def notify_failure(what: str, message: str, toast=None, box=None, folder: Path | None = None) -> None:
    """자동 실행이 작업을 찾기도 전에 실패했을 때(작업폴더.txt 없음 · 작업 설정 깨짐 등). 무엇을 고칠지 알린다."""
    _show(FAILURE_TITLE, f"{what}: {message}", folder, toast,
          box or (lambda title, body: _message_box(title, body, MB_ICONWARNING)))
