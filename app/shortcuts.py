"""바탕화면 · 시작프로그램 바로가기(.lnk). WScript.Shell 을 PowerShell 로 부른다(외부 모듈 없음).

작업마다 바탕화면에 두 개:
  「지금 실행 - 제목」 = run.bat run NN      (콘솔에 결과 요약, 키를 누르면 닫힘)
  「승인 - 제목」     = run.bat approve NN
바탕화면이 OneDrive 로 옮겨졌어도 Windows 가 알려 주는 실제 바탕화면 위치를 쓴다.
"""
from pathlib import Path

from app import powershell
from app.jobs import Job, _folder_title

RUN_NOW = "지금 실행 - {제목}.lnk"
APPROVE = "승인 - {제목}.lnk"
RUN_BAT = "run.bat"


class ShortcutError(Exception):
    """바로가기를 만들거나 위치를 알아내지 못함. 메시지는 사용자에게 그대로 보인다."""


def _special_folder(name: str, label: str, run=None) -> Path:
    result = powershell.run_script(f"[Environment]::GetFolderPath({powershell.quote(name)})", run)
    lines = [line.strip() for line in result.out.splitlines() if line.strip()]
    if not result.ok or not lines:
        raise ShortcutError(f"{label} 위치를 알아내지 못했습니다({result.err or '빈 응답'})")
    return Path(lines[-1])


def desktop_dir(run=None) -> Path:
    return _special_folder("Desktop", "바탕화면", run)


def startup_dir(run=None) -> Path:
    return _special_folder("Startup", "시작프로그램 폴더", run)


def create(path: Path, target: Path, args: str, workdir: Path, run=None, description: str = "") -> None:
    """바로가기 하나를 만든다(있으면 덮어씀). 만든 뒤 파일이 실제로 생겼는지 확인한다."""
    q = powershell.quote
    script = "\n".join([
        f"$s = (New-Object -ComObject WScript.Shell).CreateShortcut({q(path)})",
        f"$s.TargetPath = {q(target)}",
        f"$s.Arguments = {q(args)}",
        f"$s.WorkingDirectory = {q(workdir)}",
        f"$s.Description = {q(description)}",
        "$s.Save()",
        f"if (-not (Test-Path -LiteralPath {q(path)})) {{ throw '바로가기 파일이 생기지 않았습니다' }}",
    ])
    result = powershell.run_script(script, run)
    if not result.ok:
        raise ShortcutError(f"바로가기를 만들지 못했습니다: {Path(path).name} ({result.err or '알 수 없는 오류'})")


def job_shortcut_paths(job: Job, desktop: Path) -> list[Path]:
    title = _folder_title(job.title)
    return [Path(desktop) / RUN_NOW.format(제목=title), Path(desktop) / APPROVE.format(제목=title)]


def job_shortcuts(job: Job, repo_root, desktop: Path, run=None) -> list[Path]:
    """「지금 실행 - 제목」 · 「승인 - 제목」 을 만들고 그 경로를 돌려준다."""
    repo_root = Path(repo_root)
    number = f"{job.number:02d}"
    run_now, approve = job_shortcut_paths(job, desktop)
    made = []
    for path, args, text in ((run_now, f"run {number}", "지금 실행"), (approve, f"approve {number}", "승인")):
        create(path, repo_root / RUN_BAT, args, repo_root, run, description=f"사무자동화 {job.title} {text}")
        made.append(path)
    return made


def remove_job_shortcuts(job: Job, desktop: Path) -> list[Path]:
    """이 작업의 바탕화면 바로가기 두 개를 지운다(없으면 그냥 지나감). 지운 경로를 돌려준다."""
    removed = []
    for path in job_shortcut_paths(job, desktop):
        if path.exists():
            path.unlink()
            removed.append(path)
    return removed
