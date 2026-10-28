"""매일 실행 등록과 놓친 실행.

등록은 막히면 다음 방법으로 넘어간다(설계 6절):
  1. Windows 작업 스케줄러: 「사무자동화\\NN_제목」, 매일 정해진 시각, 현재 사용자 · 로그인했을 때만,
     놓친 실행은 PC 를 켜면 바로(StartWhenAvailable). 등록 뒤 다시 조회해서 있어야 성공으로 본다.
  2. 시작프로그램 폴더의 「사무자동화 놓친 작업 실행」 바로가기 하나(모든 작업이 같이 씀):
     로그인할 때 오늘 시각이 지났는데 아직 안 돈 작업을 차례로 돌린다(catch-up).
  3. 둘 다 막히면 「없음」: 바탕화면 「지금 실행」만 쓴다.
어느 방법이 쓰였는지는 작업설정.yaml 의 스케줄러에 남긴다.
"""
import time
from datetime import datetime
from pathlib import Path

from app import powershell, shortcuts
from app.jobs import Job, JobError, _folder_title, list_jobs, load_job, save_job

TASK_PREFIX = "사무자동화"
TASK_PATH = f"\\{TASK_PREFIX}\\"
SCHEDULER = "작업스케줄러"
STARTUP = "시작프로그램"
NONE = "없음"
CATCH_UP_LINK = "사무자동화 놓친 작업 실행.lnk"
TIME_LIMIT_HOURS = 2  # 한 번 실행이 이보다 길면 작업 스케줄러가 멈춘다(멈춰 있는 실행이 다음 날을 막지 않게)
# 자동 실행 · 놓친 실행이 잠긴 작업(승인 · job set · pause 등이 쥠)을 만나면 이 간격으로 이 시간까지 다시 해 본다
LOCK_RETRY_SECONDS = 30
LOCK_RETRY_LIMIT_SECONDS = 15 * 60


def _sleep(seconds: float) -> None:
    """기다리기(테스트가 바꿔 끼운다)."""
    time.sleep(seconds)


class ScheduleError(Exception):
    """작업 스케줄러에서 지우기 · 멈추기 · 다시 켜기를 하지 못함. 메시지는 사용자에게 그대로 보인다."""


def task_name(job: Job) -> str:
    return f"{job.number:02d}_{_folder_title(job.title)}"


def full_name(job: Job) -> str:
    return f"{TASK_PREFIX}\\{task_name(job)}"


def pythonw(repo_root) -> Path:
    """창 없이 도는 파이썬(저장소의 .venv). 매일 실행 · 놓친 실행이 이것으로 돈다."""
    return Path(repo_root) / ".venv" / "Scripts" / "pythonw.exe"


def _where(name: str) -> str:
    return f"-TaskPath {powershell.quote(TASK_PATH)} -TaskName {powershell.quote(name)}"


def _register_script(job: Job, repo_root) -> str:
    q = powershell.quote
    repo_root = Path(repo_root)
    return "\n".join([
        f"$action = New-ScheduledTaskAction -Execute {q(pythonw(repo_root))} "
        f"-Argument {q(f'-m app run {job.number:02d} --자동')} -WorkingDirectory {q(repo_root)}",
        f"$trigger = New-ScheduledTaskTrigger -Daily -At {q(job.run_time)}",
        "$settings = New-ScheduledTaskSettingsSet -StartWhenAvailable -AllowStartIfOnBatteries "
        f"-DontStopIfGoingOnBatteries -MultipleInstances IgnoreNew -ExecutionTimeLimit (New-TimeSpan -Hours {TIME_LIMIT_HOURS})",
        "$user = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name",
        "$principal = New-ScheduledTaskPrincipal -UserId $user -LogonType Interactive -RunLevel Limited",
        f"Register-ScheduledTask {_where(task_name(job))} -Action $action -Trigger $trigger -Settings $settings "
        f"-Principal $principal -Description {q(f'사무자동화 매일 실행: {job.title}')} -Force | Out-Null",
    ])


def _on_task(name: str, verb: str) -> str:
    """그 이름의 작업이 있을 때만 verb(Unregister · Disable · Enable)를 한다."""
    tail = " -Confirm:$false" if verb == "Unregister" else " | Out-Null"
    return (f"$t = Get-ScheduledTask {_where(name)} -ErrorAction SilentlyContinue\n"
            f"if ($t) {{ {verb}-ScheduledTask {_where(name)}{tail} }}")


def _register_task(job: Job, repo_root, run, notes: list) -> bool:
    result = powershell.run_script(_register_script(job, repo_root), run)
    if not result.ok:
        notes.append(f"작업 스케줄러에 등록하지 못함({result.err or '알 수 없는 오류'})")
        return False
    query = powershell.run_script(f"(Get-ScheduledTask {_where(task_name(job))}).TaskName", run)
    if not query.ok:
        notes.append(f"작업 스케줄러에 등록했다고 했지만 다시 조회하니 없음({query.err or '알 수 없는 오류'})")
        return False
    if job.paused:
        paused = powershell.run_script(_on_task(task_name(job), "Disable"), run)
        if not paused.ok:
            notes.append("작업 스케줄러에서 일시중지로 바꾸지 못함(자동 실행 때 일시중지를 보고 건너뜀)")
    return True


def _startup_link(repo_root, run, startup, notes: list) -> bool:
    try:
        folder = Path(startup) if startup is not None else shortcuts.startup_dir(run)
        link = folder / CATCH_UP_LINK
        if link.exists():
            return True
        shortcuts.create(link, pythonw(repo_root), "-m app catch-up", Path(repo_root), run,
                         description="사무자동화: 오늘 시각이 지났는데 아직 안 돈 작업 실행")
        return True
    except (shortcuts.ShortcutError, OSError) as e:
        notes.append(f"시작프로그램 바로가기를 만들지 못함({e})")
        return False


def register(job: Job, repo_root, run=None, startup: Path | None = None, notes: list | None = None) -> str:
    """매일 실행을 등록하고 쓰인 방법(작업스케줄러 · 시작프로그램 · 없음)을 돌려준다. job.scheduler 에 저장한다.
    notes 를 주면 막힌 이유를 거기에 더한다(화면에 보여 줄 참고)."""
    notes = notes if notes is not None else []
    if _register_task(job, repo_root, run, notes):
        method = SCHEDULER
    elif _startup_link(repo_root, run, startup, notes):
        method = STARTUP
    else:
        method = NONE
    job.scheduler = method
    save_job(job)
    return method


def _run_on_task(name: str, verb: str, run, what: str) -> None:
    result = powershell.run_script(_on_task(name, verb), run)
    if not result.ok:
        raise ScheduleError(f"작업 스케줄러에서 「{TASK_PREFIX}\\{name}」을 {what} 못했습니다({result.err or '알 수 없는 오류'})")


def unregister(job: Job, run=None) -> None:
    """작업 스케줄러에서 지운다(없으면 그냥 지나감). 실패하면 ScheduleError."""
    _run_on_task(task_name(job), "Unregister", run, "지우지")


def pause(job: Job, run=None) -> None:
    _run_on_task(task_name(job), "Disable", run, "멈추지")


def resume(job: Job, run=None) -> None:
    _run_on_task(task_name(job), "Enable", run, "다시 켜지")


def registered_names(run=None) -> list[str]:
    """작업 스케줄러의 사무자동화 폴더에 있는 작업 이름들."""
    result = powershell.run_script(
        f"Get-ScheduledTask -TaskPath {powershell.quote(TASK_PATH)} -ErrorAction SilentlyContinue | "
        "ForEach-Object { $_.TaskName }", run)
    if not result.ok:
        raise ScheduleError(f"작업 스케줄러 목록을 읽지 못했습니다({result.err or '알 수 없는 오류'})")
    return [line.strip() for line in result.out.splitlines() if line.strip()]


def remove_leftovers(keep: set[str], run=None) -> list[str]:
    """사무자동화 폴더에서 keep 에 없는 작업(지운 작업이 남긴 것)을 지우고 그 이름들을 돌려준다."""
    removed = []
    for name in registered_names(run):
        if name not in keep:
            _run_on_task(name, "Unregister", run, "지우지")
            removed.append(name)
    return removed


# --- 놓친 실행 ---

def due_jobs(jobs, now: datetime) -> list[Job]:
    """오늘 시각이 지났는데 오늘 아직 안 돈 작업(일시중지 제외), 번호 순."""
    today, clock = now.date().isoformat(), now.strftime("%H:%M")
    return sorted((j for j in jobs if not j.paused and j.last_run != today and j.run_time <= clock),
                  key=lambda j: j.number)


def _default_runner(work):
    from app.approval import approval_gate
    from app.job_run import run_job

    return lambda job: run_job(job, work, gate=approval_gate)


def retry_locked(call):
    """call() 이 JobLocked 면 LOCK_RETRY_SECONDS 마다 다시, LOCK_RETRY_LIMIT_SECONDS 까지 해 본다.
    그래도 잠겨 있으면 마지막 JobLocked 를 그대로 던진다(부른 쪽이 기록하고 알린다)."""
    from app.job_run import JobLocked

    waited = 0
    while True:
        try:
            return call()
        except JobLocked:
            if waited >= LOCK_RETRY_LIMIT_SECONDS:
                raise
        _sleep(LOCK_RETRY_SECONDS)
        waited += LOCK_RETRY_SECONDS


def catch_up(work, now=None, runner=None, locked: list | None = None) -> list[str]:
    """놓친 작업을 번호 순으로 하나씩 실행하고 실행한 작업 이름들을 돌려준다.
    실행 직전에 작업설정.yaml 을 다시 읽어 그새 다른 실행(작업 스케줄러)이 끝냈으면 건너뛴다. 그새 지워진 작업도 건너뛴다.
    잠긴 작업(승인 · job set 등)은 retry_locked 처럼 기다렸다가 다시 해 보고(다시 읽어 아직 할 일일 때만),
    끝내 잠겨 있으면 그 작업을 locked 에 더한다(부른 쪽이 기록하고 알린다. 조용히 건너뛰지 않게)."""
    from app.job_run import JobLocked

    now = now or datetime.now()
    runner = runner or _default_runner(work)
    ran = []
    for job in due_jobs(list_jobs(work), now):
        waited = 0
        while True:
            try:
                fresh = load_job(job.folder)
            except JobError:
                break
            if not due_jobs([fresh], now):
                break  # 그새 다른 실행이 끝냈거나 멈춤으로 바뀜
            try:
                runner(fresh)
            except JobLocked:
                if waited >= LOCK_RETRY_LIMIT_SECONDS:
                    if locked is not None:
                        locked.append(fresh)
                    break
                _sleep(LOCK_RETRY_SECONDS)
                waited += LOCK_RETRY_SECONDS
                continue
            except JobError:
                break
            ran.append(fresh.name)
            break
    return ran
