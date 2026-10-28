"""명령 입구. 모든 명령은 main() 의 가드(이 PC 전용) · 인터넷 차단 안에서 돈다.

  uv run python -m app                            작업 목록과 쓰는 법 안내(아무것도 처리하지 않음)
  uv run python -m app init [경로]                작업 폴더 만들기(설치.bat). 경로를 주면 묻지 않음
  uv run python -m app job new --제목 T --입력 DIR --출력 DIR --방식 회사양식|누적장부|둘다
  uv run python -m app job list                   작업 목록(시각 · 실행 방법 · 승인 현황)
  uv run python -m app job delete|pause|resume N  작업 지우기 · 멈추기 · 다시 시작(매일 실행 · 바로가기도 같이 정리)
  uv run python -m app job repair                 모든 작업의 매일 실행 · 바탕화면 바로가기 다시 만들기
  uv run python -m app run N [--자동]             작업 실행(바탕화면 「지금 실행」, 매일 자동 실행은 --자동)
  uv run python -m app approve N                  승인대기 문서 승인(바탕화면 「승인」)
  uv run python -m app catch-up                   오늘 시각이 지났는데 아직 안 돈 작업 실행(시작프로그램 바로가기)
  uv run python -m app check FILE [--작업 N]      그 작업 규칙으로 추출 · 검산 결과 표만 보기(아무것도 바꾸지 않음)
  uv run python -m app pick-folder "안내 글"      폴더 고르기 창. 고른 경로를 마지막 줄에 출력
  uv run --extra outlook python -m app outlook N  (선택) Outlook '자동화' 폴더 첨부 → 그 작업의 입력 폴더
"""
import argparse
import os
import shutil
import sys
import traceback
from datetime import date, datetime
from pathlib import Path

from app import config, schedule, shortcuts
from app import init as init_module
from app.approval import REQUIRED, approval_gate, approve
from app.excel_extract import read_excel
from app.extract import extract
from app.init import InitError, init
from app.job_run import COMPANY_MODES, LEDGER_MODES, MASTER_PREFIX, JobLocked, job_lock, ledger_path, reload_job, run_job
from app.notify import notify_failure, notify_report
from app.jobs import JOBS_DIR, JobError, broken_jobs, create_job, find_job, list_jobs, save_job
from app.local_guard import BANNER, LocalOnlyError, block_network, check_environment, check_work_dir
from app.markup import MARKUP_FILE, MarkupError, load_markup, markup_checks, missing_parts, written_markup
from app.master import MasterError, load_master, lookup
from app.organize import free_name
from app.outlook import collect_from_outlook
from app.outputs import _mark, company_form_name, company_rows
from app.rules import RuleError, load_rules
from app.verify import failed, verify

RUN_RECORD = "실행기록.txt"  # 작업 폴더. 실행 · 승인마다 결과 요약을 덧붙인다(아무도 안 볼 때의 기록)
DELETED_DIR = "지운작업"     # 작업\지운작업\: job delete 가 작업 폴더를 옮겨 두는 곳(규칙 · 처리기록 보존)
CANCELLED = "취소함"
XLS_REASON = ("옛 엑셀 형식(.xls)은 읽지 못합니다. 엑셀에서 열어 「다른 이름으로 저장」 → .xlsx 로 저장한 뒤 "
              "입력 폴더에 다시 넣으세요")
DESKTOP_HINT = "작업을 실행하려면 바탕화면의 「지금 실행 - 작업 이름」을 쓰세요. 승인은 「승인 - 작업 이름」입니다."
NO_JOBS = "아직 작업이 없습니다. Code 탭에서 작업을 만드세요(04장)."
REPAIR_HINT = "Code 탭에서 「작업 자동 실행 다시 등록해 줘」라고 말하면 정리합니다(job repair)."
JOB_FIX_HINT = ("작업 설정이 깨졌으면 Code 탭에서 「이 작업 설정 고쳐 줘」라고 말하세요. "
                "지운 작업인데 이 알림이 계속 뜨면 " + REPAIR_HINT)

# --- 명령 해석 ---

def _parser():
    p = argparse.ArgumentParser(prog="python -m app", description="PDF · 엑셀 → 회사 양식 · 누적장부 사무 자동화")
    sub = p.add_subparsers(dest="command", metavar="명령")

    s = sub.add_parser("init", help="작업 폴더 만들기(설치.bat)")
    s.add_argument("target", nargs="?", help="작업 폴더 경로(없으면 고르기 창)")

    job = sub.add_parser("job", help="작업 만들기 · 목록 · 지우기 · 멈추기")
    jobs = job.add_subparsers(dest="job_command", metavar="new|list|delete|pause|resume|repair", required=True)
    new = jobs.add_parser("new", help="작업 만들기")
    new.add_argument("--제목", dest="title", required=True)
    new.add_argument("--입력", dest="input", required=True, help="입력 폴더")
    new.add_argument("--출력", dest="output", required=True, help="출력 폴더")
    new.add_argument("--방식", dest="mode", required=True, help="회사양식 · 누적장부 · 둘다")
    jobs.add_parser("list", help="작업 목록")
    jobs.add_parser("repair", help="모든 작업의 매일 실행 · 바탕화면 바로가기 다시 만들기")
    for name, text in (("delete", "작업 지우기"), ("pause", "매일 실행 멈추기"), ("resume", "매일 실행 다시 시작")):
        jobs.add_parser(name, help=text).add_argument("job", help="작업 번호 또는 이름")

    r = sub.add_parser("run", help="작업 실행")
    r.add_argument("job", help="작업 번호 또는 이름")
    r.add_argument("--자동", dest="auto", action="store_true", help="매일 자동 실행(일시중지면 건너뜀)")

    sub.add_parser("approve", help="승인대기 문서 승인").add_argument("job", help="작업 번호 또는 이름")
    sub.add_parser("catch-up", help="오늘 시각이 지났는데 아직 안 돈 작업 실행(시작프로그램 바로가기)")

    c = sub.add_parser("check", help="추출 · 검산 결과 표만 보기")
    c.add_argument("target", help="PDF 또는 엑셀 파일 경로")
    c.add_argument("--작업", dest="job", help="작업 번호 또는 이름(작업이 하나면 생략 가능)")

    f = sub.add_parser("pick-folder", help="폴더 고르기 창")
    f.add_argument("title", nargs="?", default="폴더를 고르세요", help="창에 보일 안내 글")

    sub.add_parser("outlook", help="(선택) Outlook 첨부 → 작업 입력 폴더").add_argument("job", help="작업 번호 또는 이름")
    return p


def _utf8_console():
    """출력이 파이프여도(Claude Code 가 실행할 때) 한글이 깨지거나 죽지 않게 utf-8 로 고정."""
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8", errors="replace")
        except (AttributeError, ValueError):
            pass


def main(argv=None, repo_root=config.REPO_ROOT) -> int:
    _utf8_console()
    args = _parser().parse_args(argv)
    try:
        check_environment()
    except LocalOnlyError as e:
        print(e)
        return 1
    print(BANNER)
    print()
    restore_network = block_network()
    try:
        return _dispatch(args, repo_root)
    finally:
        restore_network()


def _dispatch(args, repo_root) -> int:
    if args.command == "catch-up" or (args.command == "run" and args.auto):
        return _automatic(args, repo_root)
    try:
        if args.command == "init":
            # 경로를 주면 묻지 않고 그 자리에 만든다(설치 스크립트용)
            init(work_dir=args.target, repo_root=repo_root)
            return 0
        if args.command == "pick-folder":
            return _pick_folder(args.title)
        work = config.work_dir(repo_root)
        check_work_dir(work)
        if args.command is None:
            return _overview(work)
        if args.command == "job":
            return _job(args, work, repo_root)
        if args.command == "run":
            return _run(work, args.job)
        if args.command == "approve":
            job = find_job(work, args.job)
            return _finish(job, approve(job, work), "승인")
        if args.command == "check":
            return _check(work, args)
        if args.command == "outlook":
            collect_from_outlook(find_job(work, args.job).input_dir)
            return 0
        raise AssertionError(args.command)
    except (config.ConfigError, InitError, LocalOnlyError, JobError, JobLocked) as e:
        print(e)
        return 1
    except RuleError as e:
        print(f"규칙.yaml 오류: {e}")
        return 1


# --- 작업 목록 · 만들기 · 지우기 · 멈추기 ---

def _approval_text(job) -> str:
    if not job.approvals:
        return "아직 없음"
    return ", ".join(f"{key} {n}/{REQUIRED}" for key, n in job.approvals.items())


def _print_jobs(work) -> int:
    """작업 목록을 출력하고 (읽힌 작업 + 깨진 작업) 수를 돌려준다."""
    jobs, broken = list_jobs(work), broken_jobs(work)
    for job in jobs:
        line = (f"{job.name}  매일 {job.run_time}  실행 방법: {job.scheduler or '등록 안 됨'}  "
                f"승인: {_approval_text(job)}")
        print(line + ("  (일시중지)" if job.paused else ""))
    for folder, reason in broken:
        print(f"{folder.name}  설정이 깨져 실행할 수 없음: {reason}")
    return len(jobs) + len(broken)


def _overview(work) -> int:
    """인자 없이 실행(run.bat 을 그냥 누름): 처리는 하지 않고 작업 목록과 쓰는 법만 보여 준다."""
    if not _print_jobs(work):
        print(NO_JOBS)
        return 0
    print()
    print(DESKTOP_HINT)
    return 0


def _job(args, work, repo_root) -> int:
    if args.job_command == "new":
        return _job_new(work, repo_root, args)
    if args.job_command == "list":
        if not _print_jobs(work):
            print(NO_JOBS)
        return 0
    if args.job_command == "repair":
        return _job_repair(work, repo_root)
    if args.job_command == "delete":
        return _job_delete(work, args.job)
    return _set_paused(work, args.job, args.job_command == "pause")


def _abs(path) -> Path:
    return Path(os.path.abspath(path))


def _job_new(work, repo_root, args) -> int:
    job = create_job(work, args.title, _abs(args.input), _abs(args.output), args.mode)
    print(f"작업 폴더: {job.folder}")
    print(f"입력 폴더: {job.input_dir}")
    print(f"출력 폴더: {job.output_dir}")
    print(f"출력 방식: {job.output_mode}")
    try:
        # 빈 뼈대: Claude 가 예시 파일을 보고 채운다. 채우기 전에는 실행하면 「규칙.yaml 오류」로 멈춘다
        shutil.copyfile(Path(repo_root) / config.SKELETON_RULES, job.rules_path)
        print(f"규칙: {job.rules_path} (빈 뼈대. 예시 파일을 보고 채워야 실행됩니다)")
    except OSError as e:
        print(f"경고: 규칙 뼈대를 넣지 못함 ({e}). {job.rules_path} 를 새로 만들어 채우세요.")
    print(f"작업 만듦: {job.name} (매일 {job.run_time})")
    _install_automation(job, repo_root)
    return 0


# --- 매일 실행 등록 · 바탕화면 바로가기 ---

def _method_line(job, method: str) -> str:
    if method == schedule.SCHEDULER:
        return f"자동 실행: Windows 작업 스케줄러 (매일 {job.run_time})"
    if method == schedule.STARTUP:
        return (f"자동 실행: 시작프로그램 바로가기 (작업 스케줄러가 막혀 있어 대신 씀. PC 에 로그인할 때 "
                f"매일 {job.run_time} 이 지났고 오늘 아직 안 돌았으면 실행합니다. 정각 실행은 안 됩니다)")
    return (f"자동 실행: 없음 (작업 스케줄러 · 시작프로그램이 모두 막혀 있습니다). "
            f"바탕화면 「지금 실행 - {job.title}」을 눌러 실행하세요.")


def _install_automation(job, repo_root) -> None:
    """매일 실행을 등록하고(막히면 다음 방법) 바탕화면 바로가기를 만든다. 쓰인 방법을 한 줄로 보여 준다."""
    notes: list[str] = []
    try:
        with job_lock(job):  # 실행이 끝날 때 작업설정.yaml 을 저장하므로 겹치지 않게
            reload_job(job)
            method = schedule.register(job, repo_root, notes=notes)
    except (JobLocked, JobError) as e:
        print(f"{job.name}: 자동 실행을 등록하지 못함({e}). 실행이 끝난 뒤 다시 하세요.")
        return
    print(_method_line(job, method))
    for note in notes:
        print(f"  참고: {note}")
    if method != schedule.NONE and not schedule.pythonw(repo_root).is_file():
        print(f"경고: {schedule.pythonw(repo_root)} 가 없어 자동 실행이 돌지 못합니다. 설치.bat 을 다시 실행하세요.")
    _make_shortcuts(job, repo_root)


def _make_shortcuts(job, repo_root) -> None:
    try:
        made = shortcuts.job_shortcuts(job, repo_root, shortcuts.desktop_dir())
    except shortcuts.ShortcutError as e:
        print(f"경고: 바탕화면 바로가기를 만들지 못함 ({e}). 대신 {Path(repo_root) / 'run.bat'} 이 있는 폴더에서 "
              f"명령 창으로 「run.bat run {job.number:02d}」(승인은 「run.bat approve {job.number:02d}」)을 실행하세요.")
        return
    print("바탕화면 바로가기: " + ", ".join(p.name for p in made))


def _remove_automation(job) -> None:
    """지운 작업의 매일 실행 · 바탕화면 바로가기를 정리한다. 실패해도 지우기는 이미 끝났으므로 경고만."""
    try:
        schedule.unregister(job)
    except schedule.ScheduleError as e:
        print(f"경고: {e}. 작업 스케줄러에 「{schedule.full_name(job)}」이 남아 있을 수 있습니다. {REPAIR_HINT}")
    try:
        removed = shortcuts.remove_job_shortcuts(job, shortcuts.desktop_dir())
    except (shortcuts.ShortcutError, OSError) as e:
        print(f"경고: 바탕화면 바로가기를 지우지 못함 ({e}). "
              f"「지금 실행 - {job.title}」 · 「승인 - {job.title}」을 직접 지우세요.")
        return
    if removed:
        print("바탕화면 바로가기 지움: " + ", ".join(p.name for p in removed))


def _job_repair(work, repo_root) -> int:
    """모든 작업의 매일 실행을 다시 등록하고 바로가기를 다시 만든다. 지운 작업이 남긴 작업 스케줄러 항목도 지운다."""
    jobs, broken = list_jobs(work), broken_jobs(work)
    if not jobs and not broken:
        print(NO_JOBS)
    for job in jobs:
        print(f"[{job.name}]")
        _install_automation(job, repo_root)
    keep = {schedule.task_name(job) for job in jobs} | {folder.name for folder, _ in broken}
    try:
        for name in schedule.remove_leftovers(keep):
            print(f"지운 작업이 남긴 자동 실행을 지움: {schedule.TASK_PREFIX}\\{name}")
    except schedule.ScheduleError as e:
        print(f"경고: {e}")
    for folder, reason in broken:
        print(f"{folder.name}: 설정이 깨져 다시 등록하지 못함. {reason}")
    return 1 if broken else 0


def _job_delete(work, key) -> int:
    job = find_job(work, key)
    with job_lock(job):  # 실행 중이면 JobLocked(지우지 않음)
        pass
    trash = Path(work) / JOBS_DIR / DELETED_DIR
    try:
        trash.mkdir(exist_ok=True)
        dst = free_name(trash / job.name)
        os.rename(job.folder, dst)
    except OSError as e:
        print(f"작업 폴더를 옮기지 못했습니다({e}). 작업 폴더의 파일을 연 프로그램을 닫고 다시 하세요.")
        return 1
    print(f"작업 지움: {job.name}")
    print(f"규칙 · 처리기록은 {dst} 에 옮겨 두었습니다. 필요 없으면 직접 지우세요.")
    _remove_automation(job)
    return 0


def _set_paused(work, key, paused: bool) -> int:
    job = find_job(work, key)
    with job_lock(job):  # 실행이 끝날 때 작업설정.yaml 을 저장하므로 겹치지 않게
        reload_job(job)
        job.paused = paused
        save_job(job)
    print(f"{job.name}: " + ("매일 실행을 멈췄습니다." if paused else "매일 실행을 다시 시작합니다."))
    if job.scheduler == schedule.SCHEDULER:
        try:
            (schedule.pause if paused else schedule.resume)(job)
        except schedule.ScheduleError as e:
            # 일시중지는 작업설정.yaml 에 저장됐고 자동 실행이 그것을 보고 건너뛰므로, 멈추기는 그래도 지켜진다
            print(f"경고: {e}. " + ("설정에는 저장했으므로 자동 실행은 건너뜁니다." if paused else REPAIR_HINT))
    return 0


# --- 실행 · 승인 ---

def _run(work, key) -> int:
    job = find_job(work, key)
    report = run_job(job, work, gate=approval_gate)
    return _finish(job, report, "지금 실행")


def _finish(job, report, label: str) -> int:
    """결과 요약을 화면 · 실행기록.txt 에 남기고, 알릴 것이 있으면 알린다(잠금은 이미 풀림)."""
    print(report.message)
    _append_run_record(job, label, report.message)
    if report.has_news:
        notify_report(job, report)
    return 1 if report.stopped else 0


def _append_record(folder, label: str, message: str):
    stamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    try:
        with open(Path(folder) / RUN_RECORD, "a", encoding="utf-8") as f:
            f.write(f"[{stamp}] {label}\n{message}\n\n")
    except OSError as e:
        print(f"경고: {RUN_RECORD} 에 쓰지 못함 ({e})")


def _append_run_record(job, label: str, message: str):
    _append_record(job.folder, label, message)


# --- 자동 실행(매일 실행 · 놓친 실행): 아무도 안 보므로 실패는 기록하고 알린다 ---

def _automatic(args, repo_root) -> int:
    """run N --자동 · catch-up. 작업을 찾기 전에 실패해도(작업폴더.txt 없음 · 작업 폴더가 동기화 폴더가 됨 ·
    작업 설정이 깨짐) 조용히 끝나지 않게 실행기록.txt 에 남기고 알린다. 창 없이(pythonw) 돌므로 화면 출력은 안 보인다."""
    what = "놓친 작업 실행" if args.command == "catch-up" else f"작업 {args.job} 자동 실행"
    log_dir = Path(repo_root)  # 작업 폴더를 모르거나 쓸 수 없으면(동기화 폴더) 저장소에 남긴다
    try:
        work = config.work_dir(repo_root)
        check_work_dir(work)
        log_dir = work
        if args.command == "catch-up":
            return _catch_up(work, what)
        return _auto_run(work, args.job)
    except config.ConfigError as e:
        return _auto_failed(log_dir, what, f"{e}. 작업 폴더 설정(작업폴더.txt)을 찾지 못했습니다.")
    except LocalOnlyError as e:
        return _auto_failed(log_dir, what, f"작업 폴더를 쓸 수 없어 실행하지 않았습니다. {e}")
    except JobError as e:
        return _auto_failed(log_dir, what, f"{e}\n{JOB_FIX_HINT}")
    except Exception as e:  # 창이 없으니 예상 못 한 오류도 기록 · 알림으로 남긴다
        detail = "".join(traceback.format_exception(e))
        return _auto_failed(log_dir, what, f"예상하지 못한 오류로 멈췄습니다: {e}", detail)


def _auto_failed(log_dir, what: str, message: str, detail: str = "") -> int:
    """기록하고 알린다. 알림 글에 기록 파일 위치를 넣고, 누르면 그 폴더가 열린다."""
    print(message)
    record = Path(log_dir) / RUN_RECORD
    _append_record(log_dir, f"{what} 실패", message + (f"\n{detail}" if detail else ""))
    notify_failure(what, f"{message}\n기록: {record}", folder=Path(log_dir))
    return 1


def _auto_run(work, key) -> int:
    job = find_job(work, key)
    if job.paused:
        print(f"{job.name}: 일시중지 중이라 실행하지 않았습니다.")
        return 0
    try:
        report = run_job(job, work, gate=approval_gate)
    except JobLocked as e:  # 지금 실행 등이 돌고 있음: 그 실행이 오늘 몫을 처리한다
        print(e)
        _append_run_record(job, "자동 실행", f"건너뜀: {e}")
        return 1
    return _finish(job, report, "자동 실행")


def _catch_up(work, what: str) -> int:
    """오늘 시각이 지났는데 아직 안 돈 작업을 차례로 실행한다. 알림은 모두 끝난 뒤에 띄운다
    (메시지 창으로 대신할 때 창이 다음 작업 실행을 막지 않게)."""
    finished, failed = [], []

    def runner(job):
        try:
            report = run_job(job, work, gate=approval_gate)
        except (JobLocked, JobError):
            raise  # 실행 중 · 그새 지워진 작업: schedule.catch_up 이 건너뛴다
        except Exception as e:  # 한 작업의 예상 못 한 오류가 다음 작업과 앞 작업의 알림을 막지 않게
            detail = "".join(traceback.format_exception(e))
            print(f"{job.name}: 예상하지 못한 오류로 멈췄습니다: {e}")
            _append_run_record(job, "놓친 실행 실패", f"예상하지 못한 오류로 멈췄습니다: {e}\n{detail}")
            failed.append(f"{job.name}: 예상하지 못한 오류로 멈췄습니다: {e} (기록: {Path(job.folder) / RUN_RECORD})")
            return
        print(report.message)
        _append_run_record(job, "놓친 실행", report.message)
        finished.append((job, report))

    ran = schedule.catch_up(work, runner=runner)
    for job, report in finished:
        if report.has_news:
            notify_report(job, report)
    if not ran:
        print("지금 실행할 놓친 작업이 없습니다.")
    problems = list(failed)
    broken = broken_jobs(work)
    if broken:
        lines = "\n".join(f"{folder.name}: {reason}" for folder, reason in broken)
        problems.append(f"설정이 깨진 작업이 있어 실행하지 못했습니다.\n{lines}\n{JOB_FIX_HINT}")
    if problems:  # 결과 알림을 먼저 띄운 뒤, 실패는 한 번에 모아 알린다
        return _auto_failed(work, what, "\n".join(problems))
    return 1 if any(report.stopped for _, report in finished) else 0


# --- 폴더 고르기 ---

def _pick_folder(title: str) -> int:
    """폴더 고르기 창. 마지막 줄 = 고른 경로(로컬 디스크일 때만). 취소면 「취소함」."""
    chosen = init_module.pick_folder(title)
    if not chosen:
        print(CANCELLED)
        return 1
    path = Path(chosen)
    try:
        check_work_dir(path)
    except LocalOnlyError as e:
        print(e)
        print("이 폴더는 쓸 수 없습니다. 내 PC 의 로컬 디스크 폴더를 다시 고르세요.")
        return 1
    print(path)
    return 0


# --- check ---

def _check(work, args) -> int:
    if args.job:
        job = find_job(work, args.job)
    else:
        jobs = list_jobs(work)
        if len(jobs) != 1:
            names = ", ".join(j.name for j in jobs) or "없음"
            print(f"어느 작업의 규칙으로 볼지 --작업 번호를 적어 주세요. 지금 있는 작업: {names}")
            return 1
        job = jobs[0]
    rules = load_rules(job.rules_path)
    return check(Path(args.target), rules, work, job)


def _show(value):
    if value is None:
        return "(비어 있음)"
    if isinstance(value, date):
        return value.isoformat()
    return str(value)


def _cell(text):
    return str(text).replace("|", "\\|")


def check(path: Path, rules, work: Path | None = None, job=None) -> int:
    """추출 결과 · 검산 결과를 표로 보여 준다. 파일을 만들거나 옮기지 않는다."""
    if not path.is_file():
        print(f"파일이 없습니다: {path}")
        return 1
    if path.suffix.lower() == ".xls":
        print(XLS_REASON)
        return 0
    result = read_excel(path, rules) if path.suffix.lower() == ".xlsx" else extract(path, rules)
    print(f"파일: {path.name}")
    if job is not None:
        print(f"작업: {job.name} (출력 방식 {job.output_mode})")
    print(f"문서종류: {result.doc_type.name if result.doc_type else '(판별 실패)'}")
    if result.source_form:
        print(f"엑셀양식: {result.source_form}")

    if result.review_reason and (result.doc_type is None or not result.has_text):
        print(f"\n확인필요: {result.review_reason}")
        if result.text_head:
            print("\n--- 앞부분 글자(30줄) ---")
            print(result.text_head)
        return 0

    print(f"번호필드: {result.doc_type.number_field}")
    print("\n| 필드 | 값 | 비고 |\n|---|---|---|")
    for name, value in result.fields.items():
        notes = []
        if value is None:
            notes.append("못 찾음")
        notes += [w for w in result.warnings if f"'{name}' 값" in w]
        print(f"| {_cell(name)} | {_cell(_show(value))} | {_cell(', '.join(notes))} |")

    columns = _item_columns(result, rules)
    if columns:
        if result.table is None:
            print("\n품목 표: 읽지 못함 (아래 경고 참고)")
        else:
            print(f"\n품목 {len(result.table)}줄")
            print("| " + " | ".join(columns) + " |")
            print("|" + "---|" * len(columns))
            for row in result.table:
                print("| " + " | ".join(_cell(_show(row.get(c, ""))) for c in columns) + " |")

    others = [w for w in result.warnings if "' 못 찾음" not in w and "' 값 " not in w]
    for w in others:
        print(f"경고: {w}")
    if result.review_reason:
        print(f"\n확인필요: {result.review_reason}")
        return 0

    bad = _print_checks(result, path)
    if bad:
        print(f"\n확인필요: 검산 불일치 {bad}건. 이대로 실행하면 이 문서는 결과를 내지 않고 확인필요로 갑니다.")
    if job is not None and job.output_mode in LEDGER_MODES:
        print(f"누적장부: {ledger_path(job).name}")
    if rules.company_form is not None and (job is None or job.output_mode in COMPANY_MODES):
        _check_company_form(result, rules, work)
    return 0


def _print_checks(result, path) -> int:
    """검산 결과 표(실행 때와 같은 검산). 불일치 건수를 돌려준다."""
    try:
        checks = verify(result, path)
    except Exception as e:
        print(f"\n확인필요: 검산을 하지 못함 ({e}). 이대로 실행하면 이 문서는 확인필요로 갑니다.")
        return 0
    print("\n| 검산 | 원본 | 결과 | 일치 |\n|---|---|---|---|")
    for c in checks:
        print(f"| {_cell(c.item)} | {_cell(c.source)} | {_cell(c.output)} | {_mark(c.ok)} |")
    return len(failed(checks))


def _item_columns(result, rules):
    if result.source_form:
        form = next(f for f in rules.excel_forms if f.name == result.source_form)
        return list(form.items.columns) if form.items else []
    return result.doc_type.table.columns if result.doc_type.table else []


def _check_company_form(result, rules, work):
    """회사 양식: 고객마스터에서 고객을 찾았는지(양식이 고객마스터 값을 쓸 때), 마크업(쓸 때), 몇 줄로 만들어질지."""
    form = rules.company_form
    row = None
    if any(c.source.startswith(MASTER_PREFIX) for c in form.columns):
        customer = result.fields.get("고객명")
        try:
            row = lookup(load_master(Path(work) / config.MASTER_FILE), customer) if work else None
        except MasterError as e:
            print(f"고객마스터: {e}")
        if row is not None:
            print(f"고객마스터: 찾음 (고객코드 {row.get('고객코드')})")
        else:
            print(f"고객마스터: 없음 ('{customer}' 줄을 고객마스터.xlsx 에 더해야 회사 양식이 만들어집니다)")
    if form.uses_markup:
        _check_markup(result, form, row, work)
    print(f"회사 양식: {company_form_name(result, form)} ({len(result.table or [])}줄)")


def _check_markup(result, form, master_row, work):
    """마크업표: 빠진 파트가 있는지, 줄마다 마크업% · 본사 단가 · 본사 금액(실행 때와 같은 계산 · 검산)."""
    try:
        if not work:
            raise MarkupError(f"작업 폴더를 몰라 {MARKUP_FILE} 를 읽지 못했습니다")
        table = load_markup(Path(work) / MARKUP_FILE)
    except MarkupError as e:
        print(f"\n마크업표: {e}")
        print("이대로 실행하면 이 문서는 확인필요로 갑니다.")
        return
    missing = missing_parts(result, form, table)
    if missing:
        print(f"\n마크업표: 없는 파트넘버 {', '.join(missing)} ({MARKUP_FILE} 에 더해야 회사 양식이 만들어집니다. "
              "이대로 실행하면 확인필요)")
        return
    _, rows = company_rows(result, master_row, form, table)
    checks = markup_checks(result, written_markup(form, rows), table, form.markup)
    print("\n| 마크업 검산 | 계산 | 결과 | 일치 |\n|---|---|---|---|")
    for c in checks:
        print(f"| {_cell(c.item)} | {_cell(c.source)} | {_cell(c.output)} | {_mark(c.ok)} |")
    bad = len(failed(checks))
    if bad:
        print(f"\n확인필요: 마크업 검산 불일치 {bad}건. 이대로 실행하면 이 문서는 확인필요로 갑니다.")


if __name__ == "__main__":
    sys.exit(main())
