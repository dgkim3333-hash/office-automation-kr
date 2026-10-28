"""매일 실행 등록(작업 스케줄러 → 시작프로그램 → 없음) · 놓친 실행 · 작업 명령의 등록 정리 · 자동 실행 실패 알림.
실제 작업 스케줄러 · 바탕화면은 건드리지 않는다(가짜 실행기 · 임시 폴더)."""
import os
import shutil
from datetime import datetime
from pathlib import Path

import pytest
import yaml

import app.__main__ as cli
from app import config, schedule
from app.__main__ import main
from app.job_run import JobLocked, job_lock
from app.jobs import JOB_FILE, JobError, create_job, find_job, load_job, save_job

REPO = Path(__file__).resolve().parent.parent
EXAMPLE_RULES = REPO / "rules" / "규칙.예시.yaml"
TITLE = "일일 PO 정리"


def _job(work, inp, out, title=TITLE, number_time=None):
    job = create_job(work, title, inp, out, "누적장부")
    if number_time:
        job.run_time = number_time
        save_job(job)
    return job


@pytest.fixture
def dirs(tmp_path):
    inp, out = tmp_path / "in", tmp_path / "out"
    inp.mkdir()
    out.mkdir()
    return tmp_path / "work", inp, out


# --- register ---

def test_register_uses_task_scheduler_when_query_finds_task(fake_ps, tmp_path, dirs):
    work, inp, out = dirs
    _job(work, inp, out, title="첫 작업")
    job = _job(work, inp, out)  # 2번 = 16:20
    repo = tmp_path / "repo"

    assert schedule.register(job, repo) == "작업스케줄러"
    (script,) = fake_ps.having("Register-ScheduledTask")
    assert "-StartWhenAvailable" in script and "16:20" in script
    assert "-Daily" in script and "-LogonType Interactive" in script
    assert "'\\사무자동화\\'" in script and f"'02_{TITLE}'" in script
    assert f"'{repo / '.venv' / 'Scripts' / 'pythonw.exe'}'" in script
    assert "'-m app run 02 --자동'" in script and f"-WorkingDirectory '{repo}'" in script
    assert fake_ps.having("Get-ScheduledTask -TaskPath")  # 등록 뒤 다시 조회해 확인
    assert not fake_ps.having("CreateShortcut")
    assert load_job(job.folder).scheduler == "작업스케줄러"


def test_register_falls_back_when_query_does_not_find_task(fake_ps, tmp_path, dirs):
    job = _job(*dirs)
    startup = tmp_path / "startup"
    startup.mkdir()
    fake_ps.fail.append("Get-ScheduledTask -TaskPath")  # 등록 명령은 성공했다고 했지만 조회하면 없음
    assert schedule.register(job, tmp_path / "repo", startup=startup) == "시작프로그램"


def test_register_falls_back_to_startup_when_blocked(fake_ps, tmp_path, dirs):
    job = _job(*dirs)
    repo, startup = tmp_path / "repo", tmp_path / "startup"
    startup.mkdir()
    fake_ps.fail.append("Register-ScheduledTask")
    notes = []

    assert schedule.register(job, repo, startup=startup, notes=notes) == "시작프로그램"
    link = startup / "사무자동화 놓친 작업 실행.lnk"
    made = fake_ps.having("CreateShortcut")
    assert len(made) == 1 and f"'{link}'" in made[0] and "'-m app catch-up'" in made[0]
    assert notes and "작업 스케줄러" in notes[0]
    assert load_job(job.folder).scheduler == "시작프로그램"


def test_register_keeps_existing_startup_link(fake_ps, tmp_path, dirs):
    job = _job(*dirs)
    startup = tmp_path / "startup"
    startup.mkdir()
    (startup / "사무자동화 놓친 작업 실행.lnk").write_bytes(b"lnk")
    fake_ps.fail.append("Register-ScheduledTask")
    assert schedule.register(job, tmp_path / "repo", startup=startup) == "시작프로그램"
    assert not fake_ps.having("CreateShortcut")


def test_register_none_when_everything_blocked(fake_ps, tmp_path, dirs):
    job = _job(*dirs)
    fake_ps.fail += ["Register-ScheduledTask", "GetFolderPath", "CreateShortcut"]
    assert schedule.register(job, tmp_path / "repo") == "없음"
    assert load_job(job.folder).scheduler == "없음"


def test_register_paused_job_disables_task(fake_ps, tmp_path, dirs):
    job = _job(*dirs)
    job.paused = True
    assert schedule.register(job, tmp_path / "repo") == "작업스케줄러"
    assert fake_ps.having("Disable-ScheduledTask")


def test_unregister_pause_resume_scripts(fake_ps, dirs):
    job = _job(*dirs)
    schedule.unregister(job)
    schedule.pause(job)
    schedule.resume(job)
    assert "Unregister-ScheduledTask" in fake_ps.scripts[0] and f"'01_{TITLE}'" in fake_ps.scripts[0]
    assert "Disable-ScheduledTask" in fake_ps.scripts[1]
    assert "Enable-ScheduledTask" in fake_ps.scripts[2]


def test_unregister_failure_raises(fake_ps, dirs):
    job = _job(*dirs)
    fake_ps.fail.append("Unregister-ScheduledTask")
    with pytest.raises(schedule.ScheduleError):
        schedule.unregister(job)


# --- 놓친 실행 ---

def test_due_jobs_skips_done_paused_and_future(dirs):
    work, inp, out = dirs
    now = datetime(2026, 10, 28, 16, 50)
    due = _job(work, inp, out, "가", "16:00")
    done = _job(work, inp, out, "나", "16:00")
    done.last_run = "2026-10-28"
    paused = _job(work, inp, out, "다", "16:20")
    paused.paused = True
    future = _job(work, inp, out, "라", "17:00")
    yesterday = _job(work, inp, out, "마", "16:40")
    yesterday.last_run = "2026-10-27"
    exact = _job(work, inp, out, "바", "16:50")

    got = schedule.due_jobs([exact, yesterday, future, paused, done, due], now)
    assert [j.title for j in got] == ["가", "마", "바"]


def test_catch_up_runs_missed_jobs_in_order_once(dirs):
    work, inp, out = dirs
    _job(work, inp, out, "가", "16:00")
    _job(work, inp, out, "나", "16:20")
    _job(work, inp, out, "다", "17:00")
    now = datetime(2026, 10, 28, 16, 50)
    calls = []

    def runner(job):  # 실제 실행처럼 마지막 실행 날짜를 남긴다
        calls.append(job.number)
        job.last_run = now.date().isoformat()
        save_job(job)

    ran = schedule.catch_up(work, now=now, runner=runner)
    assert calls == [1, 2] and len(ran) == 2
    assert schedule.catch_up(work, now=now, runner=runner) == []
    assert calls == [1, 2]


def test_catch_up_skips_locked_and_deleted_jobs(dirs, monkeypatch):
    """잠긴 작업은 기다렸다가 다시 해 보고(여기서는 한 번), 끝내 잠겨 있으면 locked 로 돌려준다. 지워진 작업은 건너뛴다."""
    monkeypatch.setattr(schedule, "LOCK_RETRY_LIMIT_SECONDS", schedule.LOCK_RETRY_SECONDS)
    work, inp, out = dirs
    _job(work, inp, out, "가", "16:00")
    _job(work, inp, out, "나", "16:00")
    _job(work, inp, out, "다", "16:00")
    now = datetime(2026, 10, 28, 16, 50)
    calls = []

    def runner(job):
        calls.append(job.number)
        if job.number == 1:
            raise JobLocked("이미 실행 중")
        if job.number == 2:
            raise JobError("작업이 지워졌습니다")

    locked = []
    assert [name[:2] for name in schedule.catch_up(work, now=now, runner=runner, locked=locked)] == ["03"]
    assert calls == [1, 1, 2, 3]
    assert [job.number for job in locked] == [1]


def test_catch_up_rechecks_job_on_disk_before_running(dirs):
    """목록을 읽은 뒤 작업 스케줄러 쪽 실행이 먼저 끝났으면(마지막 실행 = 오늘) 다시 돌리지 않는다."""
    work, inp, out = dirs
    first = _job(work, inp, out, "가", "16:00")
    second = _job(work, inp, out, "나", "16:00")
    now = datetime(2026, 10, 28, 16, 50)
    calls = []

    def runner(job):
        calls.append(job.number)
        other = load_job(second.folder)
        other.last_run = "2026-10-28"
        save_job(other)

    schedule.catch_up(work, now=now, runner=runner)
    assert calls == [first.number]


def test_job_lock_on_deleted_job_is_job_error(dirs):
    job = _job(*dirs)
    shutil.rmtree(job.folder)
    with pytest.raises(JobError, match="작업이 지워졌습니다"):
        with job_lock(job):
            pass


# --- 명령 ---

@pytest.fixture
def env(tmp_path, capsys):
    repo = tmp_path / "repo"
    shutil.copytree(REPO / "rules", repo / "rules")
    work, inp, out = tmp_path / "work", tmp_path / "in", tmp_path / "out"
    inp.mkdir()
    out.mkdir()
    assert main(["init", str(work)], repo_root=repo) == 0
    capsys.readouterr()
    return repo, work.resolve(), inp, out


def _new(repo, inp, out, title=TITLE, resume=True):
    """job new(멈춤으로 만들어짐) 뒤, 작업지시서처럼 규칙을 채웠다고 보고 job resume."""
    assert main(["job", "new", "--제목", title, "--입력", str(inp), "--출력", str(out), "--방식", "누적장부"],
                repo_root=repo) == 0
    job = find_job(config.work_dir(repo), title)
    if resume:
        assert main(["job", "resume", str(job.number)], repo_root=repo) == 0
        job = load_job(job.folder)
    return job


def _desktop(fake_ps, tmp_path):
    desk = tmp_path / "desk"
    desk.mkdir(exist_ok=True)
    fake_ps.outputs["'Desktop'"] = str(desk)
    fake_ps.outputs["'Startup'"] = str(tmp_path / "startup")
    return desk


def test_job_new_prints_which_method_was_used(env, fake_ps, tmp_path, capsys):
    repo, work, inp, out = env
    desk = _desktop(fake_ps, tmp_path)
    job = _new(repo, inp, out)
    shown = capsys.readouterr().out
    assert "자동 실행: Windows 작업 스케줄러 (매일 16:00)" in shown
    assert "지금 실행 - 일일 PO 정리.lnk" in shown and "승인 - 일일 PO 정리.lnk" in shown
    assert job.scheduler == "작업스케줄러"
    assert len(fake_ps.having(f"'{desk / '지금 실행 - 일일 PO 정리.lnk'}'")) == 1


def test_job_new_prints_fallback_method(env, fake_ps, tmp_path, capsys):
    repo, work, inp, out = env
    _desktop(fake_ps, tmp_path)
    fake_ps.fail.append("Register-ScheduledTask")
    job = _new(repo, inp, out)
    shown = capsys.readouterr().out
    assert "자동 실행: 시작프로그램" in shown and "16:00" in shown
    assert job.scheduler == "시작프로그램"


def test_job_new_when_everything_blocked_says_use_desktop(env, capsys):
    repo, work, inp, out = env  # 기본 가짜 실행기는 모든 PowerShell 을 실패시킨다
    job = _new(repo, inp, out)
    shown = capsys.readouterr().out
    assert "자동 실행: 없음" in shown and "지금 실행" in shown
    assert "경고: 바탕화면 바로가기를 만들지 못함" in shown and "Traceback" not in shown
    assert job.scheduler == "없음"


def test_job_delete_unregisters_and_removes_shortcuts(env, fake_ps, tmp_path, capsys):
    repo, work, inp, out = env
    desk = _desktop(fake_ps, tmp_path)
    job = _new(repo, inp, out)
    links = [desk / "지금 실행 - 일일 PO 정리.lnk", desk / "승인 - 일일 PO 정리.lnk"]
    for p in links:
        p.write_bytes(b"lnk")
    assert main(["job", "delete", "1"], repo_root=repo) == 0
    assert not job.folder.exists()
    (script,) = fake_ps.having("Unregister-ScheduledTask")
    assert f"'01_{TITLE}'" in script
    assert not any(p.exists() for p in links)


def test_job_delete_warns_when_unregister_fails(env, fake_ps, tmp_path, capsys):
    repo, work, inp, out = env
    _desktop(fake_ps, tmp_path)
    job = _new(repo, inp, out)
    fake_ps.fail.append("Unregister-ScheduledTask")
    capsys.readouterr()
    assert main(["job", "delete", "1"], repo_root=repo) == 0
    shown = capsys.readouterr().out
    assert not job.folder.exists()
    assert "사무자동화\\01_일일 PO 정리" in shown and "다시 등록" in shown


def test_pause_resume_disable_and_enable_task(env, fake_ps, tmp_path):
    repo, work, inp, out = env
    _desktop(fake_ps, tmp_path)
    job = _new(repo, inp, out, resume=False)
    # 새 작업은 멈춤으로 만들어져 등록 직후 작업 스케줄러에서도 꺼 둔다(Ruling 26)
    assert load_job(job.folder).paused and len(fake_ps.having("Disable-ScheduledTask")) == 1
    assert main(["job", "resume", "1"], repo_root=repo) == 0
    assert not load_job(job.folder).paused and len(fake_ps.having("Enable-ScheduledTask")) == 1
    assert main(["job", "pause", "1"], repo_root=repo) == 0
    assert load_job(job.folder).paused and len(fake_ps.having("Disable-ScheduledTask")) == 2


def test_pause_without_task_scheduler_does_not_call_it(env, capsys):
    repo, work, inp, out = env
    job = _new(repo, inp, out)  # 모두 막혀 「없음」
    assert main(["job", "pause", "1"], repo_root=repo) == 0
    assert load_job(job.folder).paused


def test_job_repair_registers_again_and_removes_leftover_tasks(env, fake_ps, tmp_path, capsys):
    repo, work, inp, out = env
    fake_ps.fail += ["Register-ScheduledTask", "GetFolderPath"]
    job = _new(repo, inp, out)  # 만들 때는 모두 막힘 → 없음
    assert job.scheduler == "없음"
    fake_ps.fail.clear()  # 그 뒤 막힘이 풀림(또는 다른 PC 로 옮김)
    desk = _desktop(fake_ps, tmp_path)
    fake_ps.outputs["ForEach-Object"] = f"01_{TITLE}\r\n09_지운 작업\r\n"
    capsys.readouterr()

    assert main(["job", "repair"], repo_root=repo) == 0
    shown = capsys.readouterr().out
    assert load_job(job.folder).scheduler == "작업스케줄러"
    assert "자동 실행: Windows 작업 스케줄러 (매일 16:00)" in shown
    assert fake_ps.having(f"'{desk / '승인 - 일일 PO 정리.lnk'}'")
    removed = fake_ps.having("Unregister-ScheduledTask")
    assert len(removed) == 1 and "'09_지운 작업'" in removed[0]
    assert "09_지운 작업" in shown


def test_catch_up_command_runs_missed_jobs_and_notifies_after(env, monkeypatch, capsys):
    repo, work, inp, out = env
    job = _new(repo, inp, out)
    shutil.copy(EXAMPLE_RULES, job.rules_path)
    shutil.copy(REPO / "samples" / "03_Invoice_DM_PDF" / "INV-2026-10231.pdf", inp)
    t = datetime.now().timestamp() - 3600
    os.utime(inp / "INV-2026-10231.pdf", (t, t))
    job.run_time = "00:00"
    save_job(job)
    calls = []
    monkeypatch.setattr(cli, "notify_report", lambda job, report: calls.append(report.counts["승인대기"]))

    assert main(["catch-up"], repo_root=repo) == 0
    assert calls == [1]
    assert load_job(job.folder).last_run == datetime.now().date().isoformat()
    assert "놓친 실행" in (job.folder / "실행기록.txt").read_text(encoding="utf-8")
    assert main(["catch-up"], repo_root=repo) == 0  # 오늘 이미 돌았으면 다시 돌리지 않는다
    assert calls == [1]


# --- 자동 실행이 작업을 찾기 전에 실패하면 조용히 넘어가지 않는다 ---

@pytest.fixture
def failures(monkeypatch):
    calls = []
    monkeypatch.setattr(cli, "notify_failure",
                        lambda what, message, folder=None: calls.append((what, message, folder)))
    return calls


@pytest.mark.parametrize("command", [["run", "1", "--자동"], ["catch-up"]])
def test_auto_run_without_work_folder_setting_logs_and_notifies(tmp_path, failures, capsys, command):
    repo = tmp_path / "repo"
    repo.mkdir()  # 작업폴더.txt 없음
    assert main(command, repo_root=repo) != 0
    record = (repo / "실행기록.txt").read_text(encoding="utf-8")
    assert "설치.bat" in record
    assert len(failures) == 1 and "설치.bat" in failures[0][1]
    assert failures[0][2] == repo and str(repo / "실행기록.txt") in failures[0][1]  # 누르면 기록 폴더, 글에 기록 위치
    assert "Traceback" not in capsys.readouterr().out


@pytest.mark.parametrize("command", [["run", "1", "--자동"], ["catch-up"]])
def test_auto_run_when_work_folder_became_sync_folder(tmp_path, failures, command):
    repo, synced = tmp_path / "repo", tmp_path / "OneDrive" / "사무자동화"
    repo.mkdir()
    synced.mkdir(parents=True)
    (repo / config.WORK_DIR_FILE).write_text(str(synced), encoding="utf-8")
    assert main(command, repo_root=repo) != 0
    assert "동기화 폴더" in (repo / "실행기록.txt").read_text(encoding="utf-8")
    assert not (synced / "실행기록.txt").exists()  # 동기화 폴더에는 쓰지 않는다
    assert len(failures) == 1 and "동기화 폴더" in failures[0][1]


def test_auto_run_with_broken_job_settings_logs_and_notifies(env, failures):
    repo, work, inp, out = env
    job = _new(repo, inp, out)
    (job.folder / JOB_FILE).write_text("번호: [깨짐", encoding="utf-8")
    assert main(["run", "1", "--자동"], repo_root=repo) != 0
    record = (work / "실행기록.txt").read_text(encoding="utf-8")
    assert "설정이 깨져" in record
    assert len(failures) == 1 and "설정이 깨져" in failures[0][1] and "Code 탭" in failures[0][1]
    assert failures[0][2] == work and str(work / "실행기록.txt") in failures[0][1]


def test_auto_run_of_deleted_job_logs_and_notifies(env, failures):
    repo, work, inp, out = env
    assert main(["run", "3", "--자동"], repo_root=repo) != 0
    assert "찾지 못했습니다" in (work / "실행기록.txt").read_text(encoding="utf-8")
    assert len(failures) == 1


def test_catch_up_with_broken_job_notifies(env, failures):
    repo, work, inp, out = env
    job = _new(repo, inp, out)
    (job.folder / JOB_FILE).write_text(yaml.safe_dump({"번호": 1}), encoding="utf-8")
    assert main(["catch-up"], repo_root=repo) != 0
    assert len(failures) == 1 and job.folder.name in failures[0][1]


def test_manual_run_failure_does_not_use_failure_notification(env, failures, capsys):
    repo, *_ = env
    assert main(["run", "7"], repo_root=repo) != 0
    assert failures == []


def test_auto_run_unexpected_error_logs_and_notifies(env, failures, monkeypatch):
    repo, work, inp, out = env
    _new(repo, inp, out)

    def boom(*a, **k):
        raise RuntimeError("예상 못 한 문제")

    monkeypatch.setattr(cli, "run_job", boom)
    assert main(["run", "1", "--자동"], repo_root=repo) != 0
    assert "예상 못 한 문제" in (work / "실행기록.txt").read_text(encoding="utf-8")
    assert len(failures) == 1


def test_catch_up_keeps_going_after_unexpected_error(env, failures, monkeypatch):
    """한 작업이 예상 못 한 오류로 멈춰도 다음 작업을 돌리고, 결과 알림과 실패 알림을 둘 다 띄운다."""
    repo, work, inp, out = env
    first = _new(repo, inp, out, title="가")
    second = _new(repo, inp, out, title="나")
    for job in (first, second):
        job.run_time = "00:00"
        save_job(job)
    from app.job_run import RunReport
    ran, notified = [], []

    def fake_run(job, work, **kw):
        ran.append(job.number)
        if job.number == 1:
            raise RuntimeError("예상 못 한 문제")
        job.last_run = datetime.now().date().isoformat()
        save_job(job)
        report = RunReport(message="나: 성공 1")
        report.counts["성공"] = 1
        return report

    monkeypatch.setattr(cli, "run_job", fake_run)
    monkeypatch.setattr(cli, "notify_report", lambda job, report: notified.append(job.number))
    assert main(["catch-up"], repo_root=repo) != 0
    assert ran == [1, 2]
    assert notified == [2]
    assert len(failures) == 1 and "01_가" in failures[0][1] and "예상 못 한 문제" in failures[0][1]
    assert "예상 못 한 문제" in (first.folder / "실행기록.txt").read_text(encoding="utf-8")
