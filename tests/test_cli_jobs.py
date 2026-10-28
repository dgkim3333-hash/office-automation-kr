"""명령(python -m app ...): 작업 만들기 · 목록 · 실행 · 승인 · check · 폴더 고르기. 옛 받은문서 흐름은 없다."""
import os
import shutil
import time
from pathlib import Path

import pytest
import yaml

import app.__main__ as cli
import app.init as init_module
from app import config
from app.__main__ import main
from app.job_run import job_lock, load_records
from app.jobs import JOB_FILE, find_job, load_job
from app.rules import RuleError, load_rules

REPO = Path(__file__).resolve().parent.parent
PDFS = REPO / "samples" / "03_Invoice_DM_PDF"
ORDERS = REPO / "samples" / "01_고객오더_엑셀"
EXAMPLE_RULES = REPO / "rules" / "규칙.예시.yaml"
SKELETON = REPO / "rules" / "규칙.뼈대.yaml"
INV = PDFS / "INV-2026-10231.pdf"
ACME_PO = ORDERS / "고객오더_ACME_PO-ACME-2026-0417.xlsx"
TITLE = "일일 PO 정리"


@pytest.fixture
def env(tmp_path, capsys):
    """가짜 저장소(rules 만) + init 한 작업 폴더 + 입력 · 출력 폴더."""
    repo = tmp_path / "repo"
    shutil.copytree(REPO / "rules", repo / "rules")
    work, inp, out = tmp_path / "work", tmp_path / "in", tmp_path / "out"
    inp.mkdir()
    out.mkdir()
    assert main(["init", str(work)], repo_root=repo) == 0
    capsys.readouterr()
    return repo, work.resolve(), inp, out


def _new(repo, inp, out, title=TITLE, mode="누적장부", rules=EXAMPLE_RULES):
    assert main(["job", "new", "--제목", title, "--입력", str(inp), "--출력", str(out), "--방식", mode],
                repo_root=repo) == 0
    job = find_job(config.work_dir(repo), title)
    if rules is not None:
        shutil.copy(rules, job.rules_path)
    return job


def _drop(src, folder, age=3600):
    dst = Path(folder) / Path(src).name
    shutil.copy(src, dst)
    t = time.time() - age
    os.utime(dst, (t, t))
    return dst


def _snapshot(folder):
    return sorted(str(p.relative_to(folder)) for p in Path(folder).rglob("*"))


# --- job new · job list ---

def test_job_new_then_list(env, capsys):
    repo, work, inp, out = env
    code = main(["job", "new", "--제목", TITLE, "--입력", str(inp), "--출력", str(out), "--방식", "둘다"],
                repo_root=repo)
    lines = capsys.readouterr().out.strip().splitlines()
    assert code == 0
    assert f"작업 만듦: 01_{TITLE} (매일 16:00)" in lines  # 그 뒤에 자동 실행 방법 · 바로가기 줄이 온다(W7)
    job = find_job(work, "1")
    assert job.output_mode == "둘다" and job.input_dir == inp.resolve()
    # 예시 규칙이 아니라 Claude 가 채울 빈 뼈대. 아직 규칙으로 읽히지 않아도 작업은 만들어진다
    assert job.rules_path.read_bytes() == SKELETON.read_bytes()
    with pytest.raises(RuleError):
        load_rules(job.rules_path)

    assert main(["job", "list"], repo_root=repo) == 0
    listed = capsys.readouterr().out
    assert f"01_{TITLE}" in listed and "16:00" in listed


def test_job_list_shows_scheduler_and_approvals(env, capsys):
    repo, work, inp, out = env
    job = _new(repo, inp, out)
    job.approvals = {"ACME": 3, "Invoice": 10}
    job.scheduler = "작업스케줄러"
    from app.jobs import save_job
    save_job(job)
    capsys.readouterr()
    main(["job", "list"], repo_root=repo)
    listed = capsys.readouterr().out
    assert "작업스케줄러" in listed
    assert "ACME 3/10" in listed and "Invoice 10/10" in listed


def test_job_list_shows_broken_jobs_with_reason(env, capsys):
    repo, work, inp, out = env
    _new(repo, inp, out)
    broken = work / "작업" / "02_깨진 작업"
    broken.mkdir()
    (broken / JOB_FILE).write_text("번호: [", encoding="utf-8")
    capsys.readouterr()
    assert main(["job", "list"], repo_root=repo) == 0
    listed = capsys.readouterr().out
    assert f"01_{TITLE}" in listed
    assert "02_깨진 작업" in listed and "작업 설정을 읽지 못했습니다" in listed


def test_job_new_rejects_bad_folder_in_korean(env, capsys):
    repo, work, inp, out = env
    code = main(["job", "new", "--제목", TITLE, "--입력", str(inp), "--출력", str(inp), "--방식", "둘다"],
                repo_root=repo)
    assert code != 0
    assert "서로 다른 폴더" in capsys.readouterr().out
    assert not (work / "작업" / f"01_{TITLE}").exists()


def test_job_pause_resume_and_delete(env, capsys):
    repo, work, inp, out = env
    job = _new(repo, inp, out)
    assert main(["job", "pause", "1"], repo_root=repo) == 0
    assert load_job(job.folder).paused is True
    assert main(["job", "resume", "1"], repo_root=repo) == 0
    assert load_job(job.folder).paused is False

    with job_lock(job):  # 실행 중이면 지우지 않는다
        assert main(["job", "delete", "1"], repo_root=repo) != 0
    assert job.folder.is_dir()
    assert "실행 중" in capsys.readouterr().out

    assert main(["job", "delete", "1"], repo_root=repo) == 0
    assert not job.folder.exists()
    capsys.readouterr()
    assert main(["job", "list"], repo_root=repo) == 0
    assert f"01_{TITLE}" not in capsys.readouterr().out
    # 다시 만들면 지운 번호를 다시 쓰지 않는다
    again = _new(repo, inp, out, title="새 작업")
    assert again.number == 2


def test_pause_while_running_is_refused(env, capsys):
    repo, work, inp, out = env
    job = _new(repo, inp, out)
    with job_lock(job):
        assert main(["job", "pause", "1"], repo_root=repo) != 0
    assert "실행 중" in capsys.readouterr().out
    assert load_job(job.folder).paused is False


# --- 인자 없음 · init · 옛 흐름 ---

def test_no_args_shows_jobs_not_inbox(env, capsys):
    repo, work, inp, out = env
    _new(repo, inp, out)
    _drop(INV, inp)
    capsys.readouterr()
    assert main([], repo_root=repo) == 0
    shown = capsys.readouterr().out
    assert f"01_{TITLE}" in shown
    assert "지금 실행 - 작업 이름" in shown
    assert "받은문서" not in shown
    assert not (work / "작업" / f"01_{TITLE}" / "처리기록.csv").exists()  # 아무것도 처리하지 않음


def test_no_args_without_jobs_explains(env, capsys):
    repo, *_ = env
    assert main([], repo_root=repo) == 0
    shown = capsys.readouterr().out
    assert "작업이 없습니다" in shown and "받은문서" not in shown


def test_init_makes_jobs_folder_and_master_only(tmp_path, capsys):
    repo = tmp_path / "repo"
    shutil.copytree(REPO / "rules", repo / "rules")
    work = tmp_path / "사무자동화"
    assert main(["init", str(work)], repo_root=repo) == 0
    assert sorted(p.name for p in work.iterdir()) == ["고객마스터.xlsx", "마크업표.xlsx", "작업"]
    assert (repo / "작업폴더.txt").read_text(encoding="utf-8").strip() == str(work.resolve())

    (work / "고객마스터.xlsx").write_bytes(b"mine")  # 사용자가 채운 고객마스터 · 마크업표는 그대로
    (work / "마크업표.xlsx").write_bytes(b"my markup")
    assert main(["init", str(work)], repo_root=repo) == 0
    assert (work / "고객마스터.xlsx").read_bytes() == b"mine"
    assert (work / "마크업표.xlsx").read_bytes() == b"my markup"


def test_old_inbox_flow_is_gone():
    for name in ("INBOX", "DUPLICATES", "ORGANIZED", "REVIEW", "LEDGER_FILE", "RULES_FILE", "OUTPUT"):
        assert not hasattr(config, name), name
    assert not (REPO / "수집.bat").exists()
    assert "uv run python -m app %*" in (REPO / "run.bat").read_text(encoding="utf-8")
    for old in ("_plan", "_move", "_write_output"):
        assert not hasattr(cli, old), old
    with pytest.raises(SystemExit):
        cli._parser().parse_args(["--덮어쓰기"])


# --- check --작업 ---

def test_check_with_job_shows_verify_table(env, capsys):
    repo, work, inp, out = env
    _new(repo, inp, out)
    before = _snapshot(work)
    capsys.readouterr()
    code = main(["check", str(INV), "--작업", "1"], repo_root=repo)
    shown = capsys.readouterr().out
    assert code == 0
    assert "| 인보이스번호 | INV-2026-10231 |" in shown
    assert "| 검산 | 원본 | 결과 | 일치 |" in shown
    assert "| 소계 + 부대비용 = 합계 | 111590.00 + 3850.00 = 115440.00 | 115440.00 | 일치 |" in shown
    assert "불일치" not in shown
    assert _snapshot(work) == before and _snapshot(out) == []


def test_check_with_job_shows_mismatch(env, capsys, monkeypatch):
    repo, work, inp, out = env
    _new(repo, inp, out)
    from app.verify import Check
    monkeypatch.setattr(cli, "verify", lambda result, path: [Check("합계금액", "100.00", "10.00", False)])
    capsys.readouterr()
    main(["check", str(INV), "--작업", "1"], repo_root=repo)
    shown = capsys.readouterr().out
    assert "| 합계금액 | 100.00 | 10.00 | 불일치 |" in shown
    assert "확인필요" in shown


def test_check_without_job_uses_the_only_job(env, capsys):
    repo, work, inp, out = env
    _new(repo, inp, out)
    capsys.readouterr()
    assert main(["check", str(INV)], repo_root=repo) == 0
    assert "| 검산 | 원본 | 결과 | 일치 |" in capsys.readouterr().out


def test_check_without_job_and_two_jobs_asks_which(env, capsys):
    repo, work, inp, out = env
    _new(repo, inp, out)
    _new(repo, inp, out, title="두번째")
    capsys.readouterr()
    assert main(["check", str(INV)], repo_root=repo) != 0
    assert "--작업" in capsys.readouterr().out


def test_check_excel_po_shows_master_and_company_form(env, capsys):
    repo, work, inp, out = env
    _new(repo, inp, out, mode="둘다")
    capsys.readouterr()
    assert main(["check", str(ACME_PO), "--작업", "1"], repo_root=repo) == 0
    shown = capsys.readouterr().out
    assert "엑셀양식: ACME" in shown
    assert "고객마스터: 찾음 (고객코드 C-KR-0417)" in shown
    assert "회사 양식: 회사양식_SO입력_PO-ACME-2026-0417.xlsx (4줄)" in shown


# --- run · approve ---

def test_run_auto_writes_record_and_notifies(env, monkeypatch, capsys):
    repo, work, inp, out = env
    job = _new(repo, inp, out)
    _drop(INV, inp)
    calls = []
    monkeypatch.setattr(cli, "notify_report", lambda job, report: calls.append((job.name, report.counts["승인대기"])))

    assert main(["run", "1", "--자동"], repo_root=repo) == 0
    record = (job.folder / "실행기록.txt").read_text(encoding="utf-8")
    assert "자동 실행" in record and f"{TITLE}: 성공 0" in record and "승인대기 1" in record
    assert calls == [(f"01_{TITLE}", 1)]
    assert load_records(job)[-1]["상태"] == "승인대기"  # 처음 10건 승인 문턱이 걸려 있음

    assert main(["run", "1"], repo_root=repo) == 0  # 새 것이 없으면 알리지 않는다
    assert len(calls) == 1
    record = (job.folder / "실행기록.txt").read_text(encoding="utf-8")
    assert record.count("자동 실행") == 1 and record.count("지금 실행") == 1


def test_run_paused_job_auto_does_nothing(env, monkeypatch, capsys):
    repo, work, inp, out = env
    job = _new(repo, inp, out)
    _drop(INV, inp)
    main(["job", "pause", "1"], repo_root=repo)
    monkeypatch.setattr(cli, "notify_report", lambda job, report: pytest.fail("알리면 안 됨"))
    capsys.readouterr()
    assert main(["run", "1", "--자동"], repo_root=repo) == 0
    assert "일시중지" in capsys.readouterr().out
    assert load_records(job) == []


def test_approve_command_outputs_and_notifies(env, monkeypatch, capsys):
    repo, work, inp, out = env
    job = _new(repo, inp, out)
    _drop(INV, inp)
    main(["run", "1"], repo_root=repo)
    calls = []
    monkeypatch.setattr(cli, "notify_report", lambda job, report: calls.append(report.counts["성공"]))
    capsys.readouterr()

    assert main(["approve", "1"], repo_root=repo) == 0
    shown = capsys.readouterr().out
    assert "Invoice 1/10" in shown
    assert calls == [1]
    assert "승인" in (job.folder / "실행기록.txt").read_text(encoding="utf-8")


@pytest.mark.parametrize("command", [["run", "1"], ["run", "1", "--자동"], ["approve", "1"]])
def test_locked_job_shows_korean_message_without_traceback(env, capsys, command):
    repo, work, inp, out = env
    job = _new(repo, inp, out)
    capsys.readouterr()
    with job_lock(job):
        code = main(command, repo_root=repo)
    captured = capsys.readouterr()
    assert code != 0
    assert "이미 실행 중" in captured.out
    assert "Traceback" not in captured.out + captured.err


def test_run_unknown_job_says_so(env, capsys):
    repo, *_ = env
    assert main(["run", "7"], repo_root=repo) != 0
    assert "찾지 못했습니다" in capsys.readouterr().out


def test_run_with_skeleton_rules_reports_rule_error(env, monkeypatch, capsys):
    repo, work, inp, out = env
    job = _new(repo, inp, out, rules=None)  # 아직 Claude 가 규칙을 채우지 않은 작업
    calls = []
    monkeypatch.setattr(cli, "notify_report", lambda job, report: calls.append(report))
    capsys.readouterr()
    assert main(["run", "1", "--자동"], repo_root=repo) != 0
    assert "규칙.yaml 오류" in capsys.readouterr().out
    assert len(calls) == 1  # 자동 실행이 조용히 건너뛰지 않게 알린다
    assert "규칙.yaml 오류" in (job.folder / "실행기록.txt").read_text(encoding="utf-8")


# --- pick-folder ---

def test_pick_folder_rejects_sync_folder(tmp_path, monkeypatch, capsys):
    synced = tmp_path / "OneDrive" / "주문"
    monkeypatch.setattr(init_module, "pick_folder", lambda title: str(synced))
    code = main(["pick-folder", "입력 폴더를 고르세요"], repo_root=tmp_path)
    shown = capsys.readouterr().out
    assert code != 0
    assert "동기화 폴더" in shown
    assert shown.strip().splitlines()[-1] != str(synced)  # 고른 경로로 쓰이지 않게 마지막 줄은 경로가 아님


def test_pick_folder_prints_chosen_path_on_last_line(tmp_path, monkeypatch, capsys):
    titles = []
    monkeypatch.setattr(init_module, "pick_folder", lambda title: titles.append(title) or str(tmp_path))
    assert main(["pick-folder", "출력 폴더를 고르세요"], repo_root=tmp_path) == 0
    assert capsys.readouterr().out.strip().splitlines()[-1] == str(tmp_path)
    assert titles == ["출력 폴더를 고르세요"]


def test_pick_folder_cancel(tmp_path, monkeypatch, capsys):
    monkeypatch.setattr(init_module, "pick_folder", lambda title: "")
    assert main(["pick-folder", "입력 폴더를 고르세요"], repo_root=tmp_path) != 0
    assert capsys.readouterr().out.strip().splitlines()[-1] == "취소함"


# --- 뼈대 규칙 ---

def test_skeleton_rules_is_commented_guide_not_empty_doc_types():
    text = SKELETON.read_text(encoding="utf-8")
    data = yaml.safe_load(text)
    assert isinstance(data, dict) and data.get("문서종류")  # 비어 있는 문서종류: {} 가 아님
    assert sum(1 for line in text.splitlines() if line.lstrip().startswith("#")) >= 10
    with pytest.raises(RuleError):
        load_rules(SKELETON)
