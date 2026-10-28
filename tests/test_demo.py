"""04장 첫 실행: 체험.bat (python -m app demo) — 가짜 샘플을 따로 된 체험 폴더에서 처리한다."""
import shutil
from pathlib import Path

import openpyxl
import pytest

import app.__main__ as cli
from app.__main__ import main
from app.init import init

REPO = Path(__file__).resolve().parent.parent


@pytest.fixture
def env(tmp_path, monkeypatch):
    repo = tmp_path / "repo"
    (repo / "rules").mkdir(parents=True)
    shutil.copy(REPO / "rules" / "규칙.예시.yaml", repo / "rules" / "규칙.예시.yaml")
    shutil.copytree(REPO / "samples" / "03_Invoice_DM_PDF", repo / "samples" / "03_Invoice_DM_PDF")
    work = init(work_dir=tmp_path / "사무자동화", repo_root=repo)
    opened = []
    monkeypatch.setattr(cli, "_open", lambda path: opened.append(Path(path)))  # 엑셀·탐색기를 실제로 열지 않음
    return repo, work, opened


def _rows(path, sheet):
    ws = openpyxl.load_workbook(path)[sheet]
    return [r for r in ws.iter_rows(min_row=2, values_only=True) if any(v is not None for v in r)]


def test_demo_processes_samples_in_separate_folder(env, capsys):
    repo, work, opened = env
    assert main(["demo"], repo_root=repo) == 0
    out = capsys.readouterr().out
    demo = work / "_체험"
    assert "성공 4 · 중복 0 · 확인필요 2" in out
    assert len(_rows(demo / "누적장부.xlsx", "Invoice")) == 2
    assert len(_rows(demo / "누적장부.xlsx", "DM")) == 2
    assert len(list((demo / "확인필요").glob("*.pdf"))) == 2
    assert opened == [demo / "누적장부.xlsx"]


def test_demo_leaves_real_work_folder_alone(env):
    repo, work, _ = env
    (work / "받은문서" / "진짜.pdf").write_bytes(b"%PDF real")
    main(["demo"], repo_root=repo)
    assert (work / "받은문서" / "진짜.pdf").read_bytes() == b"%PDF real"
    assert not (work / "누적장부.xlsx").exists()
    assert not list((work / "정리문서").iterdir())


def test_demo_again_starts_fresh(env, capsys):
    repo, work, _ = env
    main(["demo"], repo_root=repo)
    main(["demo"], repo_root=repo)
    out = capsys.readouterr().out
    assert out.count("성공 4 · 중복 0 · 확인필요 2") == 2  # 두 번째도 중복 없이 같은 결과
    assert len(_rows(work / "_체험" / "누적장부.xlsx", "Invoice")) == 2


def test_demo_uses_example_rules_not_my_rules(env, capsys):
    repo, work, _ = env
    (work / "규칙.yaml").write_text("문서종류: 망가진 내 규칙", encoding="utf-8")
    assert main(["demo"], repo_root=repo) == 0
    assert "성공 4" in capsys.readouterr().out


def test_demo_explains_result_in_plain_words(env, capsys):
    repo, _, _ = env
    main(["demo"], repo_root=repo)
    out = capsys.readouterr().out
    assert "체험 폴더" in out and "진짜" in out


def test_demo_bat_exists():
    bat = REPO / "체험.bat"
    text = bat.read_bytes().decode("utf-8")
    assert "python -m app demo" in text and "pause" in text
    assert b"\r\n" in bat.read_bytes() or "\n" in text  # .gitattributes 가 CRLF 로 내보냄
