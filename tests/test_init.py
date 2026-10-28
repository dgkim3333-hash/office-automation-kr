"""설치.bat 본체(app init)와 작업 폴더 설정."""
import shutil
from pathlib import Path

import pytest

from app import config
from app.init import InitError, init

REPO = Path(__file__).resolve().parent.parent


@pytest.fixture
def repo(tmp_path):
    """가짜 저장소 루트(규칙 예시만 있음)."""
    root = tmp_path / "repo"
    (root / "rules").mkdir(parents=True)
    shutil.copy(REPO / "rules" / "규칙.예시.yaml", root / "rules" / "규칙.예시.yaml")
    return root


def test_creates_jobs_folder_and_master_only(tmp_path, repo):
    work = tmp_path / "사무자동화"
    init(work_dir=work, repo_root=repo)
    assert sorted(p.name for p in work.iterdir()) == ["고객마스터.xlsx", "작업"]
    assert (work / "작업").is_dir()


def test_keeps_existing_master_and_jobs(tmp_path, repo):
    work = tmp_path / "사무자동화"
    (work / "작업" / "01_일일 정리").mkdir(parents=True)
    (work / "고객마스터.xlsx").write_bytes(b"mine")
    init(work_dir=work, repo_root=repo)
    assert (work / "고객마스터.xlsx").read_bytes() == b"mine"
    assert (work / "작업" / "01_일일 정리").is_dir()


@pytest.mark.parametrize("sub", ["", "rules", "a/b"])
def test_rejects_repo_folder_or_below(repo, sub):
    with pytest.raises(InitError, match="저장소"):
        init(work_dir=repo / sub, repo_root=repo)
    assert not (repo / "작업폴더.txt").exists()


def test_saves_work_folder_txt(tmp_path, repo, capsys):
    work = tmp_path / "사무자동화"
    init(work_dir=work, repo_root=repo)
    assert (repo / "작업폴더.txt").read_text(encoding="utf-8").strip() == str(work.resolve())
    assert "04장" in capsys.readouterr().out


def test_config_work_dir_reads_txt(tmp_path, repo):
    work = tmp_path / "사무자동화"
    init(work_dir=work, repo_root=repo)
    assert config.work_dir(repo) == work.resolve()


def test_config_work_dir_missing_file(repo):
    with pytest.raises(config.ConfigError, match="설치.bat를 먼저 실행하세요"):
        config.work_dir(repo)


def test_config_work_dir_vanished_folder(tmp_path, repo):
    (repo / "작업폴더.txt").write_text(str(tmp_path / "없는폴더"), encoding="utf-8")
    with pytest.raises(config.ConfigError, match="설치.bat를 먼저 실행하세요"):
        config.work_dir(repo)


def test_rerun_keeps_current_when_user_says_no(tmp_path, repo, capsys):
    work = tmp_path / "사무자동화"
    init(work_dir=work, repo_root=repo)
    answers = iter(["n"])
    result = init(repo_root=repo, ask=lambda prompt: next(answers))
    assert result == work.resolve()
    assert str(work.resolve()) in capsys.readouterr().out


def test_rerun_new_path_typed_when_picker_cancelled(tmp_path, repo):
    init(work_dir=tmp_path / "old", repo_root=repo)
    answers = iter(["y", str(tmp_path / "new")])
    result = init(repo_root=repo, ask=lambda prompt: next(answers), pick=lambda: "")
    assert result == (tmp_path / "new").resolve()
    assert (tmp_path / "new" / "작업").is_dir()
    assert (tmp_path / "old" / "작업").is_dir()  # 기존 데이터는 그대로


def test_main_init_with_path_does_not_ask(repo, tmp_path, monkeypatch, capsys):
    """설치 스크립트용: python -m app init <경로> 는 묻지 않고 바로 만든다."""
    from app.__main__ import main

    import app.init as init_module

    def must_not_ask(*args):
        raise AssertionError("물으면 안 됨")

    monkeypatch.setattr("builtins.input", must_not_ask)
    work_dir, repo_root, _, _ = init_module.init.__defaults__
    monkeypatch.setattr(init_module.init, "__defaults__", (work_dir, repo_root, must_not_ask, must_not_ask))
    work = tmp_path / "사무자동화"
    assert main(["init", str(work)], repo_root=repo) == 0
    assert (work / "작업").is_dir() and (work / "고객마스터.xlsx").is_file()
    assert (repo / "작업폴더.txt").read_text(encoding="utf-8").strip() == str(work.resolve())
    assert main(["init", str(work)], repo_root=repo) == 0  # 다시 실행해도 그대로
    assert "04장" in capsys.readouterr().out