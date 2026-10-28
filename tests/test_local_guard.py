"""이 PC 전용 가드 — 클라우드·원격에서는 돌지 않고, 실행 중 인터넷 연결을 막는다."""
import shutil
import socket
from pathlib import Path

import pytest

from app import local_guard
from app.__main__ import main
from app.init import InitError, init
from app.local_guard import LocalOnlyError, block_network, check_environment, check_work_dir

REPO = Path(__file__).resolve().parent.parent
SAMPLE = REPO / "samples" / "03_Invoice_DM_PDF" / "INV-2026-10231.pdf"


@pytest.fixture
def repo(tmp_path):
    root = tmp_path / "repo"
    (root / "rules").mkdir(parents=True)
    shutil.copy(REPO / "rules" / "규칙.예시.yaml", root / "rules" / "규칙.예시.yaml")
    return root


def test_windows_pc_passes():
    check_environment()


def test_non_windows_is_refused(monkeypatch):
    monkeypatch.setattr(local_guard.sys, "platform", "linux")
    with pytest.raises(LocalOnlyError, match="Windows"):
        check_environment()


@pytest.mark.parametrize("var", ["CLAUDE_CODE_REMOTE", "CODESPACES", "GITHUB_ACTIONS"])
def test_cloud_markers_are_refused(monkeypatch, var):
    monkeypatch.setenv(var, "true")
    with pytest.raises(LocalOnlyError, match="클라우드"):
        check_environment()


def test_local_folder_passes(tmp_path):
    check_work_dir(tmp_path)


@pytest.mark.parametrize("parts", [["OneDrive - ACME", "사무자동화"], ["OneDrive", "문서"], ["Dropbox", "a"]])
def test_sync_folders_are_refused(tmp_path, parts):
    with pytest.raises(LocalOnlyError, match="동기화"):
        check_work_dir(tmp_path.joinpath(*parts))


def test_unc_path_is_refused():
    with pytest.raises(LocalOnlyError, match="네트워크"):
        check_work_dir(Path(r"\\fileserver\share\사무자동화"))


def test_network_drive_letter_is_refused(monkeypatch, tmp_path):
    monkeypatch.setattr(local_guard, "_drive_type", lambda root: local_guard.DRIVE_REMOTE)
    with pytest.raises(LocalOnlyError, match="네트워크"):
        check_work_dir(tmp_path)


def test_block_network_refuses_connections_and_restores():
    restore = block_network()
    try:
        with pytest.raises(LocalOnlyError):
            socket.create_connection(("example.com", 80), timeout=1)
        with pytest.raises(LocalOnlyError):
            socket.socket().connect(("127.0.0.1", 9))
        with pytest.raises(LocalOnlyError):
            socket.getaddrinfo("example.com", 80)
    finally:
        restore()
    assert socket.socket.connect is not None and socket.create_connection.__module__ == "socket"


def test_main_prints_local_only_banner(repo, tmp_path, capsys):
    init(work_dir=tmp_path / "work", repo_root=repo)
    assert main([], repo_root=repo) == 0  # 작업 목록 안내(처리 없음)
    out = capsys.readouterr().out
    assert "이 PC" in out and "인터넷" in out and "Code 탭" in out


def test_main_refuses_in_cloud(repo, tmp_path, monkeypatch, capsys):
    init(work_dir=tmp_path / "work", repo_root=repo)
    monkeypatch.setenv("CLAUDE_CODE_REMOTE", "true")
    assert main(["check", str(SAMPLE)], repo_root=repo) == 1
    assert "클라우드" in capsys.readouterr().out


def test_main_refuses_sync_work_folder(repo, tmp_path, capsys):
    work = tmp_path / "OneDrive" / "사무자동화"
    work.mkdir(parents=True)
    (repo / "작업폴더.txt").write_text(str(work), encoding="utf-8")
    assert main(["check", str(SAMPLE)], repo_root=repo) == 1
    assert "동기화" in capsys.readouterr().out


def test_init_refuses_sync_folder(repo, tmp_path):
    with pytest.raises(InitError, match="동기화"):
        init(work_dir=tmp_path / "OneDrive" / "사무자동화", repo_root=repo)
    assert not (repo / "작업폴더.txt").exists()
