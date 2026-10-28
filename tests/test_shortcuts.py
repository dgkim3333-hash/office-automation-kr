"""바탕화면 · 시작프로그램 바로가기: PowerShell(WScript.Shell)로 만들고, 한글 경로는 -EncodedCommand 로 넘긴다.
실제 바탕화면은 건드리지 않는다(가짜 실행기 · 임시 폴더)."""
from pathlib import Path

import pytest

from app import powershell, shortcuts
from app.jobs import create_job


@pytest.fixture
def job(tmp_path):
    inp, out = tmp_path / "in", tmp_path / "out"
    inp.mkdir()
    out.mkdir()
    return create_job(tmp_path / "work", "일일 PO 정리", inp, out, "누적장부")


def test_command_is_encoded_without_profile():
    args = powershell.command("Write-Output '바탕 화면'")
    assert args[0].lower().endswith("powershell.exe")
    assert "-NoProfile" in args and "-EncodedCommand" in args
    assert "바탕" not in " ".join(args)  # 한글은 base64 안에만 있다
    script = powershell.decode_command(args)
    assert "Write-Output '바탕 화면'" in script
    assert "$ErrorActionPreference = 'Stop'" in script and "$ProgressPreference = 'SilentlyContinue'" in script
    assert "exit 1" in script  # 오류는 반환 코드 1 로


def test_quote_doubles_single_quotes():
    assert powershell.quote("C:\\O'Neil 폴더") == "'C:\\O''Neil 폴더'"


def test_quote_doubles_typographic_single_quotes():
    """PowerShell 은 ‘ ’ ‚ ‛ 도 작은따옴표로 본다."""
    assert powershell.quote("O\u2019Neil \u2018a\u2019 \u201ab \u201bc") == (
        "'O\u2019\u2019Neil \u2018\u2018a\u2019\u2019 \u201a\u201ab \u201b\u201bc'")


def test_title_with_typographic_quote_is_quoted_safely(fake_ps, tmp_path):
    inp, out = tmp_path / "in", tmp_path / "out"
    inp.mkdir()
    out.mkdir()
    odd = create_job(tmp_path / "work", "O\u2019Neil 주문", inp, out, "누적장부")
    shortcuts.job_shortcuts(odd, tmp_path / "repo", tmp_path / "desk")
    assert "O\u2019\u2019Neil 주문.lnk'" in fake_ps.having("'run 01'")[0]


def test_desktop_dir_follows_what_windows_says(fake_ps, tmp_path):
    moved = tmp_path / "OneDrive" / "바탕 화면"
    fake_ps.outputs["'Desktop'"] = f"{moved}\r\n"
    assert shortcuts.desktop_dir() == moved
    assert "[Environment]::GetFolderPath('Desktop')" in fake_ps.scripts[-1]


def test_startup_dir(fake_ps, tmp_path):
    fake_ps.outputs["'Startup'"] = str(tmp_path / "시작프로그램")
    assert shortcuts.startup_dir() == tmp_path / "시작프로그램"


def test_desktop_dir_failure_raises_shortcut_error(fake_ps):
    fake_ps.fail.append("GetFolderPath")
    with pytest.raises(shortcuts.ShortcutError):
        shortcuts.desktop_dir()


def test_create_uses_wscript_shell(fake_ps, tmp_path):
    link = tmp_path / "지금 실행 - 일일 PO 정리.lnk"
    shortcuts.create(link, Path("C:/repo/run.bat"), "run 01", Path("C:/repo"))
    script = fake_ps.scripts[-1]
    assert "WScript.Shell" in script and "CreateShortcut" in script
    assert f"'{link}'" in script and "'run 01'" in script


def test_create_failure_raises(fake_ps, tmp_path):
    fake_ps.fail.append("CreateShortcut")
    with pytest.raises(shortcuts.ShortcutError):
        shortcuts.create(tmp_path / "a.lnk", Path("C:/repo/run.bat"), "run 01", Path("C:/repo"))


def test_job_shortcuts_names_and_args(fake_ps, tmp_path, job):
    repo, desktop = tmp_path / "repo", tmp_path / "desk"
    made = shortcuts.job_shortcuts(job, repo, desktop)
    assert [p.name for p in made] == ["지금 실행 - 일일 PO 정리.lnk", "승인 - 일일 PO 정리.lnk"]
    assert all(p.parent == desktop for p in made)
    run_now, approve = fake_ps.having("CreateShortcut")
    assert "'run 01'" in run_now and f"'{repo / 'run.bat'}'" in run_now and f"'{repo}'" in run_now
    assert "'approve 01'" in approve


def test_remove_job_shortcuts_only_this_job(tmp_path, job):
    desktop = tmp_path / "desk"
    desktop.mkdir()
    mine = [desktop / "지금 실행 - 일일 PO 정리.lnk", desktop / "승인 - 일일 PO 정리.lnk"]
    other = desktop / "지금 실행 - 다른 작업.lnk"
    for p in mine + [other]:
        p.write_bytes(b"lnk")
    assert sorted(shortcuts.remove_job_shortcuts(job, desktop)) == sorted(mine)
    assert not any(p.exists() for p in mine) and other.exists()
