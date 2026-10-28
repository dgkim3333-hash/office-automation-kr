"""install.ps1 (한 줄 자동 설치) 검사. 실제 프로그램 설치는 -DryRun 으로 건너뛴다."""
import hashlib
import subprocess
import zipfile
from pathlib import Path

import pytest

REPO = Path(__file__).resolve().parent.parent
SCRIPT = REPO / "install.ps1"
PS = ["powershell.exe", "-NoProfile", "-ExecutionPolicy", "Bypass"]


def _run_like_irm(*args):
    """irm | iex 처럼: 파일을 UTF-8 글자로 읽어 스크립트 블록으로 실행(인자는 시험용)."""
    quoted = " ".join(a if a.startswith("-") else "'" + a.replace("'", "''") + "'" for a in args)
    command = (f"& ([scriptblock]::Create([IO.File]::ReadAllText('{SCRIPT}', [Text.Encoding]::UTF8))) {quoted}")
    return subprocess.run(PS + ["-Command", command], capture_output=True, timeout=600)


def test_script_has_no_syntax_errors():
    check = (
        "$e = $null; [System.Management.Automation.Language.Parser]::ParseInput("
        f"[IO.File]::ReadAllText('{SCRIPT}', [Text.Encoding]::UTF8), [ref]$null, [ref]$e) | Out-Null; $e.Count"
    )
    out = subprocess.run(PS + ["-Command", check], capture_output=True, text=True, timeout=60)
    assert out.stdout.strip() == "0", out.stdout + out.stderr


def test_script_has_no_bom_so_irm_iex_works():
    # irm | iex 로 받으면 BOM 이 맨 앞 글자로 남아 param(...) 블록을 못 읽는다(Windows PowerShell 5.1 실측)
    assert not SCRIPT.read_bytes().startswith(b"\xef\xbb\xbf")
    SCRIPT.read_text(encoding="utf-8")  # UTF-8 로 읽혀야 함(GitHub raw 는 charset=utf-8)


@pytest.fixture
def repo_zip(tmp_path):
    """GitHub 의 main.zip 과 같은 모양(office-automation-kr-main/…)의 압축 파일.

    커밋된 판이 아니라 지금 작업 중인 파일로 만든다(커밋 전 변경도 시험되게).
    """
    archive = tmp_path / "main.zip"
    listed = subprocess.run(["git", "ls-files", "-co", "--exclude-standard", "-z"], cwd=REPO,
                            capture_output=True, check=True).stdout.decode("utf-8").split("\0")
    with zipfile.ZipFile(archive, "w", zipfile.ZIP_DEFLATED) as z:
        for name in filter(None, listed):
            if (REPO / name).is_file():
                z.write(REPO / name, f"office-automation-kr-main/{name}")
    return archive


def test_dry_run_downloads_repo_and_makes_work_folder(tmp_path, repo_zip):
    install_dir = tmp_path / "dev"
    work_dir = tmp_path / "사무자동화"
    proc = _run_like_irm("-DryRun", "-RepoZip", str(repo_zip), "-InstallDir", str(install_dir),
                         "-WorkDir", str(work_dir))
    out = proc.stdout.decode("utf-8", "replace") + proc.stderr.decode("utf-8", "replace")
    assert proc.returncode == 0, out
    repo = install_dir / "office-automation-kr-main"
    assert (repo / "run.bat").is_file()
    assert (work_dir / "받은문서" / "중복").is_dir()
    assert (repo / "작업폴더.txt").read_text(encoding="utf-8").strip() == str(work_dir)
    assert not zipfile.is_zipfile(install_dir / "office-automation-kr-main.zip")  # 임시 zip 은 지움


def test_dry_run_twice_keeps_existing_repo(tmp_path, repo_zip):
    install_dir = tmp_path / "dev"
    args = ("-DryRun", "-RepoZip", str(repo_zip), "-InstallDir", str(install_dir),
            "-WorkDir", str(tmp_path / "사무자동화"))
    assert _run_like_irm(*args).returncode == 0
    marker = install_dir / "office-automation-kr-main" / "내메모.txt"
    marker.write_text("지우면 안 됨", encoding="utf-8")
    assert _run_like_irm(*args).returncode == 0
    assert marker.read_text(encoding="utf-8") == "지우면 안 됨"


TEMPLATE = REPO / "claude" / "전역_CLAUDE.md"


def _real_global_claude_md_hash():
    real = Path.home() / ".claude" / "CLAUDE.md"
    return hashlib.sha256(real.read_bytes()).hexdigest() if real.exists() else None


def test_global_claude_md_is_copied_when_missing(tmp_path, repo_zip):
    home = tmp_path / "claude_home"  # 없는 폴더에서 시작
    proc = _run_like_irm("-DryRun", "-RepoZip", str(repo_zip), "-InstallDir", str(tmp_path / "dev"),
                         "-WorkDir", str(tmp_path / "사무자동화"), "-ClaudeHome", str(home))
    assert proc.returncode == 0, proc.stdout.decode("utf-8", "replace")
    assert (home / "CLAUDE.md").read_bytes() == TEMPLATE.read_bytes()


def test_existing_global_claude_md_is_never_touched(tmp_path, repo_zip):
    home = tmp_path / "claude_home"
    home.mkdir()
    mine = "# 내 전역 규칙\n- 건드리면 안 됨\n".encode("utf-8")
    (home / "CLAUDE.md").write_bytes(mine)
    proc = _run_like_irm("-DryRun", "-RepoZip", str(repo_zip), "-InstallDir", str(tmp_path / "dev"),
                         "-WorkDir", str(tmp_path / "사무자동화"), "-ClaudeHome", str(home))
    assert proc.returncode == 0, proc.stdout.decode("utf-8", "replace")
    assert (home / "CLAUDE.md").read_bytes() == mine
    assert sorted(p.name for p in home.iterdir()) == ["CLAUDE.md"]  # 다른 파일도 만들지 않음
    assert "건드리지 않았습니다" in proc.stdout.decode("utf-8", "replace")


def test_dry_run_without_claude_home_skips_global_step(tmp_path, repo_zip):
    before = _real_global_claude_md_hash()
    proc = _run_like_irm("-DryRun", "-RepoZip", str(repo_zip), "-InstallDir", str(tmp_path / "dev"),
                         "-WorkDir", str(tmp_path / "사무자동화"))
    out = proc.stdout.decode("utf-8", "replace")
    assert proc.returncode == 0, out
    assert "[6/6]" in out
    assert _real_global_claude_md_hash() == before  # 이 PC 의 진짜 전역 CLAUDE.md 는 그대로


def test_global_template_rules():
    text = TEMPLATE.read_text(encoding="utf-8")
    for must in ("한국어", "허락", "Cowork", "Code 탭", "office-automation-kr-main"):
        assert must in text, must
    assert not TEMPLATE.read_bytes().startswith(b"\xef\xbb\xbf")