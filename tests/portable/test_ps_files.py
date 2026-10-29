"""무설치판 파일 자체 검사: BOM · 인터넷 코드 없음 · Windows PowerShell 5.1 문법."""
import json
import re

from .conftest import PORTABLE, decode_output, run_powershell

BOM = b"\xef\xbb\xbf"
NETWORK = re.compile(r"Invoke-WebRequest|Invoke-RestMethod|WebClient|https?://", re.IGNORECASE)
PARSE_ALL = """\
param([string]$List, [string]$Out)
$files = [IO.File]::ReadAllText($List, [Text.Encoding]::UTF8) | ConvertFrom-Json
$errors = @()
foreach ($file in $files) {
    $tokens = $null
    $found = $null
    [void][System.Management.Automation.Language.Parser]::ParseFile($file, [ref]$tokens, [ref]$found)
    foreach ($e in $found) { $errors += ($file + ': ' + $e.Message) }
}
[IO.File]::WriteAllText($Out, (ConvertTo-Json -InputObject @($errors)), (New-Object Text.UTF8Encoding $false))
"""


def _files(*suffixes):
    return sorted(p for p in PORTABLE.rglob("*") if p.suffix.lower() in suffixes)


def test_bom_and_no_network():
    files = _files(".ps1", ".bat")
    assert files, "무설치판 파일이 없음"
    for path in files:
        data = path.read_bytes()
        assert data.startswith(BOM), f"BOM 없음: {path.name}"
        found = NETWORK.search(data.decode("utf-8-sig"))
        assert found is None, f"인터넷 코드 '{found.group(0)}' 가 있음: {path.name}"


def test_no_ps7_syntax(powershell_exe, tmp_path):
    """모든 .ps1 을 powershell.exe(5.1) 파서로 읽어 오류 0개."""
    files = _files(".ps1")
    assert files, "무설치판 .ps1 이 없음"
    script = tmp_path / "parse_all.ps1"
    script.write_bytes(BOM + PARSE_ALL.encode("utf-8"))
    listing = tmp_path / "files.json"
    listing.write_text(json.dumps([str(p) for p in files], ensure_ascii=False), encoding="utf-8")
    out = tmp_path / "errors.json"

    done = run_powershell(powershell_exe, script, "-List", str(listing), "-Out", str(out))

    assert done.returncode == 0, decode_output(done.stderr)
    assert json.loads(out.read_text(encoding="utf-8")) == []
