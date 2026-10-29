"""무설치판(PowerShell 5.1) 시험 공통: 엔진 부르기 도우미와 예시 규칙.

tests\\conftest.py 가 app.powershell.default_run 을 가짜로 바꾸므로, 무설치판은 subprocess 로 powershell.exe 를 직접 부른다.
powershell.exe 가 없는 PC(클라우드 등)에서는 무설치판 시험을 건너뛴다.
"""
import json
import shutil
import subprocess
from pathlib import Path

import pytest

from app.portable import rules_to_portable
from app.rules import rules_from_text

REPO = Path(__file__).resolve().parent.parent.parent
PORTABLE = REPO / "무설치판"
DISPATCHER = PORTABLE / "무설치판.ps1"
EXAMPLE_RULES = REPO / "rules" / "규칙.예시.yaml"


class EngineError(RuntimeError):
    """엔진이 오류로 끝남(반환 코드 0 이 아님). 메시지는 엔진이 남긴 오류 글자."""


def without_master(text: str) -> str:
    """예시 규칙에서 고객마스터.* 열 줄을 뺀다(무설치판은 고객마스터를 지원하지 않음)."""
    return "\n".join(line for line in text.splitlines() if "값: 고객마스터." not in line) + "\n"


def example_rules():
    """예시 규칙(고객마스터 열 제외)을 지금 엔진으로 해석한 것."""
    return rules_from_text(without_master(EXAMPLE_RULES.read_text(encoding="utf-8")))


@pytest.fixture
def powershell_exe() -> str:
    exe = shutil.which("powershell.exe")
    if exe is None:
        pytest.skip("powershell.exe 가 없어 무설치판 시험을 건너뜀")
    return exe


def decode_output(data: bytes) -> str:
    """powershell.exe 출력 바이트를 글자로(콘솔 코드 페이지에 따라 utf-8 또는 cp949)."""
    try:
        return data.decode("utf-8")
    except UnicodeDecodeError:
        return data.decode("mbcs", errors="replace")


def run_powershell(exe: str, script: Path, *args: str) -> subprocess.CompletedProcess:
    """powershell.exe(5.1)로 스크립트 파일을 돌린다. 출력은 바이트 그대로 돌려준다."""
    return subprocess.run(
        [exe, "-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass", "-File", str(script), *args],
        capture_output=True, timeout=120)


@pytest.fixture
def run_engine(powershell_exe, tmp_path):
    """run_engine(op, payload, culture=None) -> dict. 입력 JSON 을 tmp_path 에 쓰고 무설치판.ps1 -Command 엔진 을 돌린다."""
    calls = iter(range(1, 10_000))

    def run(op: str, payload: dict, culture: str | None = None) -> dict:
        n = next(calls)
        source = tmp_path / f"엔진입력_{n}.json"
        target = tmp_path / f"엔진출력_{n}.json"
        source.write_text(json.dumps({"op": op, **payload}, ensure_ascii=False), encoding="utf-8")
        args = ["-Command", "엔진", "-Input", str(source), "-Output", str(target)]
        if culture is not None:
            args += ["-Culture", culture]
        done = run_powershell(powershell_exe, DISPATCHER, *args)
        data = json.loads(target.read_text(encoding="utf-8-sig")) if target.exists() else None
        if done.returncode != 0:
            message = data.get("error") if isinstance(data, dict) else None
            raise EngineError(message or decode_output(done.stderr))
        return data

    return run
