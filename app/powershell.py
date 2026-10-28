"""PowerShell 한 번 실행하기(작업 스케줄러 등록 · 바로가기 · Windows 알림이 쓴다).

- 한글 경로가 깨지지 않게 스크립트를 UTF-16LE base64(-EncodedCommand)로 넘기고, 출력은 UTF-8 로 받는다.
- 사용자 프로필은 읽지 않는다(-NoProfile). 창은 띄우지 않는다(CREATE_NO_WINDOW).
- 스크립트 안의 오류는 모두 멈춤 오류로 바꾸고, 잡아서 오류 글자를 stderr 에 쓴 뒤 반환 코드 1 로 끝낸다.
"""
import base64
import re
import subprocess
from dataclasses import dataclass

EXE = "powershell.exe"  # Windows PowerShell 5.1: WinRT 알림 · ScheduledTasks 모듈이 기본으로 있다
TIMEOUT = 120
# 진행 표시(「모듈을 준비하는 중」)가 오류 출력에 섞이지 않게 끈다
_SINGLE_QUOTES = re.compile("['‘’‚‛]")
PRELUDE = ("$ErrorActionPreference = 'Stop'\n$ProgressPreference = 'SilentlyContinue'\n"
           "[Console]::OutputEncoding = [System.Text.Encoding]::UTF8\n")


@dataclass
class Result:
    ok: bool
    out: str = ""
    err: str = ""


def default_run(args, **kwargs):
    """실제 실행. 테스트는 이것을 가짜로 바꿔(tests/conftest.py) 실제 PC 를 건드리지 않는다."""
    return subprocess.run(args, **kwargs)


def quote(text) -> str:
    """PowerShell 작은따옴표 문자열. PowerShell 은 ' 뿐 아니라 ‘ ’ ‚ ‛ 도 작은따옴표로 보므로 모두 두 번 쓴다."""
    return "'" + _SINGLE_QUOTES.sub(lambda m: m.group(0) * 2, str(text)) + "'"


def command(script: str) -> list[str]:
    body = PRELUDE + "try {\n" + script + "\n} catch {\n[Console]::Error.WriteLine($_.Exception.Message)\nexit 1\n}\n"
    encoded = base64.b64encode(body.encode("utf-16-le")).decode("ascii")
    return [EXE, "-NoProfile", "-NonInteractive", "-EncodedCommand", encoded]


def decode_command(args) -> str:
    """command() 가 만든 인자에서 스크립트 글자를 되살린다(테스트 · 문제 확인용)."""
    args = list(args)
    return base64.b64decode(args[args.index("-EncodedCommand") + 1]).decode("utf-16-le")


def _text(data) -> str:
    if data is None:
        return ""
    if isinstance(data, bytes):
        return data.decode("utf-8", errors="replace")
    return str(data)


def run_script(script: str, run=None, timeout: int = TIMEOUT) -> Result:
    """스크립트를 실행해 Result 를 돌려준다. PowerShell 을 찾지 못하거나 시간이 넘어도 예외 대신 실패."""
    run = run or default_run
    try:
        done = run(command(script), capture_output=True, timeout=timeout,
                   creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
    except (OSError, subprocess.SubprocessError) as e:
        return Result(False, "", str(e))
    return Result(done.returncode == 0, _text(done.stdout).strip(), _text(done.stderr).strip())
