"""모든 테스트 공통: 실제 PowerShell(작업 스케줄러 · 바탕화면 바로가기 · Windows 알림)과 실제 메시지 창을 막는다.

잠긴 작업을 기다리는 자동 실행의 기다리기(app.schedule._sleep)도 바로 돌아오게 한다.
app.powershell.default_run 을 「늘 실패하는 가짜」로 바꾸므로, 어느 테스트가 실수로 job new 등을 불러도
실제 작업 스케줄러에 등록되거나 실제 바탕화면에 바로가기가 생기지 않는다. PowerShell 명령을 흉내 내야 하는
테스트는 fake_ps 픽스처를 쓴다.
"""
import subprocess

import pytest

import app.notify
import app.powershell
import app.schedule


class FakePowerShell:
    """powershell.default_run 자리에 들어가는 가짜 실행기. 받은 스크립트 글자를 모아 두고,
    fail 의 글자가 들어 있는 스크립트는 실패(반환 코드 1), outputs 의 글자가 들어 있으면 그 출력을 돌려준다."""

    def __init__(self):
        self.scripts: list[str] = []
        self.fail: list[str] = []
        self.outputs: dict[str, str] = {}

    def __call__(self, args, **kwargs):
        script = app.powershell.decode_command(args)
        self.scripts.append(script)
        failed = any(text in script for text in self.fail)
        out = next((value for key, value in self.outputs.items() if key in script), "")
        return subprocess.CompletedProcess(args, 1 if failed else 0, stdout=out.encode("utf-8"),
                                           stderr="가짜 실패".encode("utf-8") if failed else b"")

    def having(self, text: str) -> list[str]:
        return [s for s in self.scripts if text in s]


def _refuse(args, **kwargs):
    return subprocess.CompletedProcess(args, 1, stdout=b"", stderr="테스트에서는 PowerShell 을 실행하지 않음".encode("utf-8"))


@pytest.fixture(autouse=True)
def no_real_windows_side_effects(monkeypatch):
    monkeypatch.setattr(app.powershell, "default_run", _refuse)
    monkeypatch.setattr(app.notify, "_message_box", lambda title, body: None)
    # 잠긴 작업을 기다리는 자동 실행(30초마다 15분)이 테스트를 실제로 멈춰 세우지 않게
    monkeypatch.setattr(app.schedule, "_sleep", lambda seconds: None)


@pytest.fixture
def fake_ps(monkeypatch):
    fake = FakePowerShell()
    monkeypatch.setattr(app.powershell, "default_run", fake)
    return fake
