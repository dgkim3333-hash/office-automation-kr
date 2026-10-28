"""체험(가짜 샘플 체험)은 뺐다(2026-10-28 결정). 친구는 바로 자기 파일로 시작한다."""
from pathlib import Path

import pytest

from app.__main__ import _parser

REPO = Path(__file__).resolve().parent.parent


def test_demo_bat_is_gone():
    assert not (REPO / "체험.bat").exists()


def test_demo_command_is_gone():
    with pytest.raises(SystemExit):
        _parser().parse_args(["demo"])


def test_install_does_not_point_to_demo():
    assert "체험" not in (REPO / "install.ps1").read_text(encoding="utf-8")
