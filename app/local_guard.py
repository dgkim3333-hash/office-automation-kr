"""이 PC 전용 가드.

- Windows PC 가 아니면 멈춘다(클라우드 세션·Cowork 작업은 원격 서버에서 돈다).
- 클라우드 실행 환경 표시(환경 변수)가 있으면 멈춘다.
- 작업 폴더가 동기화 폴더(OneDrive 등)나 네트워크 위치면 멈춘다. 장부·PDF 가 PC 밖으로 복사되기 때문이다.
- 실행하는 동안 이 프로그램의 인터넷 연결을 코드로 막는다.
"""
import os
import socket
import sys
from pathlib import Path

# 클라우드에서 돌 때 켜지는 것으로 알려진 환경 변수. Windows 검사가 1차, 이것은 보조다.
CLOUD_MARKERS = ("CLAUDE_CODE_REMOTE", "CODESPACES", "GITHUB_ACTIONS")
SYNC_FOLDER_NAMES = ("onedrive", "dropbox", "google drive", "icloud")
DRIVE_REMOTE = 4  # GetDriveTypeW: 네트워크 드라이브

BANNER = (
    "[이 PC 전용] 이 프로그램은 지금 이 컴퓨터 안에서만 돌고, 인터넷으로 아무것도 보내지 않습니다.\n"
    "             실행은 run.bat 또는 Claude Desktop 의 Code 탭(이 PC)에서만 하세요.\n"
    "             Cowork·클라우드 세션에는 회사 PDF 를 맡기지 마세요. Claude 에게 보여 준 것만 Claude 에게 갑니다."
)


class LocalOnlyError(Exception):
    """이 PC 밖에서 돌거나 PC 밖으로 데이터가 나갈 수 있는 상황."""


def check_environment():
    if sys.platform != "win32":
        raise LocalOnlyError(
            "이 프로그램은 내 Windows PC 에서만 동작합니다. 지금은 Windows 가 아닙니다.\n"
            "클라우드 세션이나 Cowork 에서 실행하지 말고, 내 PC 에서 run.bat 을 누르세요."
        )
    found = [name for name in CLOUD_MARKERS if os.environ.get(name)]
    if found:
        raise LocalOnlyError(
            f"클라우드 실행 환경으로 보입니다({', '.join(found)}). 이 프로그램은 내 PC 에서만 실행합니다.\n"
            "Claude Desktop 의 Code 탭에서 '로컬(Local)' 환경으로 열었는지 확인하세요."
        )


def _drive_type(root: str) -> int:
    import ctypes

    return ctypes.windll.kernel32.GetDriveTypeW(root)


def check_work_dir(path):
    path = Path(path)
    text = str(path)
    if text.startswith("\\\\"):
        raise LocalOnlyError(f"작업 폴더가 네트워크 위치입니다: {text}\n내 PC 의 로컬 디스크(예: C:\\사무자동화)를 쓰세요.")
    for part in path.parts:
        lowered = part.lower()
        if any(lowered.startswith(name) for name in SYNC_FOLDER_NAMES):
            raise LocalOnlyError(
                f"작업 폴더가 동기화 폴더({part}) 안에 있습니다: {text}\n"
                "장부와 PDF 가 클라우드로 자동 업로드되므로 쓸 수 없습니다. 예: C:\\사무자동화 처럼 로컬 디스크를 쓰세요."
            )
    if sys.platform == "win32" and path.drive and _drive_type(path.drive + "\\") == DRIVE_REMOTE:
        raise LocalOnlyError(f"작업 폴더가 네트워크 드라이브({path.drive})에 있습니다. 내 PC 의 로컬 디스크를 쓰세요.")


def _refuse(*args, **kwargs):
    raise LocalOnlyError("이 프로그램은 인터넷에 연결하지 않습니다(이 PC 전용 가드).")


def block_network():
    """이 파이썬 프로세스의 인터넷 연결을 막는다. 되돌리는 함수를 돌려준다."""
    saved = {
        (socket.socket, "connect"): socket.socket.connect,
        (socket.socket, "connect_ex"): socket.socket.connect_ex,
        (socket, "create_connection"): socket.create_connection,
        (socket, "getaddrinfo"): socket.getaddrinfo,
    }
    for (owner, name) in saved:
        setattr(owner, name, _refuse)

    def restore():
        for (owner, name), original in saved.items():
            setattr(owner, name, original)

    return restore
