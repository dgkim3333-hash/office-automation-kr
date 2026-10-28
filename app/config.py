"""경로·파일 이름 상수. 작업 폴더 위치는 저장소 루트의 작업폴더.txt 에서 읽는다."""
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parent.parent
WORK_DIR_FILE = "작업폴더.txt"
EXAMPLE_RULES = Path("rules") / "규칙.예시.yaml"

INBOX = "받은문서"
DUPLICATES = "중복"  # 받은문서/중복
ORGANIZED = "정리문서"
REVIEW = "확인필요"
LEDGER_FILE = "누적장부.xlsx"
LOG_FILE = "처리로그.csv"
RULES_FILE = "규칙.yaml"

NEED_INIT = "설치.bat를 먼저 실행하세요"


class ConfigError(Exception):
    pass


def work_dir_file(repo_root=REPO_ROOT) -> Path:
    return Path(repo_root) / WORK_DIR_FILE


def work_dir(repo_root=REPO_ROOT) -> Path:
    """작업 폴더 경로. 설정이 없거나 폴더가 사라졌으면 ConfigError."""
    txt = work_dir_file(repo_root)
    if not txt.exists():
        raise ConfigError(NEED_INIT)
    path = Path(txt.read_text(encoding="utf-8").strip())
    if not path.is_dir():
        raise ConfigError(f"{NEED_INIT} (작업 폴더가 없습니다: {path})")
    return path
