"""경로 · 파일 이름 상수. 작업 폴더 위치는 저장소 루트의 작업폴더.txt 에서 읽는다."""
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parent.parent
WORK_DIR_FILE = "작업폴더.txt"
SKELETON_RULES = Path("rules") / "규칙.뼈대.yaml"  # job new 가 새 작업의 규칙.yaml 로 복사하는 빈 뼈대

MASTER_FILE = "고객마스터.xlsx"  # 작업 폴더 맨 위. 작업들이 같이 쓴다

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
