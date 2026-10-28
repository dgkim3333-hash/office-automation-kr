"""설치.bat 본체: 작업 폴더를 고르고 빈 구조(작업\\ · 고객마스터.xlsx)를 만든다. 폴더 고르기 창도 여기 있다."""
from pathlib import Path

from app import config
from app.jobs import JOBS_DIR
from app.local_guard import LocalOnlyError, check_work_dir
from app.master import write_example


class InitError(Exception):
    pass


def pick_folder(title: str) -> str:
    """Windows 폴더 선택 창. 쓸 수 없거나 취소하면 빈 문자열."""
    try:
        import tkinter
        from tkinter import filedialog
    except ImportError:
        return ""
    root = tkinter.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    try:
        return filedialog.askdirectory(title=title) or ""
    finally:
        root.destroy()


def _pick_work_folder() -> str:
    return pick_folder("작업 폴더를 고르세요 (예: 문서\\사무자동화)")


def _choose(repo_root, ask, pick) -> Path:
    txt = config.work_dir_file(repo_root)
    if txt.exists():
        current = Path(txt.read_text(encoding="utf-8").strip())
        print(f"현재 작업 폴더: {current}")
        if ask("바꾸시겠습니까? (y/n): ").strip().lower() not in ("y", "yes", "예", "ㅇ"):
            return current
    chosen = pick()
    if not chosen:
        chosen = ask("경로를 직접 입력하세요: ").strip().strip('"')
    if not chosen:
        raise InitError("작업 폴더를 고르지 않았습니다. 설치.bat 를 다시 실행하세요.")
    return Path(chosen)


def _is_inside(path: Path, parent: Path) -> bool:
    return path == parent or parent in path.parents


def init(work_dir=None, repo_root=config.REPO_ROOT, ask=input, pick=_pick_work_folder) -> Path:
    """작업 폴더에 작업\\ 과 고객마스터.xlsx(없을 때만)를 만들고 위치를 작업폴더.txt 에 적는다.
    작업마다의 폴더 · 규칙은 job new 가 만든다."""
    repo_root = Path(repo_root).resolve()
    if work_dir is None:
        work_dir = _choose(repo_root, ask, pick)
    work_dir = Path(work_dir).resolve()

    if _is_inside(work_dir, repo_root):
        raise InitError(
            "저장소 폴더 안은 작업 폴더로 쓸 수 없습니다. ZIP 을 새로 받을 때 데이터가 지워질 수 있습니다.\n"
            "문서 폴더 같은 다른 곳을 고르세요."
        )

    try:
        check_work_dir(work_dir)
    except LocalOnlyError as e:
        raise InitError(str(e)) from None

    (work_dir / JOBS_DIR).mkdir(parents=True, exist_ok=True)
    master = work_dir / config.MASTER_FILE
    if not master.exists():  # 있으면 사용자가 채운 것이므로 건드리지 않는다
        write_example(master)
    config.work_dir_file(repo_root).write_text(str(work_dir) + "\n", encoding="utf-8")

    print(f"작업 폴더: {work_dir}")
    print("이제 04장으로 가세요.")
    return work_dir
