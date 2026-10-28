"""작업(job) 모델: 작업 폴더 만들기, 설정 저장·읽기, 번호·시각 배정, 찾기.

작업 하나 = `작업/NN_제목/` 폴더 하나. 설정은 그 폴더의 `작업설정.yaml` 에 저장한다.
"""
import os
import re
from dataclasses import dataclass, field
from pathlib import Path

import yaml

from app.local_guard import LocalOnlyError, check_work_dir
from app.organize import FORBIDDEN

JOBS_DIR = "작업"
JOB_FILE = "작업설정.yaml"
LAST_NUMBER_FILE = "마지막번호.txt"
RULES_FILE = "규칙.yaml"
OUTPUT_MODES = ("회사양식", "누적장부", "둘다")
JOB_SUBFOLDERS = ("승인대기", "확인필요", "검산")
TITLE_MAX = 30
FIRST_SLOT_MINUTES = 16 * 60
SLOT_STEP_MINUTES = 20

_NUMBER_PREFIX = re.compile(r"^(\d+)_")


class JobError(Exception):
    """작업을 만들거나 찾을 수 없는 경우. 메시지는 사용자에게 그대로 보인다."""


@dataclass
class Job:
    folder: Path
    number: int
    title: str
    input_dir: Path
    output_dir: Path
    output_mode: str
    run_time: str
    scheduler: str = ""
    paused: bool = False
    approvals: dict[str, int] = field(default_factory=dict)
    fingerprints: dict[str, str] = field(default_factory=dict)
    last_run: str = ""  # YYYY-MM-DD

    @property
    def rules_path(self) -> Path:
        return self.folder / RULES_FILE

    @property
    def name(self) -> str:
        return self.folder.name


def slot_time(number: int) -> str:
    """작업 번호 n 의 매일 실행 시각: 16:00 + 20분 x (n - 1)."""
    minutes = (FIRST_SLOT_MINUTES + SLOT_STEP_MINUTES * (number - 1)) % (24 * 60)
    return f"{minutes // 60:02d}:{minutes % 60:02d}"


def _folder_title(title: str) -> str:
    text = FORBIDDEN.sub("_", title).strip().rstrip(".").strip()
    return text[:TITLE_MAX].rstrip().rstrip(".")


def _jobs_root(work: Path) -> Path:
    return Path(work) / JOBS_DIR


def _check_folder(path, label: str) -> Path:
    path = Path(path)
    try:
        check_work_dir(path)
    except LocalOnlyError as e:
        raise JobError(str(e)) from e
    if not path.is_dir():
        raise JobError(f"{label} 폴더가 없습니다: {path}")
    return path


def _next_number(root: Path) -> int:
    """지금까지 쓴 가장 큰 번호 + 1. 지운 작업의 번호는 마지막번호.txt 가 기억해 다시 쓰지 않는다."""
    numbers = []
    if root.is_dir():
        for child in root.iterdir():
            m = _NUMBER_PREFIX.match(child.name)
            if child.is_dir() and m:
                numbers.append(int(m.group(1)))
        try:
            numbers.append(int((root / LAST_NUMBER_FILE).read_text(encoding="utf-8").strip()))
        except (OSError, ValueError):
            pass
    return max(numbers, default=0) + 1


def _atomic_write(path: Path, text: str) -> None:
    """같은 폴더의 임시 파일에 쓴 뒤 바꿔치기해서, 쓰다 끊겨도 기존 파일이 잘리지 않게 한다."""
    tmp = path.with_name(path.name + ".tmp")
    tmp.write_text(text, encoding="utf-8")
    os.replace(tmp, path)


def create_job(work, title: str, input_dir, output_dir, output_mode: str) -> Job:
    if output_mode not in OUTPUT_MODES:
        raise JobError(f"출력 방식 '{output_mode}' 을 쓸 수 없습니다. 고를 수 있는 것: {', '.join(OUTPUT_MODES)}")
    title = title.strip()
    safe = _folder_title(title)
    if not safe:
        raise JobError("작업 제목이 비어 있습니다.")
    input_dir = _check_folder(input_dir, "입력")
    output_dir = _check_folder(output_dir, "출력")
    root = _jobs_root(work)
    number = _next_number(root)
    folder = root / f"{number:02d}_{safe}"
    folder.mkdir(parents=True)
    for sub in JOB_SUBFOLDERS:
        (folder / sub).mkdir()
    _atomic_write(root / LAST_NUMBER_FILE, str(number))
    job = Job(
        folder=folder,
        number=number,
        title=title,
        input_dir=input_dir,
        output_dir=output_dir,
        output_mode=output_mode,
        run_time=slot_time(number),
    )
    save_job(job)
    return job


def save_job(job: Job) -> None:
    data = {
        "번호": job.number,
        "제목": job.title,
        "입력폴더": str(job.input_dir),
        "출력폴더": str(job.output_dir),
        "출력방식": job.output_mode,
        "실행시각": job.run_time,
        "스케줄러": job.scheduler,
        "일시중지": job.paused,
        "승인": dict(job.approvals),
        "지문": dict(job.fingerprints),
        "마지막실행": job.last_run,
    }
    text = yaml.safe_dump(data, allow_unicode=True, sort_keys=False)
    _atomic_write(job.folder / JOB_FILE, text)


def load_job(folder: Path) -> Job:
    folder = Path(folder)
    path = folder / JOB_FILE
    try:
        data = yaml.safe_load(path.read_text(encoding="utf-8"))
    except UnicodeDecodeError as e:
        raise JobError(f"작업 설정을 UTF-8 로 읽지 못했습니다(메모장에서 다른 인코딩으로 저장했는지 확인): {path}") from e
    except (OSError, yaml.YAMLError) as e:
        raise JobError(f"작업 설정을 읽지 못했습니다: {path} ({e})") from e
    if not isinstance(data, dict):
        raise JobError(f"작업 설정 형식이 올바르지 않습니다: {path}")

    def need(key):
        if data.get(key) is None:
            raise KeyError(key)
        return data[key]

    try:
        return Job(
            folder=folder,
            number=int(need("번호")),
            title=str(need("제목")),
            input_dir=Path(need("입력폴더")),
            output_dir=Path(need("출력폴더")),
            output_mode=str(need("출력방식")),
            run_time=str(need("실행시각")),
            scheduler=str(data.get("스케줄러") or ""),
            paused=bool(data.get("일시중지", False)),
            approvals={str(k): int(v) for k, v in (data.get("승인") or {}).items()},
            fingerprints={str(k): str(v) for k, v in (data.get("지문") or {}).items()},
            last_run=str(data.get("마지막실행") or ""),
        )
    except KeyError as e:
        raise JobError(f"작업 설정에 {e.args[0]} 항목이 없습니다: {path}") from e
    except (ValueError, TypeError, AttributeError) as e:
        raise JobError(f"작업 설정의 값이 올바르지 않습니다: {path} ({e})") from e


def _scan(work) -> tuple[list[Job], list[tuple[Path, str]]]:
    """작업 폴더들을 읽어 (읽힌 작업, 깨진 작업) 으로 나눈다. 깨진 하나가 나머지를 막지 않는다."""
    root = _jobs_root(work)
    jobs: list[Job] = []
    broken: list[tuple[Path, str]] = []
    if root.is_dir():
        for child in sorted(root.iterdir(), key=lambda p: p.name):
            if not (child.is_dir() and (child / JOB_FILE).is_file()):
                continue
            try:
                jobs.append(load_job(child))
            except JobError as e:
                broken.append((child, str(e)))
    return sorted(jobs, key=lambda j: j.number), broken


def list_jobs(work) -> list[Job]:
    """읽히는 작업만 번호 순으로. 깨진 작업은 broken_jobs 로 따로 본다."""
    return _scan(work)[0]


def broken_jobs(work) -> list[tuple[Path, str]]:
    """설정을 읽지 못한 작업의 (폴더, 이유)."""
    return _scan(work)[1]


def find_job(work, key: str) -> Job:
    key = key.strip()
    jobs, broken = _scan(work)
    if key.isdigit():
        for job in jobs:
            if job.number == int(key):
                return job
    for job in jobs:
        if key in (job.name, job.title):
            return job
    for folder, reason in broken:
        m = _NUMBER_PREFIX.match(folder.name)
        if key == folder.name or (key.isdigit() and m and int(m.group(1)) == int(key)):
            raise JobError(f"작업 '{key}' 의 설정이 깨져 있습니다. {reason}")
    existing = ", ".join(j.name for j in jobs) or "없음"
    raise JobError(f"작업 '{key}' 을 찾지 못했습니다. 지금 있는 작업: {existing}")
