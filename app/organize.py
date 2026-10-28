"""파일 이동(삭제·덮어쓰기 없음), 이름 충돌 처리, 확인필요 .txt 작성."""
import re
import shutil
import string
from pathlib import Path

from app import config

FORBIDDEN = re.compile(r'[\\/:*?"<>|]')
MISSING = "미상"


class _Missing:
    """값을 못 찾은 필드. 어떤 서식을 줘도 '미상' 으로 찍힌다."""

    def __format__(self, spec):
        return MISSING


class _SafeFormatter(string.Formatter):
    def format_field(self, value, spec):
        text = super().format_field(value, spec)
        return FORBIDDEN.sub("_", text).strip().rstrip(".") or MISSING


def render_filename(result) -> Path:
    """규칙의 파일명 템플릿을 채운다. 결과는 정리문서/ 아래 상대 경로."""
    values = {k: (_Missing() if v is None else v) for k, v in result.fields.items()}
    return Path(_SafeFormatter().format(result.doc_type.filename_tpl, **values))


def target_path(result, work_dir) -> Path:
    return Path(work_dir) / config.ORGANIZED / render_filename(result)


def _free_name(path: Path) -> Path:
    """같은 이름이 있으면 _2, _3 … 을 붙인다."""
    if not path.exists():
        return path
    n = 2
    while (candidate := path.with_name(f"{path.stem}_{n}{path.suffix}")).exists():
        n += 1
    return candidate


def _move(src, dst) -> Path:
    dst = _free_name(Path(dst))
    dst.parent.mkdir(parents=True, exist_ok=True)
    shutil.move(str(src), str(dst))
    return dst


def move_success(src, dst) -> Path:
    return _move(src, dst)


def move_duplicate(src, work_dir) -> Path:
    return _move(src, Path(work_dir) / config.INBOX / config.DUPLICATES / Path(src).name)


def move_review(src, work_dir, reason, text_head) -> Path:
    review_dir = Path(work_dir) / config.REVIEW
    review_dir.mkdir(parents=True, exist_ok=True)
    # PDF 와 .txt 가 같은 이름 짝이 되도록 둘 다 비어 있는 이름을 고른다
    dst, n = review_dir / Path(src).name, 2
    while dst.exists() or dst.with_suffix(".txt").exists():
        dst = review_dir / f"{Path(src).stem}_{n}{Path(src).suffix}"
        n += 1
    final = _move(src, dst)
    note = [f"사유: {reason}", "", "규칙을 고친 뒤 이 PDF 를 받은문서/ 로 옮기고 이 .txt 는 지우세요."]
    if text_head:
        note += ["", "--- 1쪽 텍스트 앞 30줄 ---", text_head]
    final.with_suffix(".txt").write_text("\n".join(note) + "\n", encoding="utf-8")
    return final
