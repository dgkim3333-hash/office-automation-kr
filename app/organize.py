"""파일 이름: 템플릿 채우기 · 못 쓰는 글자 바꾸기 · 이름 충돌 처리."""
import re
import string
from pathlib import Path

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


def render_text(template, fields) -> str:
    """{필드} 자리를 채운다. 경로에 못 쓰는 글자는 _, 빈 값은 '미상'."""
    values = {k: (_Missing() if v is None else v) for k, v in fields.items()}
    return _SafeFormatter().format(template, **values)


def render_filename(result) -> Path:
    """규칙의 파일명 템플릿을 채운다(상대 경로)."""
    return Path(render_text(result.doc_type.filename_tpl, result.fields))


def free_name(path: Path, reserved=()) -> Path:
    """같은 이름이 있거나(reserved 포함) 쓰기로 한 이름이면 _2, _3 … 을 붙인다."""
    path = Path(path)
    taken = lambda p: p.exists() or p in reserved  # noqa: E731
    if not taken(path):
        return path
    n = 2
    while taken(candidate := path.with_name(f"{path.stem}_{n}{path.suffix}")):
        n += 1
    return candidate

