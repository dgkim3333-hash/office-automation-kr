"""문서(README·01~99장) 작성 원칙 검사 — 설계 11절."""
import re
from pathlib import Path

import pytest

REPO = Path(__file__).resolve().parent.parent
# 설계 4절의 장 목록. 아직 안 쓴 장으로 가는 링크는 허용하되, 마지막에는 전부 있어야 한다.
PLANNED = [
    "01_구조_이해.md",
    "02_기본_설치.md",
    "03_프로그램_설치.md",
    "04_작업_만들기.md",
    "05_내_양식_가르치기.md",
    "06_매일_쓰기.md",
    "07_프로젝트와_예정됨.md",
    "08_선택_Outlook_자동수집.md",
    "99_막혔을_때.md",
]
LINK = re.compile(r"\]\(([^)#\s]+)(#[^)]*)?\)")


def _chapters():
    return sorted(p for p in REPO.glob("[0-9][0-9]_*.md"))


def _docs():
    return [REPO / "README.md", *_chapters()] if (REPO / "README.md").exists() else _chapters()


def _body_lines(path):
    """제목(# …) 다음부터 빈 줄을 뺀 줄들."""
    lines = [line for line in path.read_text(encoding="utf-8").splitlines() if line.strip()]
    assert lines[0].startswith("# "), f"{path.name}: 첫 줄은 # 제목"
    return lines[1:]


def test_readme_first_box_says_code_tab():
    first_box = []
    for line in _body_lines(REPO / "README.md"):
        if not line.startswith(">"):
            break
        first_box.append(line)
    box = "\n".join(first_box)
    assert "Code 탭" in box and "채팅" in box and "PC 밖" in box


@pytest.mark.parametrize("path", _chapters(), ids=lambda p: p.name)
def test_chapter_starts_with_what_you_will_see(path):
    head = "\n".join(_body_lines(path)[:8])
    assert "이 장이 끝나면 보이는 것" in head


@pytest.mark.parametrize("path", _chapters(), ids=lambda p: p.name)
def test_chapter_ends_with_next(path):
    tail = "\n".join(_body_lines(path)[-3:])
    assert "다음 →" in tail


def test_chapter_02_first_line_is_code_tab():
    path = REPO / "02_기본_설치.md"
    if path.exists():
        first = _body_lines(path)[0]
        assert "Code 탭" in first and "채팅" in first


def test_chapter_99_first_line_is_plan():
    path = REPO / "99_막혔을_때.md"
    if path.exists():
        first = _body_lines(path)[0]
        assert "Code 탭" in first and "요금제" in first


@pytest.mark.parametrize("path", _docs(), ids=lambda p: p.name)
def test_one_command_per_block_with_description_above(path):
    lines = path.read_text(encoding="utf-8").splitlines()
    i = 0
    while i < len(lines):
        if lines[i].strip() in ("```powershell", "```bat", "```cmd"):
            body = []
            j = i + 1
            while lines[j].strip() != "```":
                if lines[j].strip():
                    body.append(lines[j])
                j += 1
            assert len(body) == 1, f"{path.name}:{i + 1} 명령 블록에 {len(body)}줄"
            above = next(line for line in reversed(lines[:i]) if line.strip())
            assert not above.startswith(("#", "```", "|")), f"{path.name}:{i + 1} 블록 위에 설명 한 줄이 없음"
            i = j
        i += 1


@pytest.mark.parametrize("path", _docs(), ids=lambda p: p.name)
def test_relative_links_exist(path):
    text = path.read_text(encoding="utf-8")
    for target, _ in LINK.findall(text):
        if re.match(r"^[a-z]+:", target):  # http: · https: · mailto:
            continue
        if (path.parent / target).exists():
            continue
        assert target in PLANNED, f"{path.name}: 없는 파일로 가는 링크 {target}"


def test_all_planned_chapters_exist():
    missing = [name for name in PLANNED if not (REPO / name).exists()]
    assert missing == [], f"아직 없는 장: {missing}"