"""작업지시서.md · CLAUDE.md · 문서가 작업 방식(설계 work-orders 4 · 12 · 13절)과 맞는지 검사."""
import re
import shlex
from pathlib import Path

import pytest

from app.__main__ import _parser

REPO = Path(__file__).resolve().parent.parent
WORK_ORDER = REPO / "작업지시서.md"
CLAUDE_MD = REPO / "CLAUDE.md"

# 설계 4절 질문 순서. 작업지시서.md 에 이 순서로 나와야 한다
QUESTIONS = [
    "무엇을 만들어 드릴까요? 제목을 입력해 주세요.",
    "입력할 파일(엑셀 또는 PDF)의 예를 하나 넣어 주세요.",
    "출력은 회사 양식, 누적장부, 둘 다 중 무엇으로 할까요?",
    'pick-folder "입력 폴더를 고르세요"',
    'pick-folder "출력 폴더를 고르세요"',
    "job new",
    "--작업",
    "job resume",
    "처음 10건은 승인대기로 갑니다.",
    "다음 작업도 만드실까요, 그만하실까요?",
]


def _text(path: Path) -> str:
    return path.read_text(encoding="utf-8")


def _docs():
    chapters = sorted(REPO.glob("[0-9][0-9]_*.md"))
    return [REPO / "README.md", CLAUDE_MD, WORK_ORDER, *chapters]


def test_work_order_has_all_questions_in_order():
    text = _text(WORK_ORDER)
    at = 0
    for q in QUESTIONS:
        found = text.find(q, at)
        assert found >= 0, f"작업지시서.md 에 순서대로 없음: {q}"
        at = found + len(q)


def test_work_order_warns_original_goes_to_claude_before_asking_file():
    text = _text(WORK_ORDER)
    ask_file = text.find(QUESTIONS[1])
    warning = text.rfind("Anthropic", 0, ask_file)
    assert warning >= 0 and "원본" in text[warning - 200:ask_file], "예시 파일을 받기 전에 원본이 Claude 에게 간다는 안내"


def test_work_order_tells_friend_to_fill_master_and_markup_themselves():
    text = _text(WORK_ORDER)
    for name in ("고객마스터.xlsx", "마크업표.xlsx"):
        assert name in text
    assert "직접" in text


def test_claude_md_points_to_work_order():
    text = _text(CLAUDE_MD)
    assert "작업지시서대로 새 작업 만들어 줘" in text
    assert "작업지시서.md" in text


def test_claude_md_lists_files_claude_never_opens():
    text = _text(CLAUDE_MD)
    for name in ("고객마스터.xlsx", "마크업표.xlsx", "누적장부", "승인대기", "입력 폴더", "출력 폴더"):
        assert name in text, name
    assert "Code 탭이면" in text  # 「Code 탭이면 밖으로 안 나간다」고 말하지 않는다는 약속은 그대로
    for command in ("run", "approve", "catch-up"):
        assert f"`{command}`" in text, command


@pytest.mark.parametrize("name", ["01_구조_이해.md", "04_작업_만들기.md"])
def test_docs_mention_original_goes_to_claude(name):
    text = _text(REPO / name)
    assert "원본" in text and "Anthropic" in text


@pytest.mark.parametrize("path", _docs(), ids=lambda p: p.name)
def test_no_old_single_flow_words(path):
    """옛 흐름(받은문서 · 정리문서 · 수집.bat · 04_처음_처리하기)이 문서에 남지 않는다."""
    text = _text(path)
    for word in ("받은문서", "정리문서", "수집.bat", "04_처음_처리하기", "C:\\사무자동화\\출력"):
        assert word not in text, f"{path.name}: 옛 흐름 낱말 '{word}'"


_COMMAND = re.compile(r"python -m app((?: [^`\n|]*)?)")


@pytest.mark.parametrize("path", _docs(), ids=lambda p: p.name)
def test_every_app_command_in_docs_is_real(path):
    """문서에 적은 `python -m app …` 명령이 실제 명령 해석기에서 통과한다(없는 명령 · 옵션을 적지 않게)."""
    for match in _COMMAND.finditer(_text(path)):
        args = [a.strip('"') for a in shlex.split(match.group(1), posix=False)]
        try:
            _parser().parse_args(args)
        except SystemExit:
            pytest.fail(f"{path.name}: 없는 명령 또는 옵션: python -m app{match.group(1)}")


def test_chapter_04_and_06_explain_email_body_po():
    for name in ("04_작업_만들기.md", "06_매일_쓰기.md"):
        text = _text(REPO / name)
        assert "Ctrl+P" in text and "PDF" in text and "Microsoft Print to PDF" in text, name


def test_chapter_05_markup_and_forms():
    text = _text(REPO / "05_내_양식_가르치기.md")
    for word in ("마크업표.xlsx", "파트넘버", "마크업%", "반올림자리", "0.5%", "10건", "번호"):
        assert word in text, word


def test_chapter_06_daily():
    text = _text(REPO / "06_매일_쓰기.md")
    for word in ("16:00", "20분", "지금 실행 - ", "승인 - ", "대조표", "어느 작업에도 맞지 않는 파일", ".xls",
                 "메시지 창"):
        assert word in text, word


def test_chapter_99_scheduler_and_it():
    text = _text(REPO / "99_막혔을_때.md")
    for word in ("시작프로그램", "없음", "job repair", "IT 담당자", "검산 불일치", "이 작업이 이미 실행 중입니다"):
        assert word in text, word


def test_install_script_points_to_job_chapter_and_removes_old_files():
    text = _text(REPO / "install.ps1")
    assert "04장(작업 만들기)" in text
    legacy = text[text.index("$LegacyFiles"):]
    legacy = legacy[:legacy.index(")")]
    for name in ("04_처음_처리하기.md", "수집.bat", "체험.bat"):
        assert f"'{name}'" in legacy, name


def test_folder_and_mode_changes_use_job_set_not_hand_edits():
    """작업설정.yaml 을 손으로 고치지 않고 job set 을 쓴다(Ruling 27)."""
    claude = _text(CLAUDE_MD)
    assert "job set" in claude and "job set" in _text(WORK_ORDER)
    assert "`입력폴더`" not in claude and "작업설정.yaml` 을 `.bak`" not in claude


def test_docs_say_check_hides_company_data():
    """check 에는 예시 파일의 값과 일치 여부만 나오고 고객마스터 · 마크업표 값은 나오지 않는다(Ruling 25)."""
    for path in (REPO / "01_구조_이해.md", REPO / "04_작업_만들기.md", WORK_ORDER):
        text = _text(path)
        assert "일치 여부" in text, path.name
    assert "C-KR-0417)" not in _text(REPO / "05_내_양식_가르치기.md")


def test_chapter_06_master_same_day_retry():
    text = _text(REPO / "06_매일_쓰기.md")
    assert "| 확인필요(고객마스터에 없는 고객) |" in text and "같은 날 다시 처리" in text
