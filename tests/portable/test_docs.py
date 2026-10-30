"""무설치판 문서(09장 · README · 99장 · 작업지시서 · CLAUDE.md)가 설계와 코드에 맞는지 검사(작업 12)."""
import re
from pathlib import Path

import pytest

REPO = Path(__file__).resolve().parent.parent.parent
CHAPTER = REPO / "09_무설치판.md"
README = REPO / "README.md"
TROUBLE = REPO / "99_막혔을_때.md"
WORK_ORDER = REPO / "작업지시서.md"
CLAUDE_MD = REPO / "CLAUDE.md"
PORTABLE = REPO / "무설치판"


def _text(path: Path) -> str:
    return path.read_text(encoding="utf-8")


def _ps_string(file: str, name: str) -> str:
    """무설치판 .ps1 의 `$script:<name> = '…'`(또는 "…") 글자. 문서가 화면 문구를 코드와 같게 옮겼는지 보려고."""
    text = _text(PORTABLE / "lib" / file)
    match = re.search(rf"^\$script:{name}\s*=\s*(['\"])(.*?)\1\s*(#.*)?$", text, re.M)
    assert match, f"{file} 에 $script:{name} 이 없음"
    return match.group(2)


def _section(text: str, title: str) -> str:
    """title 이 든 첫 '## ' 제목 줄부터 다음 '## ' 앞(또는 끝)까지."""
    match = re.search(rf"^## [^\n]*{re.escape(title)}", text, re.M)
    assert match, f"'## …{title}…' 절이 없음"
    end = text.find("\n## ", match.end())
    return text[match.start():] if end < 0 else text[match.start():end]


# --- 09장 ---

def test_chapter_09_says_original_leaves_friend_company():
    """설계 3절: 이 문장을 그대로 적는다."""
    text = _text(CHAPTER)
    assert "원본 1개가 친구 회사 밖(대표)으로 나가고 Claude 에게도 한 번 간다" in text
    assert "보내기 전에 회사에 확인" in text


def test_chapter_09_has_setup_bat_and_daily_parts():
    text = _text(CHAPTER)
    for word in ("준비.bat", "지금 실행 - ", "16:00", "20분", "로그인", "사무자동화 무설치판 놓친 작업 실행"):
        assert word in text, word


def test_chapter_09_pdf_warning_is_users_own_click():
    """「다시 표시 안 함」은 친구가 워드에서 직접 체크한다. 레지스트리 · 보안 설정을 고치라고 하지 않는다."""
    text = _text(CHAPTER)
    assert "다시 표시 안 함" in text
    guide = _ps_string("Setup.ps1", "PdfWarningGuide")
    assert guide in text, "준비 화면의 안내 글자와 같게"
    for word in ("레지스트리", "regedit", "DisableConvertPdfWarning", "보안 설정을 바꾸"):
        assert word not in text, word


@pytest.mark.parametrize("path", [CHAPTER, TROUBLE], ids=lambda p: p.name)
def test_no_code_tab_keeps_everything_inside_claim(path):
    """「Code 탭이면 밖으로 안 나간다」 류 문장이 없다."""
    text = _text(CHAPTER) if path == CHAPTER else _section(_text(TROUBLE), "무설치판")
    for pattern in (r"Code 탭(이면|이라서|이니까|이라면)", r"밖으로 안 나", r"밖으로 나가지 않", r"아무것도 나가지 않"):
        assert not re.search(pattern, text), f"{path.name}: {pattern}"


def test_chapter_09_quotes_setup_messages_like_code():
    text = _text(CHAPTER)
    for name in ("PdfDifferent", "ExcelDifferent"):
        assert _ps_string("Setup.ps1", name) in text, name
    assert "폴더를 고르지 않아 준비를 멈췄습니다" in _text(PORTABLE / "lib" / "Setup.ps1")
    assert "폴더를 고르지 않아 준비를 멈췄습니다" in text


def test_chapter_09_differences_from_installed_edition():
    section = _section(_text(CHAPTER), "지금 판과 다른 점")
    for word in ("승인", "10건", "고객마스터", "Outlook"):
        assert word in section, word


def test_chapter_09_explains_rulings():
    text = _text(CHAPTER)
    # 마크업지문 = 마크업표 파일 전체 지문(지금 판과 뜻이 다름)
    at = text.index("마크업지문")
    assert "지금 판" in text[at:at + 600] and "마크업표" in text[at:at + 600]
    # 실행 중 숨은 엑셀에 xlsx 가 열릴 수 있음 → 그 엑셀은 끝내지 않고 화면에 보이게 넘겨 준다(최종 검토 I3), 시간 초과면 닫힐 수 있음
    assert "숨은 엑셀" in text
    at = text.index("숨은 엑셀")
    assert "화면에 보이게 넘겨" in text[at:at + 900] and "시간 초과" in text[at:at + 900]
    # 제목을 바꾸면 새 누적장부, 옛 장부는 그대로
    assert "새 누적장부" in text and "옛 누적장부" in text
    # bundle-portable: --작업 없으면 멈추지 않은 작업만, 새 빈 폴더, 실패하면 지우고 새 폴더
    assert "bundle-portable" in text and "--작업" in text and "멈춘 작업" in text
    assert "비어 있는 새 폴더" in text and "반쯤 만든" in text
    # 자체 시험: 엑셀이 다르면 준비가 멈춤, PDF 가 다르면 PDF 만 꺼지고 준비는 계속
    assert "준비가 멈춥니다" in text and "준비는 계속" in text
    # 업데이트: 덮어쓰고 준비.bat 다시, 번호로 맞춤
    update = _section(text, "업데이트")
    assert "덮어" in update and "준비.bat" in update and "번호" in update


def test_chapter_09_commands_use_real_portable_commands():
    """문서의 `-Command X` 는 무설치판.ps1 에 있는 명령이다."""
    dispatcher = _text(PORTABLE / "무설치판.ps1")
    for path in (CHAPTER, TROUBLE, WORK_ORDER):
        for name in re.findall(r"-Command ([^\s`\"'|)]+)", _text(path)):
            assert f"'{name}' = {{" in dispatcher, f"{path.name}: 없는 무설치판 명령 {name}"


# --- README · 99장 · 작업지시서 · CLAUDE.md ---

def test_readme_lists_chapter_09_between_08_and_99():
    text = _text(README)
    at_08 = text.index("](08_선택_Outlook_자동수집.md)")
    at_09 = text.index("](09_무설치판.md)")
    at_99 = text.index("](99_막혔을_때.md)")
    assert at_08 < at_09 < at_99


def test_chapter_99_portable_section():
    text = _text(TROUBLE)
    section = _section(text, "무설치판")
    assert "09_무설치판.md" in section
    assert _ps_string("Office.ps1", "WordTimeoutText") in section  # 「워드가 응답하지 않음」
    assert "숨은 엑셀에 이 프로그램이 열지 않은 통합 문서" in section   # 넘겨 준 기록(Office.ps1 OfficeLeftText)
    assert "새로 띄운 엑셀에 이미 다른 통합 문서가 열려 있어" in section   # 우리 것이 아닌 엑셀(OfficeForeignText)
    assert "다르게 읽습니다" in section
    assert "스크립트" in section
    it = _section(text, "IT 담당자")
    assert "PowerShell 스크립트(-ExecutionPolicy Bypass) 실행" in it


def test_work_order_portable_section_at_end():
    text = _text(WORK_ORDER)
    assert text.index("## 기존 작업을 고칠 때") < text.index("## 무설치판으로 넘길 작업")
    section = _section(text, "무설치판으로 넘길 작업")
    for word in ("export-portable", "-Command 점검", "bundle-portable", "--작업", "둘 다", "고객마스터"):
        assert word in section, word


def test_claude_md_portable_rule():
    lines = [line for line in _text(CLAUDE_MD).splitlines() if "무설치판" in line]
    assert any("고객마스터" in line and "마크업표" in line and "열지 않" in line for line in lines)


@pytest.mark.parametrize("path", [CHAPTER, README, TROUBLE, WORK_ORDER, CLAUDE_MD], ids=lambda p: p.name)
def test_no_company_initials(path):
    initials = chr(68) + chr(71)  # 글자 그대로 저장소에 남기지 않는다
    assert not re.search(rf"\b{initials}\b", _text(path))
