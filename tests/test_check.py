"""python -m app check <pdf> — 추출 결과 표만 출력하고 아무것도 바꾸지 않는다."""
import shutil
from pathlib import Path

import pytest

from app.__main__ import main
from app.init import init

REPO = Path(__file__).resolve().parent.parent
SAMPLES = REPO / "samples" / "03_Invoice_DM_PDF"


@pytest.fixture
def setup(tmp_path):
    repo = tmp_path / "repo"
    (repo / "rules").mkdir(parents=True)
    shutil.copy(REPO / "rules" / "규칙.예시.yaml", repo / "rules" / "규칙.예시.yaml")
    work = init(work_dir=tmp_path / "work", repo_root=repo)
    return repo, work


def _snapshot(folder):
    return sorted(str(p.relative_to(folder)) for p in folder.rglob("*"))


def test_check_prints_fields_and_table(setup, capsys):
    repo, work = setup
    before = _snapshot(work)
    code = main(["check", str(SAMPLES / "INV-2026-10231.pdf")], repo_root=repo)
    out = capsys.readouterr().out

    assert code == 0
    assert "| 필드 | 값 | 비고 |" in out
    assert "| 인보이스번호 | INV-2026-10231 |" in out
    assert "| 발행일 | 2026-10-14 |" in out
    assert "| 고객명 | ACME Semiconductor Korea Co., Ltd. |" in out
    assert "| 합계금액 | 115440.00 |" in out
    assert "| 줄 | 품번 | 품명 | 수량 | 단가 | 금액 |" in out
    assert "| 1 | VTS-CDU-300 | Coolant Distribution Unit, 300 kW, N+1 pump | 2 | 48500.00 | 97000.00 |" in out
    assert "Invoice_INV-2026-10231.pdf" in out  # 정리될 경로 미리보기
    assert _snapshot(work) == before  # 작업 폴더·장부 변화 없음
    assert (SAMPLES / "INV-2026-10231.pdf").exists()


def test_check_unknown_shows_reason_and_text_head(setup, capsys):
    repo, _ = setup
    main(["check", str(SAMPLES / "Letter_warranty_notice.pdf")], repo_root=repo)
    out = capsys.readouterr().out
    assert "확인필요" in out and "판별 실패" in out
    assert "Vertex" in out  # 1쪽 텍스트 앞부분


def test_check_missing_field_shows_note(setup, capsys):
    repo, work = setup
    rules = work / "규칙.yaml"
    rules.write_text(rules.read_text(encoding="utf-8").replace('"Bill To:', '"Sold To:'), encoding="utf-8")
    main(["check", str(SAMPLES / "INV-2026-10231.pdf")], repo_root=repo)
    out = capsys.readouterr().out
    assert "| 고객명 | (비어 있음) | 못 찾음 |" in out


def test_check_bad_rules_prints_korean_error(setup, capsys):
    repo, work = setup
    (work / "규칙.yaml").write_text("문서종류: 없음", encoding="utf-8")
    code = main(["check", str(SAMPLES / "INV-2026-10231.pdf")], repo_root=repo)
    assert code == 1
    assert "규칙" in capsys.readouterr().out


def test_without_init_says_run_setup(tmp_path, capsys):
    code = main(["check", str(SAMPLES / "INV-2026-10231.pdf")], repo_root=tmp_path)
    assert code == 1
    assert "설치.bat를 먼저 실행하세요" in capsys.readouterr().out
