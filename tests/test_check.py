"""python -m app check <파일> --작업 N: 그 작업 규칙으로 추출 · 검산 결과 표만 출력하고 아무것도 바꾸지 않는다."""
import shutil
from pathlib import Path

import pytest

from app.__main__ import main
from app.init import init
from app.jobs import create_job

REPO = Path(__file__).resolve().parent.parent
SAMPLES = REPO / "samples" / "03_Invoice_DM_PDF"


@pytest.fixture
def setup(tmp_path):
    repo = tmp_path / "repo"
    (repo / "rules").mkdir(parents=True)
    shutil.copy(REPO / "rules" / "규칙.예시.yaml", repo / "rules" / "규칙.예시.yaml")
    work = init(work_dir=tmp_path / "work", repo_root=repo)
    (tmp_path / "in").mkdir()
    (tmp_path / "out").mkdir()
    job = create_job(work, "일일 정리", tmp_path / "in", tmp_path / "out", "둘다")
    shutil.copy(REPO / "rules" / "규칙.예시.yaml", job.rules_path)
    return repo, job.folder


def _check(path, repo):
    return main(["check", str(path), "--작업", "1"], repo_root=repo)


def _snapshot(folder):
    return sorted(str(p.relative_to(folder)) for p in folder.rglob("*"))


def test_check_prints_fields_and_table(setup, capsys):
    repo, work = setup
    before = _snapshot(work)
    code = _check(str(SAMPLES / "INV-2026-10231.pdf"), repo)
    out = capsys.readouterr().out

    assert code == 0
    assert "| 필드 | 값 | 비고 |" in out
    assert "| 인보이스번호 | INV-2026-10231 |" in out
    assert "| 발행일 | 2026-10-14 |" in out
    assert "| 고객명 | ACME Semiconductor Korea Co., Ltd. |" in out
    assert "| 합계금액 | 115440.00 |" in out
    assert "| 줄 | 품번 | 품명 | 수량 | 단가 | 금액 |" in out
    assert "| 1 | VTS-CDU-300 | Coolant Distribution Unit, 300 kW, N+1 pump | 2 | 48500.00 | 97000.00 |" in out
    assert "| 검산 | 원본 | 결과 | 일치 |" in out
    assert _snapshot(work) == before  # 작업 폴더·장부 변화 없음
    assert (SAMPLES / "INV-2026-10231.pdf").exists()


def test_check_unknown_shows_reason_and_text_head(setup, capsys):
    repo, _ = setup
    _check(str(SAMPLES / "Letter_warranty_notice.pdf"), repo)
    out = capsys.readouterr().out
    assert "확인필요" in out and "판별 실패" in out
    assert "Vertex" in out  # 1쪽 텍스트 앞부분


def test_check_missing_field_shows_note(setup, capsys):
    repo, work = setup
    rules = work / "규칙.yaml"  # 작업 폴더의 규칙
    rules.write_text(rules.read_text(encoding="utf-8").replace('"Bill To:', '"Sold To:'), encoding="utf-8")
    _check(str(SAMPLES / "INV-2026-10231.pdf"), repo)
    out = capsys.readouterr().out
    assert "| 고객명 | (비어 있음) | 못 찾음 |" in out


def test_check_bad_rules_prints_korean_error(setup, capsys):
    repo, work = setup
    (work / "규칙.yaml").write_text("문서종류: 없음", encoding="utf-8")
    code = _check(str(SAMPLES / "INV-2026-10231.pdf"), repo)
    assert code == 1
    assert "규칙" in capsys.readouterr().out


def test_without_init_says_run_setup(tmp_path, capsys):
    code = main(["check", str(SAMPLES / "INV-2026-10231.pdf")], repo_root=tmp_path)
    assert code == 1
    assert "설치.bat를 먼저 실행하세요" in capsys.readouterr().out


# --- 엑셀 PO 도 check 로 본다 ---

ORDERS = REPO / "samples" / "01_고객오더_엑셀"


def test_check_excel_shows_head_items_master_and_form(setup, capsys):
    repo, work = setup
    before = _snapshot(work)
    code = _check(str(ORDERS / "고객오더_ACME_PO-ACME-2026-0417.xlsx"), repo)
    out = capsys.readouterr().out
    assert code == 0
    assert "엑셀양식: ACME" in out
    assert "| 고객PO번호 | PO-ACME-2026-0417 |" in out
    assert "| 발주일 | 2026-09-18 |" in out
    assert "품목 4줄" in out
    assert "| 1 | VTS-CDU-300 |" in out
    assert "고객마스터: 찾음 (고객코드 C-KR-0417)" in out
    assert "| 마크업 검산 | 계산 | 결과 | 일치 |" in out  # 예시 마크업표(init 이 만든 것)로 줄마다 본사 단가
    assert "| 1줄 마크업% (VTS-CDU-300) | 마크업표 3% | 3 | 일치 |" in out
    assert "| 1줄 본사 단가 | 48500 x (100 + 3) / 100 = 49955 | 49955.00 | 일치 |" in out
    assert "회사 양식: 회사양식_SO입력_PO-ACME-2026-0417.xlsx (4줄)" in out
    assert _snapshot(work) == before  # 아무것도 만들거나 옮기지 않음


def test_check_excel_unknown_customer_says_so(setup, capsys):
    repo, work = setup
    import openpyxl
    master = work.parent.parent / "고객마스터.xlsx"
    wb = openpyxl.load_workbook(master)
    wb.active.delete_rows(2, 2)
    wb.save(master)
    _check(str(ORDERS / "고객오더_HANIL_HDC-PUR-26-0908.xlsx"), repo)
    assert "고객마스터: 없음" in capsys.readouterr().out


def test_check_old_xls_explains(setup, capsys, tmp_path):
    repo, _ = setup
    old = tmp_path / "옛주문.xls"
    old.write_bytes(b"old")
    _check(str(old), repo)
    assert ".xlsx" in capsys.readouterr().out