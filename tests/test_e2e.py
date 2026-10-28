"""03장 '첫 실행 10분' 과 같은 시나리오: samples/03_Invoice_DM_PDF 6종 → run → 폴더·장부·로그 검증.

기대 장부는 samples/04_장부_엑셀/누적장부_기대결과.xlsx (PO 는 2차 범위라 제외).
"""
import csv
import shutil
from decimal import Decimal, InvalidOperation
from pathlib import Path

import openpyxl
import pytest

from app.__main__ import main
from app.init import init

REPO = Path(__file__).resolve().parent.parent
SAMPLES = REPO / "samples" / "03_Invoice_DM_PDF"
EXPECTED = REPO / "samples" / "04_장부_엑셀" / "누적장부_기대결과.xlsx"
ACME_DIR = Path("정리문서") / "ACME Semiconductor Korea Co., Ltd" / "2026-10"
HANIL_DIR = Path("정리문서") / "Hanil DataCenter Co., Ltd" / "2026-10"
SKIP_COLUMNS = {"처리일시", "정리된경로"}  # 실행 시각·폴더 이름은 기대 장부와 비교하지 않음


@pytest.fixture
def env(tmp_path):
    repo = tmp_path / "repo"
    (repo / "rules").mkdir(parents=True)
    shutil.copy(REPO / "rules" / "규칙.예시.yaml", repo / "rules" / "규칙.예시.yaml")
    work = init(work_dir=tmp_path / "work", repo_root=repo)
    return repo, work


def _drop_samples(work):
    for pdf in SAMPLES.glob("*.pdf"):
        shutil.copy(pdf, work / "받은문서" / pdf.name)


def _records(ws):
    """시트를 [{열이름: 값}] 으로."""
    rows = list(ws.iter_rows(values_only=True))
    header = [h for h in rows[0] if h is not None]
    return [dict(zip(header, r)) for r in rows[1:] if any(v is not None for v in r)]


def _same(a, b):
    """숫자는 값으로(97000 == Decimal('97000.00') == '97000'), 날짜는 날짜로, 나머지는 글자로 비교."""
    if hasattr(a, "date") and hasattr(b, "date"):
        return a.date() == b.date()
    try:
        return Decimal(str(a)) == Decimal(str(b))
    except (InvalidOperation, ValueError):
        return a == b


def _csv_rows(work):
    with open(work / "처리로그.csv", encoding="utf-8-sig", newline="") as f:
        return list(csv.reader(f))[1:]


def _pdfs(folder):
    return sorted(p.name for p in folder.glob("*.pdf"))


def test_first_run(env, capsys):
    repo, work = env
    _drop_samples(work)
    assert main([], repo_root=repo) == 0
    out = capsys.readouterr().out
    assert "성공 4 · 중복 0 · 확인필요 2" in out

    assert _pdfs(work / ACME_DIR) == ["DM_DM-2026-0092.pdf", "Invoice_INV-2026-10231.pdf"]
    assert _pdfs(work / HANIL_DIR) == ["DM_DM-2026-0097.pdf", "Invoice_INV-2026-10278.pdf"]
    assert _pdfs(work / "확인필요") == ["Letter_warranty_notice.pdf", "Scan_signed_packing_list.pdf"]
    assert sorted(p.name for p in (work / "확인필요").glob("*.txt")) == [
        "Letter_warranty_notice.txt", "Scan_signed_packing_list.txt"]
    assert _pdfs(work / "받은문서") == []
    assert len(_csv_rows(work)) == 6


@pytest.mark.parametrize("sheet", ["Invoice", "DM"])
def test_first_run_matches_expected_ledger(env, sheet):
    repo, work = env
    _drop_samples(work)
    main([], repo_root=repo)
    got = {r["원본파일"]: r for r in _records(openpyxl.load_workbook(work / "누적장부.xlsx")[sheet])}
    expected = {r["원본파일"]: r for r in _records(openpyxl.load_workbook(EXPECTED)[sheet])}
    assert sorted(got) == sorted(expected)
    for name, exp in expected.items():
        for column, value in exp.items():
            if column in SKIP_COLUMNS:
                continue
            assert _same(got[name][column], value), f"{sheet} {name} {column}: {got[name][column]!r} != {value!r}"
        assert Path(got[name]["정리된경로"]).name == Path(exp["정리된경로"]).name


def test_first_run_matches_expected_items_and_log(env):
    repo, work = env
    _drop_samples(work)
    main([], repo_root=repo)
    wb = openpyxl.load_workbook(work / "누적장부.xlsx")
    exp_wb = openpyxl.load_workbook(EXPECTED)

    key = lambda r: (r["문서종류"], r["문서번호"], str(r["줄"]))  # noqa: E731
    got_items = sorted(_records(wb["품목"]), key=key)
    exp_items = sorted((r for r in _records(exp_wb["품목"]) if r["문서종류"] != "PO"), key=key)
    assert len(got_items) == len(exp_items) == 9
    for g, e in zip(got_items, exp_items):
        for column, value in e.items():
            assert _same(g[column], value), f"품목 {key(e)} {column}: {g[column]!r} != {value!r}"

    got_log = {r["파일"]: r["결과"] for r in _records(wb["처리로그"])}
    exp_log = {r["파일"]: r["결과"] for r in _records(exp_wb["처리로그"]) if r["파일"].endswith(".pdf")}
    assert got_log == exp_log


def test_second_run_marks_duplicates(env, capsys):
    repo, work = env
    _drop_samples(work)
    main([], repo_root=repo)
    _drop_samples(work)
    main([], repo_root=repo)
    out = capsys.readouterr().out
    assert "성공 0 · 중복 4 · 확인필요 2" in out
    assert len(_pdfs(work / "받은문서" / "중복")) == 4
    assert _pdfs(work / "받은문서") == []
    wb = openpyxl.load_workbook(work / "누적장부.xlsx")
    assert len(_records(wb["Invoice"])) == 2
    assert len(_records(wb["품목"])) == 9
    assert len(_csv_rows(work)) == 12


def test_overwrite_keeps_row_count_and_updates_time(env, capsys, monkeypatch):
    repo, work = env
    _drop_samples(work)
    main([], repo_root=repo)
    first = _records(openpyxl.load_workbook(work / "누적장부.xlsx")["Invoice"])[0]["처리일시"]

    shutil.copy(SAMPLES / "INV-2026-10231.pdf", work / "받은문서")
    import app.__main__ as cli
    monkeypatch.setattr(cli, "_now", lambda: first.replace(year=first.year + 1))
    main(["--덮어쓰기"], repo_root=repo)
    out = capsys.readouterr().out
    assert "성공 1 · 중복 0 · 확인필요 0" in out

    wb = openpyxl.load_workbook(work / "누적장부.xlsx")
    rows = {r["인보이스번호"]: r for r in _records(wb["Invoice"])}
    assert len(rows) == 2
    assert rows["INV-2026-10231"]["처리일시"] > first
    assert len(_records(wb["품목"])) == 9
    assert "Invoice_INV-2026-10231_2.pdf" in _pdfs(work / ACME_DIR)


def test_locked_ledger_moves_nothing(env, capsys):
    repo, work = env
    _drop_samples(work)
    main([], repo_root=repo)
    capsys.readouterr()

    _drop_samples(work)
    with open(work / "누적장부.xlsx", "rb"):  # 엑셀이 열어 둔 상태
        code = main([], repo_root=repo)
    out = capsys.readouterr().out
    assert code != 0
    assert "엑셀을 닫고 다시 실행" in out
    assert len(_pdfs(work / "받은문서")) == 6
    assert _pdfs(work / "받은문서" / "중복") == []
    assert _pdfs(work / "확인필요") == ["Letter_warranty_notice.pdf", "Scan_signed_packing_list.pdf"]
    assert len(_csv_rows(work)) == 6 + 6  # csv 는 잠김과 무관하게 기록


def test_empty_inbox(env, capsys):
    repo, _ = env
    assert main([], repo_root=repo) == 0
    assert "받은문서에 PDF 가 없습니다" in capsys.readouterr().out
