"""문서별 결과 엑셀 — 엑셀 PO → 회사 SO 양식, PDF → 추출결과. 정답은 samples/02·04."""
from decimal import Decimal, InvalidOperation
from pathlib import Path

import openpyxl
import pytest

from app.excel_extract import read_excel
from app.extract import extract
from app.master import load_master, lookup, write_example
from app.outputs import company_rows, write_company_form, write_extract_result
from app.rules import load_rules

REPO = Path(__file__).resolve().parent.parent
SAMPLES = REPO / "samples"
CASES = [
    ("고객오더_ACME_PO-ACME-2026-0417.xlsx", "회사양식_SO입력_PO-ACME-2026-0417.xlsx"),
    ("고객오더_HANIL_HDC-PUR-26-0908.xlsx", "회사양식_SO입력_HDC-PUR-26-0908.xlsx"),
]


@pytest.fixture(scope="module")
def rules():
    return load_rules(REPO / "rules" / "규칙.예시.yaml")


@pytest.fixture(scope="module")
def master(tmp_path_factory):
    path = tmp_path_factory.mktemp("m") / "고객마스터.xlsx"
    write_example(path)
    return load_master(path)


def _same(a, b):
    if hasattr(a, "year") and hasattr(b, "year"):
        return (a.year, a.month, a.day) == (b.year, b.month, b.day)
    try:
        return Decimal(str(a)) == Decimal(str(b))
    except (InvalidOperation, ValueError):
        return a == b


def _sheet(path, name):
    rows = [r for r in openpyxl.load_workbook(path)[name].iter_rows(values_only=True)
            if any(v is not None for v in r)]
    return rows[0], rows[1:]


@pytest.mark.parametrize("order, expected", CASES)
def test_company_rows_match_expected_form(rules, master, order, expected):
    result = read_excel(SAMPLES / "01_고객오더_엑셀" / order, rules)
    header, rows = company_rows(result, lookup(master, result.fields["고객명"]), rules.company_form)
    exp_header, exp_rows = _sheet(SAMPLES / "02_회사양식_변환" / expected, "SO_Upload")
    assert header == list(exp_header)
    assert len(rows) == len(exp_rows)
    for got, exp in zip(rows, exp_rows):
        for name, g, e in zip(header, got, exp):
            assert _same(g, e), f"{order} {name}: {g!r} != {e!r}"


@pytest.mark.parametrize("order, expected", CASES)
def test_written_company_form_file(rules, master, tmp_path, order, expected):
    result = read_excel(SAMPLES / "01_고객오더_엑셀" / order, rules)
    path = write_company_form(tmp_path, result, lookup(master, result.fields["고객명"]), rules.company_form)
    assert path.name == expected
    got_header, got_rows = _sheet(path, "SO_Upload")
    exp_header, exp_rows = _sheet(SAMPLES / "02_회사양식_변환" / expected, "SO_Upload")
    assert got_header == exp_header and len(got_rows) == len(exp_rows)
    assert all(_same(g, e) for gr, er in zip(got_rows, exp_rows) for g, e in zip(gr, er))


def test_customer_missing_from_master_leaves_blank(rules):
    result = read_excel(SAMPLES / "01_고객오더_엑셀" / CASES[0][0], rules)
    header, rows = company_rows(result, None, rules.company_form)
    assert rows[0][header.index("Customer Code")] is None


def test_extract_result_file_matches_sample(rules, tmp_path):
    pdf = SAMPLES / "03_Invoice_DM_PDF" / "INV-2026-10231.pdf"
    path = write_extract_result(tmp_path, extract(pdf, rules), pdf.name)
    assert path.name == "추출결과_INV-2026-10231.xlsx"
    sample = SAMPLES / "04_장부_엑셀" / "추출결과_단건_INV-2026-10231.xlsx"
    for sheet in ("추출결과", "품목"):
        got_header, got_rows = _sheet(path, sheet)
        exp_header, exp_rows = _sheet(sample, sheet)
        assert got_header[:2] == exp_header[:2]
        assert len(got_rows) == len(exp_rows), sheet
        for gr, er in zip(got_rows, exp_rows):
            assert all(_same(g, e) for g, e in zip(gr[:2] if sheet == "추출결과" else gr, er)), (sheet, gr, er)


def test_output_name_collision_adds_suffix(rules, master, tmp_path):
    result = read_excel(SAMPLES / "01_고객오더_엑셀" / CASES[1][0], rules)
    row = lookup(master, result.fields["고객명"])
    first = write_company_form(tmp_path, result, row, rules.company_form)
    second = write_company_form(tmp_path, result, row, rules.company_form)
    assert first.exists() and second.name == first.stem + "_2.xlsx"
