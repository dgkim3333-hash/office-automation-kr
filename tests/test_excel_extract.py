"""고객 엑셀 PO 읽기 — samples/01_고객오더_엑셀 기준."""
from datetime import date
from decimal import Decimal
from pathlib import Path

import openpyxl
import pytest

from app.excel_extract import read_excel
from app.rules import load_rules

REPO = Path(__file__).resolve().parent.parent
ORDERS = REPO / "samples" / "01_고객오더_엑셀"
ACME = ORDERS / "고객오더_ACME_PO-ACME-2026-0417.xlsx"
HANIL = ORDERS / "고객오더_HANIL_HDC-PUR-26-0908.xlsx"


@pytest.fixture(scope="module")
def rules():
    return load_rules(REPO / "rules" / "규칙.예시.yaml")


def test_acme_head(rules):
    r = read_excel(ACME, rules)
    assert r.doc_type.name == "PO" and r.source_form == "ACME"
    assert r.fields["고객PO번호"] == "PO-ACME-2026-0417"
    assert r.fields["발주일"] == date(2026, 9, 18) and type(r.fields["발주일"]) is date
    assert r.fields["고객명"] == "ACME Semiconductor Korea Co., Ltd."
    assert r.fields["납기요청일"] == date(2026, 11, 20)
    assert r.fields["인도조건"] == "DAP Cheongun"
    assert r.fields["통화"] == "USD"
    assert r.fields["합계금액"] == Decimal("111590")
    assert r.fields["담당"] == "Mr. J. Park (Procurement)"
    assert r.number == "PO-ACME-2026-0417"
    assert r.warnings == [] and r.review_reason is None


def test_acme_items_stop_before_total_row(rules):
    rows = read_excel(ACME, rules).table
    assert len(rows) == 4
    assert rows[0] == {"줄": 1, "품번": "VTS-CDU-300", "품명": "Coolant Distribution Unit, 300 kW, N+1 pump",
                       "수량": 2, "단위": "EA", "단가": Decimal("48500"), "금액": Decimal("97000")}
    assert rows[3]["품번"] == "VTS-SENS-LEAK"


def test_hanil_korean_form(rules):
    r = read_excel(HANIL, rules)
    assert r.source_form == "HANIL"
    assert r.fields["고객명"] == "Hanil DataCenter Co., Ltd."  # A1 셀
    assert r.fields["고객PO번호"] == "HDC-PUR-26-0908"
    assert r.fields["발주일"] == date(2026, 9, 25)
    assert r.fields["합계금액"] == Decimal("224100")
    assert [row["품번"] for row in r.table] == ["VTS-CHL-500RT", "VTS-PMP-CDU-K", "VTS-FLT-50"]


def test_unknown_excel_is_review(rules):
    r = read_excel(REPO / "samples" / "02_회사양식_변환" / "회사양식_SO입력_HDC-PUR-26-0908.xlsx", rules)
    assert r.doc_type is None
    assert r.review_reason.startswith("판별 실패")
    assert "Order Date" in r.text_head  # 무엇이 들어 있었는지 확인필요 메모에 남김


def test_missing_label_is_warning_and_missing_number_is_required(rules, tmp_path):
    wb = openpyxl.load_workbook(ACME)
    ws = wb.active
    ws["A4"] = "Order No:"   # 'PO Number:' 글자가 없어짐 → 번호필드 못 찾음
    ws["A11"] = None         # 'Buyer Contact:' 글자가 없어짐 → 담당 못 찾음
    broken = tmp_path / "broken.xlsx"
    wb.save(broken)
    r = read_excel(broken, rules)
    assert r.fields["고객PO번호"] is None and r.missing_required is True
    assert r.fields["담당"] is None
    assert any("담당" in w for w in r.warnings)


# --- 엑셀 float(수식 결과)은 엑셀처럼 유효 15자리 ---

@pytest.mark.parametrize("value, expected", [
    (0.30000000000000004, "0.3"),  # =3*0.1
    (12.3, "12.3"),
    (97000.0, "97000"),
    (1234.5, "1234.5"),
    (0.1 + 0.7, "0.8"),            # 0.7999999999999999
    (2, "2"),
])
def test_excel_float_is_normalized_to_15_digits(value, expected):
    from app.excel_extract import _as_amount, excel_number
    assert str(excel_number(value)) == expected
    assert str(_as_amount(value)) == expected


def test_formula_float_in_items_becomes_clean_decimal(tmp_path, rules):
    wb = openpyxl.load_workbook(ACME)
    ws = wb.active
    ws["D14"], ws["F14"], ws["G14"] = 3, 0.1, 0.375
    path = tmp_path / "po.xlsx"
    wb.save(path)
    _raw_float(path, "0.375", "0.30000000000000004")
    row = read_excel(path, rules).table[0]
    assert row["단가"] == Decimal("0.1") and str(row["금액"]) == "0.3"


def _raw_float(path, placeholder, raw):
    """openpyxl 은 float 을 짧게 쓰므로, 엑셀이 저장하는 긴 값(=3*0.1 → 0.30000000000000004)을 XML 에 직접 넣는다."""
    import zipfile
    with zipfile.ZipFile(path) as z:
        files = {n: z.read(n) for n in z.namelist()}
    sheet = "xl/worksheets/sheet1.xml"
    assert files[sheet].count(f"<v>{placeholder}</v>".encode()) == 1
    files[sheet] = files[sheet].replace(f"<v>{placeholder}</v>".encode(), f"<v>{raw}</v>".encode())
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        for n, data in files.items():
            z.writestr(n, data)
