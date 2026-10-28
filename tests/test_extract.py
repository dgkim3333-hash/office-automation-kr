"""텍스트 추출·문서종류 판별·필드·표 추출 — samples/03_Invoice_DM_PDF 기준."""
from datetime import date
from decimal import Decimal
from pathlib import Path

import pytest

from app.extract import detect_type, extract, extract_from_text, read_text, to_date
from app.rules import load_rules

REPO = Path(__file__).resolve().parent.parent
SAMPLES = REPO / "samples" / "03_Invoice_DM_PDF"
INV_ACME = SAMPLES / "INV-2026-10231.pdf"
INV_HANIL = SAMPLES / "INV-2026-10278.pdf"
DM_ACME = SAMPLES / "DM-2026-0092.pdf"
DM_HANIL = SAMPLES / "DM-2026-0097.pdf"
LETTER = SAMPLES / "Letter_warranty_notice.pdf"
SCAN = SAMPLES / "Scan_signed_packing_list.pdf"


@pytest.fixture(scope="module")
def rules():
    return load_rules(REPO / "rules" / "규칙.예시.yaml")


# --- 1부: 텍스트·판별 ---

def test_read_text_contains_invoice_no():
    assert "Invoice No" in read_text(INV_ACME)


@pytest.mark.parametrize("pdf, expected", [
    (INV_ACME, "Invoice"), (INV_HANIL, "Invoice"), (DM_ACME, "DM"), (DM_HANIL, "DM"),
])
def test_detect_type(rules, pdf, expected):
    assert detect_type(read_text(pdf), rules).name == expected


def test_dm_with_reference_invoice_is_not_invoice(rules):
    text = read_text(DM_ACME)
    assert "Reference Invoice" in text
    assert detect_type(text, rules).name == "DM"


def test_detect_type_is_case_insensitive(rules):
    assert detect_type("debit memo no: x", rules).name == "DM"


def test_letter_returns_none(rules):
    assert detect_type(read_text(LETTER), rules) is None


def test_scan_has_no_text():
    assert len(read_text(SCAN)) == 0


# --- 2부: 필드·표 ---

def test_invoice_fields(rules):
    r = extract(INV_ACME, rules)
    assert r.doc_type.name == "Invoice"
    assert r.fields == {
        "인보이스번호": "INV-2026-10231",
        "발행일": date(2026, 10, 14),
        "고객명": "ACME Semiconductor Korea Co., Ltd.",
        "고객PO번호": "PO-ACME-2026-0417",
        "판매오더번호": "SO-26-11892",
        "결제조건": "Net 60",
        "인도조건": "DAP Cheongun",
        "선적일": date(2026, 10, 13),
        "운송사": "DHL Global Forwarding",
        "추적번호": "DHLGF-7731-884210",
        "도착예정일": date(2026, 10, 17),
        "소계": Decimal("111590.00"),
        "운임": Decimal("3850.00"),
        "합계금액": Decimal("115440.00"),
    }
    assert r.missing_required is False
    assert r.warnings == []


def test_dm_fields(rules):
    r = extract(DM_HANIL, rules)
    assert r.fields["DM번호"] == "DM-2026-0097"
    assert r.fields["참조인보이스"] == "INV-2026-10278"
    assert r.fields["합계금액"] == Decimal("300.00")
    assert r.warnings == []


def test_date_is_date_object(rules):
    assert type(extract(INV_HANIL, rules).fields["선적일"]) is date


def test_table_rows_with_row_pattern_and_types(rules):
    rows = extract(INV_ACME, rules).table
    assert len(rows) == 4
    assert rows[0] == {
        "줄": "1", "품번": "VTS-CDU-300", "품명": "Coolant Distribution Unit, 300 kW, N+1 pump",
        "수량": 2, "단가": Decimal("48500.00"), "금액": Decimal("97000.00"),
    }
    assert len(extract(INV_HANIL, rules).table) == 3
    assert extract(DM_ACME, rules).table[0]["품명"] == "Expedited air freight surcharge"
    assert extract(DM_HANIL, rules).table[0]["품명"] == "Price difference 20 x (USD 155.00 - 140.00)"


INVOICE_TEXT = """INVOICE
Invoice No: INV-9
Date: 10/15/2026
Bill To: Ship To:
Foo Co., Ltd. Bar Dock
Line Item Description Qty Unit Price Amount
1 A-1 Big Pump, 2 kW 2 1.00 2.00
Subtotal: $2.00
Total: $2.00"""


def test_to_date_formats():
    assert to_date("10/15/2026") == date(2026, 10, 15)
    assert to_date("2026-10-15") == date(2026, 10, 15)
    assert to_date("15 Oct") is None


def test_number_field_missing_is_required_failure(rules):
    r = extract_from_text(INVOICE_TEXT.replace("Invoice No: INV-9", "Invoice No: (pending)"), rules)
    assert r.doc_type.name == "Invoice"
    assert r.missing_required is True
    assert "인보이스번호" in r.review_reason


def test_optional_field_missing_is_none_with_warning(rules):
    r = extract_from_text(INVOICE_TEXT, rules)
    assert r.fields["운송사"] is None
    assert r.missing_required is False
    assert any("운송사" in w for w in r.warnings)


def test_crafted_table_row(rules):
    rows = extract_from_text(INVOICE_TEXT, rules).table
    assert rows == [{"줄": "1", "품번": "A-1", "품명": "Big Pump, 2 kW", "수량": 2,
                     "단가": Decimal("1.00"), "금액": Decimal("2.00")}]


def test_row_not_matching_pattern_skips_table(rules):
    r = extract_from_text(INVOICE_TEXT.replace("1 A-1 Big Pump, 2 kW 2 1.00 2.00", "1 A-1 2.00"), rules)
    assert r.table is None
    assert any("표" in w for w in r.warnings)
    assert r.fields["인보이스번호"] == "INV-9"


def test_text_head_is_first_30_lines(rules):
    text = "\n".join(f"line {i}" for i in range(50))
    r = extract_from_text(text, rules)
    assert r.text_head.splitlines() == [f"line {i}" for i in range(30)]
    assert r.doc_type is None
    assert r.review_reason.startswith("판별 실패")


def test_blank_scan_review_reason(rules):
    assert extract(SCAN, rules).review_reason.startswith("텍스트 없음")
