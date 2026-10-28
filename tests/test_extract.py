"""텍스트 추출·문서종류 판별·필드·표 추출."""
from datetime import date
from decimal import Decimal
from pathlib import Path

import pytest

from app.extract import detect_type, extract, extract_from_text, read_text, to_date
from app.rules import load_rules

REPO = Path(__file__).resolve().parent.parent
SAMPLES = REPO / "samples"
INVOICE = SAMPLES / "ACME_Invoice_INV-2026-0001.pdf"
PO = SAMPLES / "ACME_PO_PO-77001.pdf"
DM = SAMPLES / "ACME_DM_DM-5101.pdf"


@pytest.fixture(scope="module")
def rules():
    return load_rules(REPO / "rules" / "규칙.예시.yaml")


# --- 1부: 텍스트·판별 ---

def test_read_text_contains_invoice_no():
    assert "Invoice No" in read_text(INVOICE)


@pytest.mark.parametrize("pdf, expected", [(INVOICE, "Invoice"), (PO, "PO"), (DM, "DM")])
def test_detect_type(rules, pdf, expected):
    assert detect_type(read_text(pdf), rules).name == expected


def test_detect_type_is_case_insensitive(rules):
    assert detect_type("debit memo no: x", rules).name == "DM"


def test_detect_unknown_returns_none(rules):
    assert detect_type(read_text(SAMPLES / "Unknown_letter.pdf"), rules) is None


def test_scan_blank_has_no_text():
    assert len(read_text(SAMPLES / "Scan_blank.pdf")) == 0


# --- 2부: 필드·표 ---

INVOICE_TEXT = """INVOICE
Invoice No: INV-9
Date: 10/15/2026
Bill To: Foo Corp
Item Description Qty Unit Price Amount
A-1 Big Pump 2 1.00 2.00
Subtotal: 2.00
Total: $2.00"""


def test_invoice_fields(rules):
    r = extract(INVOICE, rules)
    assert r.doc_type.name == "Invoice"
    assert r.fields == {
        "인보이스번호": "INV-2026-0001",
        "발행일": date(2026, 10, 15),
        "고객명": "ACME Semiconductor Korea",
        "합계금액": Decimal("12345.67"),
    }
    assert r.missing_required is False
    assert r.warnings == []


def test_date_is_date_object(rules):
    assert type(extract(PO, rules).fields["발주일"]) is date


def test_amount_is_decimal(rules):
    assert extract(DM, rules).fields["합계금액"] == Decimal("12345.67")


def test_table_rows(rules):
    rows = extract(INVOICE, rules).table
    assert len(rows) == 3
    assert rows[0] == {"품번": "CP-200", "품명": "Chiller Pump Assembly", "수량": "2", "단가": "4,500.00", "금액": "9,000.00"}
    assert rows[2]["품번"] == "VL-055"
    assert len(extract(PO, rules).table) == 2
    assert len(extract(DM, rules).table) == 2


def test_to_date_formats():
    assert to_date("10/15/2026") == date(2026, 10, 15)
    assert to_date("2026-10-15") == date(2026, 10, 15)
    assert to_date("15 Oct") is None


def test_number_field_missing_is_required_failure(rules):
    r = extract_from_text(INVOICE_TEXT.replace("Invoice No: INV-9", "Invoice"), rules)
    assert r.doc_type.name == "Invoice"
    assert r.missing_required is True
    assert "인보이스번호" in r.review_reason


def test_optional_field_missing_is_none_with_warning(rules):
    r = extract_from_text(INVOICE_TEXT.replace("Bill To: Foo Corp\n", ""), rules)
    assert r.fields["고객명"] is None
    assert r.missing_required is False
    assert any("고객명" in w for w in r.warnings)


def test_table_column_mismatch_skips_table(rules):
    r = extract_from_text(INVOICE_TEXT.replace("A-1 Big Pump 2 1.00 2.00", "A-1 2.00"), rules)
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
    r = extract(SAMPLES / "Scan_blank.pdf", rules)
    assert r.review_reason == "텍스트 없음 — 스캔본, OCR 범위 밖"


def test_extra_words_without_multiword_column_is_mismatch(rules):
    import copy
    strict = copy.deepcopy(rules)
    strict.get("Invoice").table.multiword_column = None
    r = extract_from_text(INVOICE_TEXT, strict)
    assert r.table is None
