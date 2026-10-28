"""텍스트 추출·문서종류 판별·필드·표 추출."""
from pathlib import Path

import pytest

from app.extract import detect_type, read_text
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
