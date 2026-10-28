"""가짜 샘플 생성기 검사."""
import importlib.util
from pathlib import Path

import pdfplumber

REPO = Path(__file__).resolve().parent.parent
SAMPLE_NAMES = [
    "ACME_Invoice_INV-2026-0001.pdf",
    "ACME_PO_PO-77001.pdf",
    "ACME_DM_DM-5101.pdf",
    "Unknown_letter.pdf",
    "Scan_blank.pdf",
]


def _load_make_samples():
    spec = importlib.util.spec_from_file_location("make_samples", REPO / "samples" / "make_samples.py")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def _text(path):
    with pdfplumber.open(path) as pdf:
        return "\n".join(page.extract_text() or "" for page in pdf.pages)


def test_build_creates_five_pdfs(tmp_path):
    _load_make_samples().build(tmp_path)
    assert sorted(p.name for p in tmp_path.glob("*.pdf")) == sorted(SAMPLE_NAMES)


def test_generated_text_is_extractable(tmp_path):
    _load_make_samples().build(tmp_path)
    assert "Invoice No: INV-2026-0001" in _text(tmp_path / "ACME_Invoice_INV-2026-0001.pdf")
    assert "PO Number: PO-77001" in _text(tmp_path / "ACME_PO_PO-77001.pdf")
    assert "Debit Memo No: DM-5101" in _text(tmp_path / "ACME_DM_DM-5101.pdf")
    assert _text(tmp_path / "Scan_blank.pdf").strip() == ""


def test_committed_samples_exist():
    for name in SAMPLE_NAMES:
        assert (REPO / "samples" / name).is_file(), name
