"""샘플 세트(samples/)와 테스트 전용 생성기(tests/fixtures/build_fixtures.py) 검사."""
import importlib.util
from pathlib import Path

import pdfplumber

REPO = Path(__file__).resolve().parent.parent
SAMPLE_PDFS = REPO / "samples" / "03_Invoice_DM_PDF"
SAMPLE_SET = [
    "03_Invoice_DM_PDF/INV-2026-10231.pdf",
    "03_Invoice_DM_PDF/INV-2026-10278.pdf",
    "03_Invoice_DM_PDF/DM-2026-0092.pdf",
    "03_Invoice_DM_PDF/DM-2026-0097.pdf",
    "03_Invoice_DM_PDF/Letter_warranty_notice.pdf",
    "03_Invoice_DM_PDF/Scan_signed_packing_list.pdf",
    "04_장부_엑셀/누적장부_기대결과.xlsx",
    "README_샘플설명.md",
    "make_samples.py",
]


def _load_builder():
    spec = importlib.util.spec_from_file_location("build_fixtures", REPO / "tests" / "fixtures" / "build_fixtures.py")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def test_sample_set_is_committed():
    for name in SAMPLE_SET:
        assert (REPO / "samples" / name).is_file(), name
    assert len(list(SAMPLE_PDFS.glob("*.pdf"))) == 6


def test_fixture_builder_writes_only_into_given_folder(tmp_path):
    files = _load_builder().build(tmp_path)
    assert len(files) == 5
    assert all(f.parent == tmp_path for f in files)
    with pdfplumber.open(tmp_path / "ACME_Invoice_INV-2026-0001.pdf") as pdf:
        assert "Invoice No: INV-2026-0001" in pdf.pages[0].extract_text()
