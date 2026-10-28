"""파일 이름: 템플릿 채우기 · 이름 충돌."""
from datetime import date
from decimal import Decimal
from pathlib import Path

from app.extract import ExtractResult
from app.organize import free_name, render_filename
from app.rules import load_rules

REPO = Path(__file__).resolve().parent.parent
RULES = load_rules(REPO / "rules" / "규칙.예시.yaml")


def _result(**fields):
    base = {"인보이스번호": "INV-1", "발행일": date(2026, 10, 15), "고객명": "ACME Korea", "합계금액": Decimal("1.00")}
    base.update(fields)
    return ExtractResult(doc_type=RULES.get("Invoice"), fields=base)


def _pdf(folder, name="a.pdf", content=b"%PDF-1.4 test"):
    folder.mkdir(parents=True, exist_ok=True)
    path = folder / name
    path.write_bytes(content)
    return path


def test_render_filename_with_date_format():
    assert render_filename(_result()) == Path("ACME Korea/2026-10/Invoice_INV-1.pdf")


def test_render_filename_replaces_forbidden_characters():
    path = render_filename(_result(고객명='A/B:C*D?"E<F>G|H\\I'))
    assert path.parts[0] == "A_B_C_D__E_F_G_H_I"
    assert len(path.parts) == 3


def test_render_filename_missing_value_is_misang():
    assert render_filename(_result(고객명=None, 발행일=None)) == Path("미상/미상/Invoice_INV-1.pdf")


def test_free_name_adds_suffix_and_never_overwrites(tmp_path):
    taken = _pdf(tmp_path, "b.pdf", b"old")
    assert free_name(tmp_path / "new.pdf") == tmp_path / "new.pdf"
    assert free_name(taken) == tmp_path / "b_2.pdf"
    assert free_name(taken, reserved={tmp_path / "b_2.pdf"}) == tmp_path / "b_3.pdf"
    assert taken.read_bytes() == b"old"
