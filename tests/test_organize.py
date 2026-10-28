"""파일 이동·이름 충돌·확인필요 .txt·중복 폴더."""
from datetime import date
from decimal import Decimal
from pathlib import Path

from app.extract import ExtractResult
from app.organize import move_duplicate, move_review, move_success, render_filename, target_path
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


def test_target_path_is_under_organized(tmp_path):
    assert target_path(_result(), tmp_path) == tmp_path / "정리문서" / "ACME Korea" / "2026-10" / "Invoice_INV-1.pdf"


def test_move_success_moves_file(tmp_path):
    src = _pdf(tmp_path / "받은문서")
    dst = tmp_path / "정리문서" / "x" / "b.pdf"
    final = move_success(src, dst)
    assert final == dst
    assert not src.exists()
    assert dst.read_bytes() == b"%PDF-1.4 test"


def test_move_success_name_collision_adds_suffix(tmp_path):
    dst = tmp_path / "정리문서" / "b.pdf"
    _pdf(dst.parent, "b.pdf", b"old")
    first = move_success(_pdf(tmp_path / "in", "1.pdf"), dst)
    second = move_success(_pdf(tmp_path / "in", "2.pdf"), dst)
    assert first.name == "b_2.pdf"
    assert second.name == "b_3.pdf"
    assert dst.read_bytes() == b"old"  # 기존 파일은 덮어쓰지 않음


def test_move_review_writes_txt(tmp_path):
    src = _pdf(tmp_path / "받은문서", "letter.pdf")
    final = move_review(src, tmp_path, "판별 실패 — 판별키워드가 하나도 없습니다", "첫 줄\n둘째 줄")
    assert final == tmp_path / "확인필요" / "letter.pdf"
    assert final.exists() and not src.exists()
    note = (tmp_path / "확인필요" / "letter.txt").read_text(encoding="utf-8")
    assert "판별 실패" in note
    assert "첫 줄\n둘째 줄" in note


def test_move_review_collision_keeps_txt_pair(tmp_path):
    _pdf(tmp_path / "확인필요", "letter.pdf", b"old")
    final = move_review(_pdf(tmp_path / "받은문서", "letter.pdf"), tmp_path, "사유", "")
    assert final.name == "letter_2.pdf"
    assert (tmp_path / "확인필요" / "letter_2.txt").exists()


def test_move_duplicate(tmp_path):
    src = _pdf(tmp_path / "받은문서", "dup.pdf")
    final = move_duplicate(src, tmp_path)
    assert final == tmp_path / "받은문서" / "중복" / "dup.pdf"
    assert final.exists() and not src.exists()
