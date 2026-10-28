"""규칙.yaml 로드·검증."""
from pathlib import Path

import pytest
import yaml

from app.rules import RuleError, load_rules

REPO = Path(__file__).resolve().parent.parent
EXAMPLE = REPO / "rules" / "규칙.예시.yaml"


def _broken(tmp_path, change):
    """예시 규칙을 불러와 change(dict) 로 한 군데 망가뜨린 파일 경로."""
    data = yaml.safe_load(EXAMPLE.read_text(encoding="utf-8"))
    change(data["문서종류"]["Invoice"])
    path = tmp_path / "규칙.yaml"
    path.write_text(yaml.safe_dump(data, allow_unicode=True), encoding="utf-8")
    return path


def test_load_example_rules():
    rules = load_rules(EXAMPLE)
    assert [d.name for d in rules.doc_types] == ["DM", "Invoice", "PO"]  # 검사 순서: DM 먼저
    invoice = rules.get("Invoice")
    assert invoice.keywords == ["Invoice No"]
    assert invoice.number_field == "인보이스번호"
    assert list(invoice.fields)[:3] == ["인보이스번호", "발행일", "고객명"]
    assert invoice.fields["발행일"].fmt == "날짜"
    assert invoice.fields["합계금액"].fmt == "금액"
    assert invoice.fields["고객명"].fmt is None
    table = invoice.table
    assert table.columns == ["줄", "품번", "품명", "수량", "단가", "금액"]
    assert table.amount_columns == ["단가", "금액"]
    assert table.quantity_columns == ["수량"]
    assert table.row_pattern.match("1 P-1 Big Pump, 2 kW 2 1,000.00 2,000.00").groups() == (
        "1", "P-1", "Big Pump, 2 kW", "2", "1,000.00", "2,000.00")
    assert rules.get("PO").table is None


def test_row_pattern_group_count_must_match_columns(tmp_path):
    path = _broken(tmp_path, lambda d: d["표"].update(열=["줄", "품번"]))
    with pytest.raises(RuleError, match="행패턴"):
        load_rules(path)


def test_amount_column_must_be_in_columns(tmp_path):
    path = _broken(tmp_path, lambda d: d["표"].update(금액열=["합계"]))
    with pytest.raises(RuleError, match="합계"):
        load_rules(path)


def test_table_without_row_pattern_raises(tmp_path):
    path = _broken(tmp_path, lambda d: d["표"].pop("행패턴"))
    with pytest.raises(RuleError, match="행패턴"):
        load_rules(path)


def test_old_multiword_key_is_rejected(tmp_path):
    path = _broken(tmp_path, lambda d: d["표"].update(여러단어열="품명"))
    with pytest.raises(RuleError, match="행패턴"):
        load_rules(path)


def test_bad_regex_raises_with_type_and_field(tmp_path):
    path = _broken(tmp_path, lambda d: d["필드"]["고객명"].update(패턴="Bill To ((("))
    with pytest.raises(RuleError) as e:
        load_rules(path)
    assert "Invoice" in str(e.value) and "고객명" in str(e.value)


def test_pattern_without_group_raises(tmp_path):
    path = _broken(tmp_path, lambda d: d["필드"]["고객명"].update(패턴="Bill To"))
    with pytest.raises(RuleError, match="괄호"):
        load_rules(path)


def test_number_field_missing_raises(tmp_path):
    path = _broken(tmp_path, lambda d: d.update(번호필드="없는필드"))
    with pytest.raises(RuleError, match="없는필드"):
        load_rules(path)


def test_filename_variable_missing_raises(tmp_path):
    path = _broken(tmp_path, lambda d: d.update(파일명="{거래처}/Invoice_{인보이스번호}.pdf"))
    with pytest.raises(RuleError, match="거래처"):
        load_rules(path)


def test_unknown_format_raises(tmp_path):
    path = _broken(tmp_path, lambda d: d["필드"]["합계금액"].update(형식="원화"))
    with pytest.raises(RuleError, match="원화"):
        load_rules(path)
