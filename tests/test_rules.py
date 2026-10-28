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
    assert [d.name for d in rules.doc_types] == ["Invoice", "PO", "DM"]
    invoice = rules.doc_types[0]
    assert invoice.keywords == ["INVOICE", "Invoice No"]
    assert invoice.number_field == "인보이스번호"
    assert list(invoice.fields) == ["인보이스번호", "발행일", "고객명", "합계금액"]
    assert invoice.fields["발행일"].fmt == "날짜"
    assert invoice.fields["합계금액"].fmt == "금액"
    assert invoice.fields["고객명"].fmt is None
    assert invoice.fields["고객명"].pattern.search("Bill To: X").group(1) == "X"
    assert invoice.table.columns == ["품번", "품명", "수량", "단가", "금액"]
    assert invoice.table.multiword_column == "품명"


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
