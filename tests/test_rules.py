"""규칙.yaml 로드·검증."""
from pathlib import Path

import pytest
import yaml

from app.rules import CheckSpec, RuleError, load_rules

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


def test_example_check_specs():
    rules = load_rules(EXAMPLE)
    assert rules.get("Invoice").check == CheckSpec(subtotal="소계", extras=["운임"], total="합계금액")
    assert rules.get("DM").check == CheckSpec(subtotal=None, extras=[], total="합계금액")
    assert rules.get("PO").check == CheckSpec(subtotal=None, extras=[], total="합계금액")  # 엑셀 PO 도 합계 검산


def test_check_spec_unknown_field_is_rule_error(tmp_path):
    path = _broken(tmp_path, lambda d: d["필드"].pop("운임"))
    with pytest.raises(RuleError, match="검산의 '운임' 은 필드 목록에 없습니다"):
        load_rules(path)


def test_check_spec_field_must_be_amount(tmp_path):
    path = _broken(tmp_path, lambda d: d["필드"]["소계"].pop("형식"))
    with pytest.raises(RuleError, match="소계"):
        load_rules(path)


def test_check_spec_unknown_key_is_rule_error(tmp_path):
    path = _broken(tmp_path, lambda d: d["검산"].update(세금="운임"))
    with pytest.raises(RuleError, match="세금"):
        load_rules(path)


def test_check_spec_extras_need_total(tmp_path):
    path = _broken(tmp_path, lambda d: d["검산"].pop("합계"))
    with pytest.raises(RuleError, match="부대비용"):
        load_rules(path)


@pytest.mark.parametrize("check", [{}, {"소계": None}, {"부대비용": []}, {"합계": ""}])
def test_empty_check_spec_is_rule_error(tmp_path, check):
    path = _broken(tmp_path, lambda d: d.update(검산=check))
    with pytest.raises(RuleError, match="검산에 적은 것이 없습니다"):
        load_rules(path)


def test_example_check_columns_default():
    spec = load_rules(EXAMPLE).get("Invoice").check
    assert (spec.quantity, spec.price, spec.amount) == ("수량", "단가", "금액")


def _renamed_table(d):
    """Invoice 표 열 이름을 영어로 바꾼다."""
    d["표"].update(열=["줄", "품번", "품명", "Qty", "Price", "Amount"], 금액열=["Price", "Amount"], 수량열=["Qty"])


def test_check_columns_can_be_named(tmp_path):
    def change(d):
        _renamed_table(d)
        d["검산"].update(수량열="Qty", 단가열="Price", 금액열="Amount")
    spec = load_rules(_broken(tmp_path, change)).get("Invoice").check
    assert (spec.quantity, spec.price, spec.amount) == ("Qty", "Price", "Amount")
    assert spec.subtotal == "소계" and spec.total == "합계금액"


def test_default_amount_column_missing_with_sum_check_is_rule_error(tmp_path):
    path = _broken(tmp_path, _renamed_table)  # 소계 · 합계 검산이 있으면 기본 '금액' 열이 있어야 한다
    with pytest.raises(RuleError, match="금액열 '금액'"):
        load_rules(path)


def test_total_only_check_without_price_column_loads(tmp_path):
    def change(d):  # 단가 열이 없는 품목 표에 합계 검산만
        d["표"].update(열=["줄", "품번", "품명", "수량", "비고", "금액"], 금액열=["금액"])
        d["검산"] = {"합계": "합계금액"}
    spec = load_rules(_broken(tmp_path, change)).get("Invoice").check
    assert spec.total == "합계금액" and spec.price == "단가"


def test_renamed_quantity_price_must_be_named(tmp_path):
    def change(d):
        _renamed_table(d)
        d["검산"].update(금액열="Amount")  # 수량 · 단가를 안 적으면 줄 검산이 조용히 해당 없음이 되므로 막는다
    with pytest.raises(RuleError, match="검산: 에 수량열 · 단가열 · 금액열 을 적으세요"):
        load_rules(_broken(tmp_path, change))


def test_named_amount_column_checked_even_without_sum(tmp_path):
    path = _broken(tmp_path, lambda d: d.update(검산={"금액열": "Amount"}))
    with pytest.raises(RuleError, match="금액열 'Amount'"):
        load_rules(path)


def test_empty_check_block_is_rule_error(tmp_path):
    path = _broken(tmp_path, lambda d: d.update(검산=None))  # "검산:" 만 쓰고 내용 없음
    with pytest.raises(RuleError, match="검산에 적은 것이 없습니다"):
        load_rules(path)


def test_named_check_column_missing_from_table_is_rule_error(tmp_path):
    path = _broken(tmp_path, lambda d: d["검산"].update(단가열="Price"))
    with pytest.raises(RuleError, match="단가열 'Price'"):
        load_rules(path)


def test_check_column_missing_from_excel_items_is_rule_error(tmp_path):
    data = yaml.safe_load(EXAMPLE.read_text(encoding="utf-8"))
    data["문서종류"]["PO"]["검산"] = {"합계": "합계금액"}
    data["엑셀양식"]["HANIL"]["품목"]["열"] = {"줄": "A", "품번": "B", "품명": "C", "수량": "D", "단가": "F", "Amount": "G"}
    data["엑셀양식"]["HANIL"]["품목"]["금액열"] = ["단가", "Amount"]
    path = tmp_path / "규칙.yaml"
    path.write_text(yaml.safe_dump(data, allow_unicode=True), encoding="utf-8")
    with pytest.raises(RuleError, match="HANIL.*금액열 '금액'"):
        load_rules(path)


def test_named_price_column_missing_from_excel_items_is_rule_error(tmp_path):
    data = yaml.safe_load(EXAMPLE.read_text(encoding="utf-8"))
    data["문서종류"]["PO"]["검산"] = {"합계": "합계금액", "단가열": "Price"}
    path = tmp_path / "규칙.yaml"
    path.write_text(yaml.safe_dump(data, allow_unicode=True), encoding="utf-8")
    with pytest.raises(RuleError, match="ACME.*단가열 'Price'"):
        load_rules(path)


def test_check_spec_single_extra_may_be_text(tmp_path):
    path = _broken(tmp_path, lambda d: d["검산"].update(부대비용="운임"))
    assert load_rules(path).get("Invoice").check.extras == ["운임"]


def test_sheet_order_from_rules():
    rules = load_rules(EXAMPLE)
    assert rules.sheet_order == ["Invoice", "DM", "PO"]
    assert [d.name for d in rules.doc_types] == ["DM", "Invoice", "PO"]  # 판별 순서는 그대로


def test_sheet_order_unknown_type_raises(tmp_path):
    data = yaml.safe_load(EXAMPLE.read_text(encoding="utf-8"))
    data["시트순서"] = ["Invoice", "없는종류"]
    path = tmp_path / "규칙.yaml"
    path.write_text(yaml.safe_dump(data, allow_unicode=True), encoding="utf-8")
    with pytest.raises(RuleError, match="없는종류"):
        load_rules(path)


def test_sheet_order_missing_types_are_appended(tmp_path):
    data = yaml.safe_load(EXAMPLE.read_text(encoding="utf-8"))
    data["시트순서"] = ["PO"]
    path = tmp_path / "규칙.yaml"
    path.write_text(yaml.safe_dump(data, allow_unicode=True), encoding="utf-8")
    assert load_rules(path).sheet_order == ["PO", "DM", "Invoice"]