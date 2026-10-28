"""규칙.yaml 의 엑셀양식(고객별 엑셀 읽는 법)·회사양식(회사 SO 양식) 로드·검증."""
from pathlib import Path

import pytest
import yaml

from app.rules import RuleError, load_rules

REPO = Path(__file__).resolve().parent.parent
EXAMPLE = REPO / "rules" / "규칙.예시.yaml"


def _broken(tmp_path, change):
    data = yaml.safe_load(EXAMPLE.read_text(encoding="utf-8"))
    change(data)
    path = tmp_path / "규칙.yaml"
    path.write_text(yaml.safe_dump(data, allow_unicode=True), encoding="utf-8")
    return path


def test_example_excel_forms():
    rules = load_rules(EXAMPLE)
    assert [f.name for f in rules.excel_forms] == ["ACME", "HANIL"]
    acme = rules.excel_forms[0]
    assert acme.doc_type == "PO"
    assert acme.keywords == ["Buyer:"]
    assert acme.head["고객PO번호"].where == "오른쪽:PO Number:"
    assert acme.head["발주일"].fmt == "날짜"
    assert acme.head["합계금액"].fmt == "금액"
    assert acme.items.header_text == "Supplier Part No"
    assert acme.items.columns["수량"] == "D"
    assert acme.items.amount_columns == ["단가", "금액"]
    assert rules.excel_forms[1].head["고객명"].where == "A1"


def test_example_company_form():
    form = load_rules(EXAMPLE).company_form
    assert form.sheet == "SO_Upload"
    assert form.filename_tpl == "회사양식_SO입력_{고객PO번호}.xlsx"
    names = [c.name for c in form.columns]
    assert len(names) == 19 and names[0] == "Order Date" and names[-1] == "Remarks"
    sources = {c.name: c.source for c in form.columns}
    assert sources["Customer Code"] == "고객마스터.고객코드"
    assert sources["Line No"] == "줄번호"
    assert sources["Ext Price"] == "수량x단가"
    assert sources["Remarks"] == "빈칸"


def test_rules_without_excel_sections_still_load(tmp_path):
    def drop(d):
        d.pop("엑셀양식")
        d.pop("회사양식")
    rules = load_rules(_broken(tmp_path, drop))
    assert rules.excel_forms == [] and rules.company_form is None


@pytest.mark.parametrize("cell", ["4B", "오른쪽:", "B0", "왼쪽:Total"])
def test_bad_cell_address_raises(tmp_path, cell):
    path = _broken(tmp_path, lambda d: d["엑셀양식"]["ACME"]["머리"].update(고객명=cell))
    with pytest.raises(RuleError, match="ACME"):
        load_rules(path)


def test_unknown_doc_type_raises(tmp_path):
    path = _broken(tmp_path, lambda d: d["엑셀양식"]["ACME"].update(문서종류="견적서"))
    with pytest.raises(RuleError, match="견적서"):
        load_rules(path)


def test_number_field_must_be_in_head(tmp_path):
    path = _broken(tmp_path, lambda d: d["엑셀양식"]["ACME"]["머리"].pop("고객PO번호"))
    with pytest.raises(RuleError, match="고객PO번호"):
        load_rules(path)


def test_item_column_letter_must_be_valid(tmp_path):
    path = _broken(tmp_path, lambda d: d["엑셀양식"]["ACME"]["품목"]["열"].update(수량="4"))
    with pytest.raises(RuleError, match="수량"):
        load_rules(path)


@pytest.mark.parametrize("value", ["머리", "엉뚱한.값", "수량+단가", ""])
def test_bad_company_value_raises(tmp_path, value):
    path = _broken(tmp_path, lambda d: d["회사양식"]["열"][0].update(값=value))
    with pytest.raises(RuleError, match="Order Date"):
        load_rules(path)
