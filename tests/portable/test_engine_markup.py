"""무설치판 엔진 markup: 같은 마크업표 · 같은 품목을 넣으면 지금 판 app\\markup.py 와 같은 결과를 낸다.

마크업표는 openpyxl 로 tmp_path 에 만들고(값 + 칸 서식) tests\\portable\\grid.py 로 격자 JSON 으로 옮겨 엔진에 넣는다.
지금 판 쪽은 markup_from_bytes(파일 바이트)로 같은 파일을 읽는다. 오류 문구는 한 글자도 다르면 안 된다.

엔진 op `markup` 의 입력(action 으로 고른다, grid = 마크업표 첫 시트의 격자):
  read         {grid}                                    -> [{key, part, rate(글자), places}]  (오류면 EngineError)
  value        {grid, items, sources, columns?}          -> {values: [[값 글자 또는 null, ...] 품목마다]}
  missing      {grid, result: {table}, form}             -> {reason, missing, blanks}
  checks       {grid, result: {table}, rows_written, columns?} -> 검산 목록
  fingerprint  {grid, parts}                             -> {fingerprint}
  written      {form, rows}                              -> [{값 출처: 값}]
"""
import json
from collections import namedtuple
from datetime import datetime
from decimal import Decimal
from types import SimpleNamespace

import openpyxl
import pytest

from app import markup as M
from app.portable import _company_form
from app.rules import CompanyColumn, CompanyForm, MarkupColumns

from .conftest import PORTABLE, EngineError, decode_output, run_powershell
from .grid import grid_from_xlsx

HEADER = ("파트넘버", "마크업%", "반올림자리")
F = namedtuple("F", "value fmt")  # 칸 값 + 칸 서식
SOURCES = [M.RATE, M.PRICE, M.AMOUNT]


# --- 도우미 ---

def markup_file(tmp_path, rows, header=HEADER, name="마크업표.xlsx", extra_first_sheet=None):
    """rows 의 각 줄은 칸 값 목록(None = 빈 칸, F(값, 서식) = 서식 있는 칸). 줄 자체가 None 이면 빈 줄."""
    wb = openpyxl.Workbook()
    ws = wb.active
    if extra_first_sheet is not None:  # 첫 시트는 이것, 마크업표는 둘째 시트(활성)
        ws.append(extra_first_sheet)
        ws = wb.create_sheet("마크업표")
        wb.active = 1
    for c, h in enumerate(header, 1):
        if h is not None:
            ws.cell(row=1, column=c, value=h)
    for r, row in enumerate(rows, 2):
        for c, spec in enumerate(row or (), 1):
            if spec is None:
                continue
            value, fmt = (spec.value, spec.fmt) if isinstance(spec, F) else (spec, None)
            cell = ws.cell(row=r, column=c, value=value)
            if fmt:
                cell.number_format = fmt
    path = tmp_path / name
    wb.save(path)
    return path


def python_read(path):
    try:
        table = M.markup_from_bytes(path.read_bytes())
    except M.MarkupError as e:
        return "error", str(e)
    return "ok", [{"key": k, "part": r.part, "rate": format(r.rate, "f"), "places": r.places} for k, r in table.items()]


def engine_read(run_engine, path, culture=None):
    try:
        out = run_engine("markup", {"action": "read", "grid": grid_from_xlsx(path)}, culture)
    except EngineError as e:
        return "error", str(e)
    return "ok", out


def same_read(run_engine, path, culture=None):
    expected, got = python_read(path), engine_read(run_engine, path, culture)
    assert got == expected
    return expected


def jsonable(value):
    """파이썬 값을 엔진 JSON 표기로: 금액(Decimal)은 자리 그대로의 글자, 나머지는 그대로."""
    if isinstance(value, Decimal):
        return format(value, "f")
    if isinstance(value, dict):
        return {k: jsonable(v) for k, v in value.items()}
    if isinstance(value, list):
        return [jsonable(v) for v in value]
    return value


def table_of(path):
    return M.markup_from_bytes(path.read_bytes())


def python_values(path, items, sources, columns=None):
    table = table_of(path)
    cols = MarkupColumns(**columns) if columns else MarkupColumns()
    out = []
    for item in items:
        row = []
        for source in sources:
            value = M.markup_value(source, item, cols, table)
            row.append(None if value is None else format(value, "f"))
        out.append(row)
    return out


def engine_values(run_engine, path, items, sources, columns=None, culture=None):
    payload = {"action": "value", "grid": grid_from_xlsx(path), "items": jsonable(items), "sources": sources}
    if columns is not None:
        payload["columns"] = columns
    return run_engine("markup", payload, culture)["values"]


def company_form(columns=None):
    cols = MarkupColumns(**columns) if columns else MarkupColumns()
    return CompanyForm("x_{PO}.xlsx", "시트", [
        CompanyColumn("품번", "품목.품번"), CompanyColumn("마크업%", M.RATE),
        CompanyColumn("본사 단가", M.PRICE), CompanyColumn("본사 금액", M.AMOUNT)], cols)


def view(checks):
    return [{"item": c.item, "source_value": c.source, "result_value": c.output, "ok": c.ok} for c in checks]


def written_json(rows):
    return [{k: (None if v is None else format(v, "f")) for k, v in row.items()} for row in rows]


# --- 마크업표 읽기: 같은 결과 ---

OK_CASES = [
    ("basic", [["VTS-1", 3, None], ["vts-2", 2.5, 2], ["X", 0, 0], ["Y", 10, 10]]),
    ("percent_format", [["A", F(0.03, "0%"), None], ["B", F(0.005, "0.00%"), 2], ["C", F(0.1, "0.0%"), None],
                        ["D", F(0.0, "0%"), None]]),
    ("percent_text", [["A", "0.5%", None], ["B", " 3 % ", None], ["C", "2.50", None], ["D", "1,0", None],
                      ["E", "3", None], ["F", "0%", None], ["G", ".5%", None], ["H", "5.", None], ["I", "+4", None]]),
    ("places_variants", [["A", 1, "2"], ["B", 1, 2.0], ["C", 1, " 3 "], ["D", 1, "  "], ["E", 1, 0],
                         ["F", 1, "4.0"], ["G", 1, "1e1"]]),
    ("part_variants", [[12345.0, 3, None], ["  Abc  ", 1, None], [67890, 2, None], ["ÄÖ", 1, None],
                       [1.5, 1, None], ["007", 2, None], [datetime(2026, 1, 2), 3, None]]),
    ("literal_percent_is_not_percent", [["A", F(5, '0.00"%"'), None], ["B", F(2, "0\\%"), None]]),
    ("percent_after_quoted_text", [["A", F(0.04, '"율 "0%'), None]]),
    ("blank_lines_and_memo_column", [["A", 1, None], None, [None, None, None, "메모만 있는 줄"], ["B", 2, None], []]),
    ("single_row", [["ONLY", 1.5, None]]),
    ("header_only", []),
    ("rate_ten_and_zero_text", [["A", "10%", None], ["B", "0", None], ["C", "10.000", None]]),
    ("rate_float_noise", [["A", 2.675, None], ["B", 0.1 + 2, None]]),
]


@pytest.mark.parametrize("rows", [c[1] for c in OK_CASES], ids=[c[0] for c in OK_CASES])
def test_read_same_as_python(rows, run_engine, tmp_path):
    status, _ = same_read(run_engine, markup_file(tmp_path, rows))
    assert status == "ok"


def test_read_known_values(run_engine, tmp_path):
    """같은 결과이면서 값 자체도 맞는지: 3 → 3%, 백분율 서식 0.005 → 0.5%, 글자 0.5% → 0.5%, 반올림자리."""
    path = markup_file(tmp_path, [["A", 3, None], ["B", F(0.005, "0.00%"), 2], ["C", "0.5%", None], ["D", 1, "  "]])

    _, rows = same_read(run_engine, path)

    assert [(r["part"], r["rate"], r["places"]) for r in rows] == [
        ("A", "3", None), ("B", "0.5", 2), ("C", "0.5", None), ("D", "1", None)]


def test_read_header_variants(run_engine, tmp_path):
    reordered = markup_file(tmp_path, [[2, "x", 3, "A"], [None, "y", 4, "b"]], ("반올림자리", "메모", "마크업%", "파트넘버"), "a.xlsx")
    padded = markup_file(tmp_path, [["A", 1], ["B", 2]], (" 파트넘버 ", "마크업% "), "b.xlsx")
    no_places = markup_file(tmp_path, [["A", 1]], ("파트넘버", "마크업%"), "c.xlsx")

    for path in (reordered, padded, no_places):
        assert same_read(run_engine, path)[0] == "ok"


def test_read_uses_first_sheet_not_active(run_engine, tmp_path):
    path = markup_file(tmp_path, [["A", 1, None]], extra_first_sheet=["메모", "다른 표"])

    status, message = same_read(run_engine, path)

    assert status == "error" and "첫 줄(머리글)" in message


ERROR_CASES = [
    ("duplicate_part", [["A", 1, None], ["a ", 2, None]]),
    ("duplicate_number_vs_text", [[12345, 1, None], ["12345", 2, None], ["  12345.0", 3, None]]),
    ("duplicate_casefold", [["Straße", 1, None], ["STRASSE", 2, None]]),
    ("duplicate_after_bad_row", [["A", "bad", None], ["A", 1, None], ["A", 2, None]]),
    ("rate_empty", [["A", None, None], ["B", "  ", None]]),
    ("rate_bool", [["A", True, None]]),
    ("rate_range", [["A", 11, None], ["B", -1, None], ["C", 10.01, None], ["D", "11%", None], ["E", F(0.5, "0%"), None]]),
    ("rate_unmarked_fraction", [["A", 0.03, None], ["B", 0.5, None], ["C", "0.005", None]]),
    ("rate_float_noise_fraction", [["A", 0.1 + 0.2, None]]),
    ("rate_not_number", [["A", "abc", None], ["B", "3 percent", None], ["C", "%", None], ["D", "NaN", None],
                         ["E", "Infinity", None], ["F", "1.2.3", None]]),
    # (지금 판은 28자리를 넘는 정수 율(1e30)에서 MarkupError 가 아니라 decimal.InvalidOperation 으로 죽는다 - 같게 맞출 대상이 아님)
    ("rate_huge_and_tiny", [["A", "1e20", None], ["B", "1e-40", None], ["C", "1234567890123456789012345678", None]]),
    ("rate_date", [["A", datetime(2026, 1, 2), None]]),
    ("rate_literal_percent_fraction", [["A", F(0.5, '0.00"%"'), None]]),
    ("places_bad", [["A", 1, 11], ["B", 1, -1], ["C", 1, 2.5], ["D", 1, "x"], ["E", 1, True], ["F", 1, "11"],
                    ["G", 1, "1e30"], ["H", 1, datetime(2026, 1, 2)]]),
    ("part_blank", [[None, 3, None], ["  ", 3, None], [None, None, 5]]),
    ("part_bool", [[True, 3, None]]),
    ("label_of_number_part", [[12345.0, "bad", None]]),
    ("missing_rate_header", [["A", "x"]]),
    ("more_than_five", [[f"P{i}", "bad", None] for i in range(7)]),
    ("exactly_five", [[f"P{i}", "bad", None] for i in range(5)]),
]


@pytest.mark.parametrize("rows", [c[1] for c in ERROR_CASES], ids=[c[0] for c in ERROR_CASES])
def test_read_errors_same_text_as_python(rows, run_engine, tmp_path):
    header = ("파트넘버", "메모") if rows == [["A", "x"]] else HEADER
    status, message = same_read(run_engine, markup_file(tmp_path, rows, header))
    assert status == "error"
    assert message.startswith("마크업표.xlsx ")


def test_read_error_messages_are_the_known_ones(run_engine, tmp_path):
    """문구를 눈으로도 고정한다(지금 판 문구 그대로)."""
    path = markup_file(tmp_path, [["A", 1, None], ["a", 2, None], ["B", 11, None], ["C", 0.03, None], ["D", 1, 11]])

    status, message = same_read(run_engine, path)

    assert status == "error"
    assert message == (
        "마크업표.xlsx 를 고쳐야 합니다: "
        "마크업표.xlsx 3번째 줄: 파트넘버 'a' 가 2번째 줄('A')에도 있습니다(대소문자 · 앞뒤 공백은 같은 파트로 봄. 한 파트는 한 줄만); "
        "마크업표.xlsx 4번째 줄 (B): 마크업% 11 는 0 이상 10 이하여야 합니다; "
        "마크업표.xlsx 5번째 줄 (C): 0.03 은 0.03% 로 읽힙니다. 3% 면 3 을, 1% 미만이면 0.5% 처럼 % 를 붙여 적으세요; "
        "마크업표.xlsx 6번째 줄 (D): 반올림자리 '11' 는 0 부터 10 까지의 정수여야 합니다(비우면 반올림하지 않음)")


def test_read_no_cells_at_all(run_engine, tmp_path):
    wb = openpyxl.Workbook()
    path = tmp_path / "빈.xlsx"
    wb.save(path)

    status, message = same_read(run_engine, path)

    assert status == "error"
    assert message == "마크업표.xlsx 첫 줄(머리글)에 '파트넘버' 와 '마크업%' 열이 있어야 합니다."


def test_read_empty_first_row(run_engine, tmp_path):
    path = markup_file(tmp_path, [["A", 1, None]], header=(None, None, None))

    assert same_read(run_engine, path)[0] == "error"


@pytest.mark.parametrize("rows", [OK_CASES[0][1], OK_CASES[1][1], OK_CASES[2][1], ERROR_CASES[6][1], ERROR_CASES[9][1]],
                         ids=["basic", "percent_format", "percent_text", "rate_range", "rate_not_number"])
def test_read_culture_de_de_same_result(rows, run_engine, tmp_path):
    path = markup_file(tmp_path, rows)

    assert engine_read(run_engine, path, "de-DE") == python_read(path)


# --- 값 계산: 같은 결과 ---

STD_ROWS = [
    ["VTS-1", 3, None],
    ["vts-2", 2.5, 2],
    ["R-0", 0, None],
    ["R-HALF", 0, 2],
    ["ROUND0", 1, 0],
    ["TEN", 10, 10],
    [12345, 2, None],
    ["HALFPCT", F(0.005, "0.00%"), 3],
    ["LONGRATE", "1.234567", None],
]

ITEMS = [
    {"품번": "VTS-1", "단가": Decimal("100.00"), "수량": 3},
    {"품번": " VTS-2 ", "단가": Decimal("12.345"), "수량": 10},
    {"품번": "R-HALF", "단가": Decimal("12.345"), "수량": 2},          # 사사오입: 12.345 → 12.35
    {"품번": "R-HALF", "단가": Decimal("12.344"), "수량": 2},
    {"품번": "R-HALF", "단가": Decimal("-12.345"), "수량": 2},         # 음수도 사사오입(0 에서 먼 쪽)
    {"품번": "ROUND0", "단가": Decimal("2.5"), "수량": 1},
    {"품번": 12345.0, "단가": 19.99, "수량": "5"},                       # 숫자 칸 품번 · 실수 단가 · 글자 수량
    {"품번": "12345", "단가": "1,234.50", "수량": 2},                    # 글자 단가
    {"품번": "NOPE", "단가": Decimal("1"), "수량": 1},                   # 표에 없음 → 모두 None
    {"품번": "VTS-1", "단가": None, "수량": 1},
    {"품번": "VTS-1", "단가": Decimal("1"), "수량": None},
    {"품번": "VTS-1"},
    {"품번": "VTS-1", "단가": 0.1 + 0.2, "수량": 1},                     # 엑셀처럼 유효 15자리 → 0.3
    {"품번": "VTS-1", "단가": Decimal("-5.00"), "수량": -2},
    {"품번": "VTS-1", "단가": "abc", "수량": "x"},
    {"품번": "VTS-1", "단가": True, "수량": True},
    {"품번": None, "단가": Decimal("1"), "수량": 1},
    {"품번": "VTS-1", "단가": Decimal("123456789012.123456"), "수량": 1000000},
    {"품번": "VTS-1", "단가": Decimal("0"), "수량": 0},
    {"품번": "TEN", "단가": 100, "수량": 7},
    {"품번": "HALFPCT", "단가": Decimal("1000.0005"), "수량": 3},
    {"품번": "LONGRATE", "단가": Decimal("99.99"), "수량": 3},
    {"품번": "R-0", "단가": Decimal("33.333"), "수량": 3},
    None,
]


def test_values_same_as_python(run_engine, tmp_path):
    path = markup_file(tmp_path, STD_ROWS)

    got = engine_values(run_engine, path, ITEMS, SOURCES)

    assert got == python_values(path, ITEMS, SOURCES)


def test_values_known(run_engine, tmp_path):
    """값 자체: 사사오입 12.345 → 12.35, 반올림자리가 없으면 정확한 값(자리 그대로)."""
    path = markup_file(tmp_path, STD_ROWS)
    items = [{"품번": "R-HALF", "단가": Decimal("12.345"), "수량": 2}, {"품번": "VTS-1", "단가": Decimal("100.00"), "수량": 3}]

    got = engine_values(run_engine, path, items, SOURCES)

    assert got == [["0", "12.35", "24.70"], ["3", "103.0000", "309.0000"]]


def test_values_custom_columns_and_default_columns(run_engine, tmp_path):
    path = markup_file(tmp_path, STD_ROWS)
    columns = {"part": "PN", "price": "Unit", "quantity": "Qty"}
    items = [{"PN": "vts-1", "Unit": Decimal("10.00"), "Qty": 4}, {"품번": "VTS-1", "단가": Decimal("10.00"), "수량": 4}]

    assert engine_values(run_engine, path, items, SOURCES, columns) == python_values(path, items, SOURCES, columns)
    assert engine_values(run_engine, path, items, SOURCES) == python_values(path, items, SOURCES)  # 열 이름 생략 = 품번 · 단가 · 수량


def test_values_empty_table_and_no_items(run_engine, tmp_path):
    empty = markup_file(tmp_path, [], name="e.xlsx")
    path = markup_file(tmp_path, STD_ROWS)

    assert engine_values(run_engine, empty, ITEMS[:3], SOURCES) == python_values(empty, ITEMS[:3], SOURCES)
    assert engine_values(run_engine, path, [], SOURCES) == []


def test_values_culture_de_de(run_engine, tmp_path):
    path = markup_file(tmp_path, STD_ROWS)

    got = engine_values(run_engine, path, ITEMS, SOURCES, culture="de-DE")

    assert got == python_values(path, ITEMS, SOURCES)


@pytest.mark.parametrize("part, price", [
    ("VTS-1", "99999999999999999999999999.99"),    # 크기가 [decimal] 최대(약 7.9e28)를 넘는 단가
    ("LONGRATE", "1.1234567890123456789012"),      # 소수 22자리 x 율 1.234567(몫 8자리) = 30자리: [decimal] 소수 28자리를 넘음
    ("LONGRATE", "0.0000000000000000000000001"),   # 계수는 작아도 소수 33자리
], ids=["too_big", "too_many_places", "tiny_but_scale_33"])
def test_value_too_long_for_decimal_is_an_error_not_a_rounding(part, price, run_engine, tmp_path):
    """무설치판 [decimal] 은 28~29자리까지다. 정확히 담을 수 없으면 조용히 반올림하지 않고 멈춘다(지금 판은 계산해 낸다)."""
    path = markup_file(tmp_path, [["VTS-1", 3, None], ["LONGRATE", "1.234567", None]])
    items = [{"품번": part, "단가": Decimal(price), "수량": 1}]

    assert python_values(path, items, [M.PRICE])[0][0] is not None
    with pytest.raises(EngineError, match="28자리"):
        engine_values(run_engine, path, items, [M.PRICE])


# --- 없는 파트 · 빈 품번 ---

def missing_case(run_engine, path, table_items, form, columns=None):
    result = SimpleNamespace(table=table_items)
    table = table_of(path)
    expected = (M.missing_reason(result, form, table), M.missing_parts(result, form, table), M.blank_parts(result, form))
    out = run_engine("markup", {"action": "missing", "grid": grid_from_xlsx(path),
                                "result": {"table": jsonable(table_items)},
                                "form": None if form is None else _company_form(form)})
    assert (out["reason"], out["missing"], out["blanks"]) == expected
    return expected


def test_missing_part_reason(run_engine, tmp_path):
    """표에 없는 파트 → 사유에 그 파트넘버를 적고, 0% 로 넘기지 않는다."""
    path = markup_file(tmp_path, STD_ROWS)
    items = [{"품번": "VTS-1", "단가": Decimal("1"), "수량": 1}, {"품번": "ZZ-9", "단가": Decimal("1"), "수량": 1}]

    reason, missing, _ = missing_case(run_engine, path, items, company_form())

    assert "ZZ-9" in reason and missing == ["ZZ-9"]
    assert M.markup_value(M.RATE, items[1], MarkupColumns(), table_of(path)) is None
    assert engine_values(run_engine, path, [items[1]], SOURCES) == [[None, None, None]]


def test_missing_reason_variants(run_engine, tmp_path):
    path = markup_file(tmp_path, STD_ROWS)
    item = lambda part: {"품번": part, "단가": Decimal("1"), "수량": 1}
    cases = {
        "all_present": [item("VTS-1"), item(" vts-2")],
        "two_missing_dedup": [item("A-1"), item("a-1"), item("B-2"), item("A-1")],
        "leading_zero": [item("0123"), item("X")],
        "blank_only": [item(None), item("  "), item("VTS-1")],
        "both": [item("QQ"), item(None)],
        "numeric_part": [item(777.0), item(12345.0)],
        "empty_table": [],
        "no_table": None,
    }
    for name, items in cases.items():
        missing_case(run_engine, path, items, company_form())


def test_missing_reason_not_used_without_markup_columns(run_engine, tmp_path):
    path = markup_file(tmp_path, STD_ROWS)
    plain = CompanyForm("x.xlsx", "시트", [CompanyColumn("품번", "품목.품번")])
    items = [{"품번": "ZZ", "단가": Decimal("1"), "수량": 1}, {"품번": None}]

    assert missing_case(run_engine, path, items, plain) == (None, [], 0)
    assert missing_case(run_engine, path, items, None) == (None, [], 0)


def test_missing_reason_custom_part_column(run_engine, tmp_path):
    path = markup_file(tmp_path, STD_ROWS)
    form = company_form({"part": "PN", "price": "Unit", "quantity": "Qty"})
    items = [{"PN": "NOPE", "Unit": Decimal("1"), "Qty": 1}, {"PN": "VTS-1", "Unit": Decimal("1"), "Qty": 1}]

    reason, _, _ = missing_case(run_engine, path, items, form)

    assert "NOPE" in reason


# --- 검산: 같은 목록 ---

def checks_case(run_engine, path, items, rows_written, columns=None, culture=None):
    table = table_of(path)
    cols = MarkupColumns(**columns) if columns else None
    expected = view(M.markup_checks(SimpleNamespace(table=items), rows_written, table, cols))
    payload = {"action": "checks", "grid": grid_from_xlsx(path), "result": {"table": jsonable(items)},
               "rows_written": written_json(rows_written)}
    if columns is not None:
        payload["columns"] = columns
    got = run_engine("markup", payload, culture)
    assert got == expected
    return expected


def correct_rows(path, items, columns=None):
    table = table_of(path)
    cols = MarkupColumns(**columns) if columns else MarkupColumns()
    return [{s: M.markup_value(s, item, cols, table) for s in SOURCES} for item in items]


def test_checks_all_correct(run_engine, tmp_path):
    path = markup_file(tmp_path, STD_ROWS)
    items = [i for i in ITEMS if i is not None and i.get("품번") in ("VTS-1", " VTS-2 ", "R-HALF", "ROUND0", "TEN", "R-0")
             and isinstance(i.get("단가"), (Decimal, int)) and i.get("수량") is not None and not isinstance(i.get("단가"), bool)]

    checks = checks_case(run_engine, path, items, correct_rows(path, items))

    assert checks and all(c["ok"] for c in checks)


def test_checks_full_item_set_including_bad_inputs(run_engine, tmp_path):
    path = markup_file(tmp_path, STD_ROWS)
    items = [i for i in ITEMS if i is not None]

    checks_case(run_engine, path, items, correct_rows(path, items))


def test_checks_wrong_values_and_missing_keys(run_engine, tmp_path):
    path = markup_file(tmp_path, STD_ROWS)
    items = [{"품번": "VTS-1", "단가": Decimal("100.00"), "수량": 3}] * 7
    good = correct_rows(path, items)[0]
    rows = [
        dict(good),
        {**good, M.RATE: Decimal("4")},                       # 율이 다름
        {**good, M.PRICE: Decimal("103.0001")},               # 단가가 다름
        {**good, M.AMOUNT: Decimal("309.01")},                # 금액이 다름
        {**good, M.PRICE: None, M.AMOUNT: None},              # 값이 None
        {M.PRICE: good[M.PRICE]},                             # 율 · 금액 열 없음(율은 표 값으로 대조)
        {M.RATE: good[M.RATE], M.AMOUNT: good[M.AMOUNT]},     # 단가 열 없음(단가는 직접 계산)
    ]

    checks = checks_case(run_engine, path, items, rows)

    assert [c["ok"] for c in checks if c["item"].endswith("마크업% (VTS-1)")] == [True, False, True, True, True, True, True]


def test_checks_rows_written_shorter_than_table_and_empty_items(run_engine, tmp_path):
    path = markup_file(tmp_path, STD_ROWS)
    items = [{"품번": "VTS-1", "단가": Decimal("100.00"), "수량": 3}, {"품번": "R-HALF", "단가": Decimal("12.345"), "수량": 2}]

    checks_case(run_engine, path, items, correct_rows(path, items)[:1])
    assert checks_case(run_engine, path, [], []) == []


def test_checks_unknown_part_price_and_quantity_gaps(run_engine, tmp_path):
    path = markup_file(tmp_path, STD_ROWS)
    items = [{"품번": "NOPE", "단가": Decimal("1"), "수량": 1},
             {"품번": "VTS-1", "단가": None, "수량": 1},
             {"품번": "VTS-1", "단가": Decimal("2"), "수량": None},
             {"품번": None, "단가": Decimal("2"), "수량": 1}]
    rows = [{M.RATE: Decimal("3"), M.PRICE: Decimal("1"), M.AMOUNT: Decimal("1")}] * 4

    checks = checks_case(run_engine, path, items, rows)

    assert checks[0] == {"item": "1줄 마크업% (NOPE)", "source_value": "마크업표에 없음", "result_value": "", "ok": False}
    assert any(c["item"] == "4줄 마크업% (" + M.BLANK_PART + ")" for c in checks)


def test_checks_custom_columns_and_places_text(run_engine, tmp_path):
    path = markup_file(tmp_path, STD_ROWS)
    columns = {"part": "PN", "price": "Unit", "quantity": "Qty"}
    items = [{"PN": "r-half", "Unit": Decimal("12.345"), "Qty": 2}, {"PN": "vts-2", "Unit": Decimal("12.345"), "Qty": 10}]

    checks = checks_case(run_engine, path, items, correct_rows(path, items, columns), columns)

    assert any("소수 2자리 반올림 12.35" in c["source_value"] for c in checks)


def test_checks_culture_de_de(run_engine, tmp_path):
    path = markup_file(tmp_path, STD_ROWS)
    items = [i for i in ITEMS if i is not None]

    checks_case(run_engine, path, items, correct_rows(path, items), culture="de-DE")


# --- 지문 ---

def stamp_python(path, parts):
    result = SimpleNamespace(table=[{"품번": p} for p in parts])
    return M.markup_stamp(result, company_form(), table_of(path))


def stamp_engine(run_engine, path, parts):
    return run_engine("markup", {"action": "fingerprint", "grid": grid_from_xlsx(path), "parts": parts})["fingerprint"]


def test_fingerprint_changes_only_with_the_documents_parts(run_engine, tmp_path):
    base = markup_file(tmp_path, [["A", 3, None], ["B", 2, 2], ["C", 1, None]], name="1.xlsx")
    other_part = markup_file(tmp_path, [["A", 3, None], ["B", 2, 2], ["C", 9, None]], name="2.xlsx")
    extra_part = markup_file(tmp_path, [["A", 3, None], ["B", 2, 2], ["C", 1, None], ["D", 5, None]], name="3.xlsx")
    rate_changed = markup_file(tmp_path, [["A", 4, None], ["B", 2, 2], ["C", 1, None]], name="4.xlsx")
    places_changed = markup_file(tmp_path, [["A", 3, None], ["B", 2, 3], ["C", 1, None]], name="5.xlsx")
    parts = ["A", "b", " A "]

    values = {n: stamp_engine(run_engine, p, parts) for n, p in
              {"base": base, "other_part": other_part, "extra": extra_part, "rate": rate_changed, "places": places_changed}.items()}

    assert values["base"] == values["other_part"] == values["extra"]       # 이 문서에 없는 파트가 바뀌어도 그대로
    assert len({values["base"], values["rate"], values["places"]}) == 3    # 쓰는 파트의 율 · 반올림자리가 바뀌면 달라짐
    assert len(values["base"]) == 16


def test_fingerprint_same_as_python_value(run_engine, tmp_path):
    """같은 값일 필요는 없지만 지금 판 markup_stamp 와 같게 만들어 둔다."""
    path = markup_file(tmp_path, [["A", 3, None], ["B", 0.5 + 1, 2], [12345, 1, None]])

    for parts in (["A"], ["B", "a", "NOPE"], [12345.0, None, "  "], [], ["Straße"]):
        assert stamp_engine(run_engine, path, parts) == stamp_python(path, parts)


def test_fingerprint_missing_part_differs_from_present(run_engine, tmp_path):
    without = markup_file(tmp_path, [["A", 3, None]], name="1.xlsx")
    with_b = markup_file(tmp_path, [["A", 3, None], ["B", 0, None]], name="2.xlsx")

    assert stamp_engine(run_engine, without, ["A", "B"]) != stamp_engine(run_engine, with_b, ["A", "B"])


# --- 회사 양식 행에서 쓴 값 골라내기 ---

def test_written_markup_same_as_python(run_engine):
    form = company_form()
    rows = [["P-1", Decimal("3"), Decimal("103.0000"), Decimal("309.0000")],
            ["P-2", None, Decimal("1.5"), None],
            ["P-3", Decimal("1")],
            []]

    got = run_engine("markup", {"action": "written", "form": _company_form(form), "rows": jsonable(rows)})

    assert got == written_json(M.written_markup(form, rows))


# --- PS 안 호출(JSON 을 거치지 않는 메모리 객체) ---

MEMORY_SCRIPT = """\
param([string]$Lib, [string]$In, [string]$Out)
. (Join-Path $Lib 'Values.ps1')
. (Join-Path $Lib 'Rules.ps1')
. (Join-Path $Lib 'ExtractText.ps1')
. (Join-Path $Lib 'ExtractGrid.ps1')
. (Join-Path $Lib 'Verify.ps1')
. (Join-Path $Lib 'Markup.ps1')
$data = [IO.File]::ReadAllText($In, [Text.Encoding]::UTF8) | ConvertFrom-Json
$table = Read-MarkupGrid $data.grid
$columns = @{ part = '품번'; price = '단가'; quantity = '수량' }
$item = [ordered]@{ '품번' = 'VTS-1'; '단가' = [decimal]'100.00'; '수량' = [long]3 }
$rate = Get-MarkupValue '마크업.율' $item $columns $table
$unit = Get-MarkupValue '마크업.단가' $item $columns $table
$amount = Get-MarkupValue '마크업.금액' $item $columns $table
$rows = New-Object 'System.Collections.Generic.List[object]'
$rows.Add(@{ '마크업.율' = $rate; '마크업.단가' = $unit; '마크업.금액' = $amount })
$items = New-Object 'System.Collections.Generic.List[object]'
$items.Add($item)
$result = [ordered]@{ table = $items }
$checks = Get-MarkupChecks $result $rows $table $columns
$form = @{ columns = @(@{ name = 'a'; value = '마크업.율' }); markup_columns = $columns }
$none = New-Object 'System.Collections.Generic.List[object]'
$none.Add([ordered]@{ '품번' = 'NOPE'; '단가' = [decimal]'1'; '수량' = [long]1 })
$reason = Get-MissingReason ([ordered]@{ table = $none }) $form $table
$double = [ordered]@{ '품번' = [double]12345; '단가' = [decimal]'10.00'; '수량' = [long]1 }      # 엑셀 숫자 칸 품번(실수 12345.0)
$doubleRate = Get-MarkupValue '마크업.율' $double $columns $table
$odd = New-OrderedMap                                                                       # 읽기를 거치지 않은 표: 율 11 은 범위 밖
$odd['x'] = @{ part = 'X'; rate = [decimal]11; places = $null }
$oddItems = New-Object 'System.Collections.Generic.List[object]'
$oddItems.Add([ordered]@{ '품번' = 'X'; '단가' = [decimal]'10.00'; '수량' = [long]1 })
$oddRows = New-Object 'System.Collections.Generic.List[object]'
$oddRows.Add(@{ '마크업.율' = [decimal]11 })
$oddChecks = Get-MarkupChecks ([ordered]@{ table = $oddItems }) $oddRows $odd $columns
$report = [ordered]@{
    rate = (ConvertTo-JsonValue $rate); unit = (ConvertTo-JsonValue $unit); amount = (ConvertTo-JsonValue $amount)
    checks = ($checks.ToArray()); reason = $reason; double_rate = (ConvertTo-JsonValue $doubleRate); odd_checks = ($oddChecks.ToArray())
}
[IO.File]::WriteAllText($Out, (ConvertTo-Json -InputObject $report -Depth 20), (New-Object Text.UTF8Encoding $false))
"""


def test_in_memory_objects(powershell_exe, tmp_path):
    """JSON 을 거치지 않은 메모리 객체(해시표 품목 · [decimal] 단가 · List)도 같은 결과."""
    path = markup_file(tmp_path, STD_ROWS)
    script = tmp_path / "memory.ps1"
    script.write_bytes(b"\xef\xbb\xbf" + MEMORY_SCRIPT.encode("utf-8"))
    source = tmp_path / "in.json"
    source.write_text(json.dumps({"grid": grid_from_xlsx(path)}, ensure_ascii=False), encoding="utf-8")
    out = tmp_path / "out.json"
    item = {"품번": "VTS-1", "단가": Decimal("100.00"), "수량": 3}
    table = table_of(path)
    unit, amount = M.markup_value(M.PRICE, item, MarkupColumns(), table), M.markup_value(M.AMOUNT, item, MarkupColumns(), table)
    rows = [{M.RATE: M.markup_value(M.RATE, item, MarkupColumns(), table), M.PRICE: unit, M.AMOUNT: amount}]

    done = run_powershell(powershell_exe, script, "-Lib", str(PORTABLE / "lib"), "-In", str(source), "-Out", str(out))

    assert done.returncode == 0, decode_output(done.stderr)
    got = json.loads(out.read_text(encoding="utf-8-sig"))
    assert (got["rate"], got["unit"], got["amount"]) == ("3", format(unit, "f"), format(amount, "f"))
    assert got["checks"] == view(M.markup_checks(SimpleNamespace(table=[item]), rows, table))
    none_item = {"품번": "NOPE", "단가": Decimal("1"), "수량": 1}
    assert got["reason"] == M.missing_reason(SimpleNamespace(table=[none_item]), company_form(), table)
    assert got["double_rate"] == "2"   # 실수 12345.0 품번은 숫자 칸 12345 와 같은 파트
    odd = {"x": M.MarkupRow("X", Decimal(11))}
    odd_item = {"품번": "X", "단가": Decimal("10.00"), "수량": 1}
    expected_odd = view(M.markup_checks(SimpleNamespace(table=[odd_item]), [{M.RATE: Decimal(11)}], odd))
    assert got["odd_checks"] == expected_odd and expected_odd[0]["ok"] is False   # 율 11 은 표 값과 같아도 범위(0~10) 밖이라 불일치
