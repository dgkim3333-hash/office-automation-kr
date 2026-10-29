"""무설치판 엔진 extract_text: 같은 글자를 넣으면 지금 판 extract_from_text 와 같은 결과를 낸다."""
from datetime import date
from decimal import Decimal

import pytest

from app import extract
from app.extract import NO_TEXT, extract_from_text
from app.portable import rules_to_portable
from app.rules import rules_from_text

from .conftest import REPO, EngineError, example_rules

SAMPLES = sorted((REPO / "samples" / "03_Invoice_DM_PDF").glob("*.pdf"))
UNMATCHED = "표 행패턴과 맞지 않는 줄이 있어 품목 건너뜀"


def _plain(value):
    """파이썬 값을 엔진 JSON 표기로: 금액은 str(Decimal), 날짜는 ISO, 나머지는 그대로."""
    if isinstance(value, Decimal):
        return str(value)
    if isinstance(value, date):
        return value.isoformat()
    return value


def python_view(result) -> dict:
    return {
        "doc_type": result.doc_type.name if result.doc_type else None,
        "fields": {k: _plain(v) for k, v in result.fields.items()},
        "table": None if result.table is None else [{k: _plain(v) for k, v in row.items()} for row in result.table],
        "warnings": list(result.warnings),
        "review_reason": result.review_reason,
    }


def engine_view(out: dict) -> dict:
    return {k: out[k] for k in ("doc_type", "fields", "table", "warnings", "review_reason")}


def _invoice_text() -> str:
    return extract.read_text(next(p for p in SAMPLES if p.name.startswith("INV-")))


def test_samples_present():
    assert len(SAMPLES) == 6


@pytest.mark.parametrize("pdf", SAMPLES, ids=[p.name for p in SAMPLES])
def test_same_as_python_on_samples(pdf, run_engine):
    """샘플 PDF 6개 각각: 같은 글자 → doc_type · fields · table · warnings · review_reason 이 같다."""
    rules = example_rules()
    text = extract.read_text(pdf)
    expected = python_view(extract_from_text(text, rules))

    out = run_engine("extract_text", {"rules": rules_to_portable(rules), "text": text})

    assert engine_view(out) == expected
    assert out["form"] is None and out["is_excel"] is False
    if pdf.name == "Scan_signed_packing_list.pdf":
        assert out["review_reason"] == NO_TEXT
    if pdf.name.startswith(("INV-", "DM-")):
        assert out["table"], "품목 표가 비어 있으면 비교가 약하다"


def test_unmatched_row_goes_to_review(run_engine):
    """표 사이에 행패턴과 안 맞는 줄이 하나 있으면 표 전체를 버리고(table None) 경고를 남긴다."""
    rules = example_rules()
    lines = _invoice_text().splitlines()
    start = next(i for i, line in enumerate(lines) if "Line Item Description" in line)
    lines.insert(start + 2, "Back-ordered items follow")
    text = "\n".join(lines)
    expected = python_view(extract_from_text(text, rules))
    assert expected["table"] is None  # 지금 판도 버린다

    out = run_engine("extract_text", {"rules": rules_to_portable(rules), "text": text})

    assert out["table"] is None
    assert f"{UNMATCHED}: Back-ordered items follow" in out["warnings"]
    assert engine_view(out) == expected


def test_culture_de_de_same_result(run_engine):
    """독일어 문화권(소수점이 쉼표)으로 바꿔 돌려도 금액 · 날짜가 기본 문화권과 같다."""
    rules = rules_to_portable(example_rules())
    text = _invoice_text()

    default = run_engine("extract_text", {"rules": rules, "text": text})
    german = run_engine("extract_text", {"rules": rules, "text": text}, culture="de-DE")

    assert german == default
    assert "." in default["fields"]["합계금액"]
    assert default["fields"]["발행일"].count("-") == 2


QTY_RULES = """\
문서종류:
  T:
    판별키워드: ["TEST DOC"]
    번호필드: 번호
    필드:
      번호: {패턴: "No: (\\\\S+)"}
    표:
      시작키워드: "BEGIN"
      끝키워드: "END"
      행패턴: "^(\\\\S+)\\\\s+([\\\\d.]+)$"
      열: [품번, 수량]
      수량열: [수량]
    파일명: "T_{번호}.pdf"
"""


@pytest.mark.parametrize("qty, table", [
    ("2.0", [{"품번": "P-1", "수량": 1}, {"품번": "P-2", "수량": 2}]),  # 2.0 은 정수로 본다
    ("1.5", None),
])
def test_quantity_must_be_integer(qty, table, run_engine):
    """수량이 정수가 아닌 줄은 행 불일치 → 표 전체 None (지금 판과 같음)."""
    rules = rules_from_text(QTY_RULES)
    text = f"TEST DOC\nNo: A-1\nBEGIN\nP-1 1\nP-2 {qty}\nEND\n"
    expected = python_view(extract_from_text(text, rules))
    assert expected["table"] == table

    out = run_engine("extract_text", {"rules": rules_to_portable(rules), "text": text})

    assert engine_view(out) == expected
    if table:
        assert all(type(row["수량"]) is int for row in out["table"])  # 글자 "2" 가 아니라 숫자


@pytest.mark.parametrize("line, table", [
    ("x P-1 1", None),                          # 줄 중간에서만 맞음 → 불일치
    ("P-1 1 extra", [{"품번": "P-1", "수량": 1}]),  # 끝은 고정하지 않음 → 맞음(한 줄짜리 표도 배열)
])
def test_row_pattern_anchored_at_start_only(line, table, run_engine):
    """행패턴은 파이썬 re.match 처럼 줄 처음에서만 맞춘다(^ 를 안 써도). 끝은 고정하지 않는다."""
    rules = rules_from_text(QTY_RULES.replace('"^(\\\\S+)\\\\s+([\\\\d.]+)$"', '"(\\\\S+)\\\\s+(\\\\d+)"'))
    assert rules.doc_types[0].table.row_pattern.pattern == r"(\S+)\s+(\d+)"
    text = f"TEST DOC\nNo: A-1\nBEGIN\n{line}\nEND"
    expected = python_view(extract_from_text(text, rules))
    assert expected["table"] == table

    out = run_engine("extract_text", {"rules": rules_to_portable(rules), "text": text})

    assert engine_view(out) == expected


@pytest.mark.parametrize("body", ["", "\n   \n\n"], ids=["no_lines", "blank_lines"])
def test_empty_table_is_empty_list(body, run_engine):
    """시작 · 끝키워드 사이에 품목 줄이 없으면(빈 줄만 있어도) 표는 None 이 아니라 빈 목록 [] (지금 판과 같음)."""
    rules = rules_from_text(QTY_RULES)
    text = f"TEST DOC\nNo: A-1\nBEGIN\n{body}END"
    expected = python_view(extract_from_text(text, rules))
    assert expected["table"] == []

    out = run_engine("extract_text", {"rules": rules_to_portable(rules), "text": text})

    assert out["table"] == []
    assert engine_view(out) == expected


def test_no_type_and_missing_number(run_engine):
    """판별키워드가 없으면 판별 실패, 번호필드가 안 맞으면 번호필드 미추출 사유(글자까지 지금 판과 같음)."""
    rules = rules_from_text(QTY_RULES)
    portable = rules_to_portable(rules)
    for text in ("nothing here", "TEST DOC\nBEGIN\nP-1 1\nEND"):
        expected = python_view(extract_from_text(text, rules))
        assert engine_view(run_engine("extract_text", {"rules": portable, "text": text})) == expected


def test_bad_regex_reported(run_engine):
    """.NET 이 컴파일하지 못하는 정규식은 「규칙 정규식 오류: <이름>」 으로 멈춘다."""
    portable = rules_to_portable(rules_from_text(QTY_RULES))
    portable["doc_types"][0]["fields"][0]["pattern"] = "No: (\\S+"
    with pytest.raises(EngineError, match="규칙 정규식 오류: .*번호"):
        run_engine("extract_text", {"rules": portable, "text": "TEST DOC"})
