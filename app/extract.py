"""PDF 텍스트 추출 → 문서종류 판별 → 필드·표 추출. 파일을 옮기거나 쓰지 않는다."""
from dataclasses import dataclass, field
from datetime import date, datetime
from decimal import Decimal, InvalidOperation
from pathlib import Path

import pdfplumber

HEAD_LINES = 30
DATE_FORMATS = ("%m/%d/%Y", "%Y-%m-%d")
NO_TEXT = "텍스트 없음 (스캔본, OCR 범위 밖)"


@dataclass
class ExtractResult:
    doc_type: object  # rules.DocType 또는 None
    fields: dict = field(default_factory=dict)
    table: list[dict] | None = None
    text_head: str = ""
    missing_required: bool = False
    warnings: list[str] = field(default_factory=list)
    has_text: bool = True
    source_form: str | None = None  # 엑셀 PO 면 엑셀양식 이름(고객), PDF 면 None
    is_excel: bool = False

    @property
    def review_reason(self) -> str | None:
        """확인필요로 보낼 사유. 정상이면 None."""
        if not self.has_text:
            return NO_TEXT
        if self.doc_type is None and self.is_excel:
            return "판별 실패 (규칙.yaml 의 엑셀양식 중 맞는 고객이 없음. 05장대로 이 고객의 엑셀 양식을 가르치세요)"
        if self.doc_type is None:
            return "판별 실패 (판별키워드가 하나도 없음)"
        if self.missing_required:
            return f"번호필드 미추출 ('{self.doc_type.number_field}' 패턴이 맞지 않음)"
        return None

    @property
    def number(self):
        return self.fields.get(self.doc_type.number_field) if self.doc_type else None


def read_pages(pdf_path) -> list[str]:
    with pdfplumber.open(Path(pdf_path)) as pdf:
        return [page.extract_text() or "" for page in pdf.pages]


def read_text(pdf_path) -> str:
    return "\n".join(read_pages(pdf_path)).strip()


def detect_type(text, rules):
    """규칙에 적힌 순서대로 판별키워드를 찾는다(대소문자 무시). 없으면 None."""
    lowered = text.lower()
    for doc_type in rules.doc_types:
        if any(k.lower() in lowered for k in doc_type.keywords):
            return doc_type
    return None


def to_date(value: str) -> date | None:
    for fmt in DATE_FORMATS:
        try:
            return datetime.strptime(value.strip(), fmt).date()
        except ValueError:
            pass
    return None


def to_amount(value: str) -> Decimal | None:
    try:
        return Decimal(value.replace(",", "").replace("$", "").strip())
    except InvalidOperation:
        return None


def extract(pdf_path, rules) -> ExtractResult:
    pages = read_pages(pdf_path)
    return extract_from_text("\n".join(pages).strip(), rules, first_page=pages[0] if pages else "")


def extract_from_text(text, rules, first_page=None) -> ExtractResult:
    head = "\n".join(text.splitlines()[:HEAD_LINES])
    if not text.strip():
        return ExtractResult(doc_type=None, text_head=head, has_text=False)

    doc_type = detect_type(first_page if first_page is not None else text, rules)
    result = ExtractResult(doc_type=doc_type, text_head=head)
    if doc_type is None:
        return result

    for name, spec in doc_type.fields.items():
        result.fields[name] = _field_value(spec, text, result.warnings)
    result.missing_required = result.fields[doc_type.number_field] is None
    if doc_type.table is not None:
        result.table = _table_rows(doc_type.table, text, result.warnings)
    return result


def _field_value(spec, text, warnings):
    match = spec.pattern.search(text)
    if not match or not (match.group(1) or "").strip():
        warnings.append(f"필드 '{spec.name}' 못 찾음")
        return None
    raw = match.group(1).strip()
    if spec.fmt == "날짜":
        value = to_date(raw)
    elif spec.fmt == "금액":
        value = to_amount(raw)
    else:
        return raw
    if value is None:
        warnings.append(f"필드 '{spec.name}' 값 '{raw}' 를 {spec.fmt}로 바꾸지 못함")
    return value


def _table_rows(table, text, warnings):
    lines = text.splitlines()
    start = next((i for i, line in enumerate(lines) if table.start in line), None)
    if start is None:
        warnings.append(f"표 시작키워드 '{table.start}' 없음, 품목 건너뜀")
        return None
    end = next((i for i in range(start + 1, len(lines)) if table.end in lines[i]), None)
    if end is None:
        warnings.append(f"표 끝키워드 '{table.end}' 없음, 품목 건너뜀")
        return None

    rows = []
    for line in lines[start + 1:end]:
        if not line.strip():
            continue
        row = _parse_row(line.strip(), table)
        if row is None:
            warnings.append(f"표 행패턴과 맞지 않는 줄이 있어 품목 건너뜀: {line.strip()}")
            return None
        rows.append(row)
    return rows


def _parse_row(line, table):
    """행패턴으로 한 줄을 나누고 금액열·수량열을 숫자로 바꾼다. 안 맞으면 None."""
    match = table.row_pattern.match(line)
    if not match:
        return None
    row = {name: (value or "").strip() for name, value in zip(table.columns, match.groups())}
    for name in table.amount_columns:
        row[name] = to_amount(row[name])
        if row[name] is None:
            return None
    for name in table.quantity_columns:
        amount = to_amount(row[name])
        if amount is None or amount != amount.to_integral_value():
            return None
        row[name] = int(amount)
    return row
