"""검산: 결과 값이 원본에 그대로 있는지, 품목 줄 · 합계가 맞는지 확인한다. 파일을 옮기거나 쓰지 않는다.

틀린 숫자가 조용히 통과하지 않는 것이 우선이다. 애매하면 불일치(확인필요)로 본다.
"""
import re
from dataclasses import dataclass, field
from datetime import date, datetime
from decimal import ROUND_HALF_EVEN, ROUND_HALF_UP, Decimal, InvalidOperation
from pathlib import Path

import openpyxl

from app.extract import read_text, to_date

# 원본 글자에서 숫자 · 날짜 토큰을 찾는 정규식. 숫자는 쉼표 · $ 를 뺀 Decimal(소수 자리 유지)로 바꾼다.
NUMBER = re.compile(r"-?\d[\d,]*(?:\.\d+)?")
DATE = re.compile(r"(?=(?<!\d)(\d{1,4}[/\-]\d{1,2}[/\-]\d{1,4})(?!\d))")  # 겹쳐서 찾음. to_date 로 바뀌는 것만 날짜
NOT_FOUND = "원본에 없음"
LINE_COLUMNS = ("수량", "단가", "금액")


@dataclass
class Check:
    item: str
    source: str
    output: str
    ok: bool | None  # None = 해당 없음


@dataclass
class _Source:
    """원본에서 다시 읽은 값 모음."""
    parts: list = field(default_factory=list)  # 원본 글자 조각(번호필드 대조용)
    numbers: dict = field(default_factory=dict)  # Decimal → 그 값을 준 원본 토큰들
    dates: dict = field(default_factory=dict)    # date(또는 datetime) → 원본 토큰들
    has_places: bool = False                     # PDF 만 소수 자리 정보가 있다

    @property
    def text(self) -> str:
        return "\n".join(self.parts)


def verify(result, source_path) -> list[Check]:
    """결과 하나를 원본과 대조한다. 실패 항목만 보려면 failed() 로 거른다."""
    if result.doc_type is None:
        return []
    source = _excel_source(source_path) if result.is_excel else _pdf_source(source_path)
    return _source_checks(result, source) + _line_checks(result.table) + _total_checks(result)


def failed(checks: list[Check]) -> list[Check]:
    return [c for c in checks if c.ok is False]


# --- 원본 다시 읽기 ---

def _pdf_source(path) -> _Source:
    source = _Source(has_places=True)
    _scan_text(source, read_text(Path(path)))
    return source


def _excel_source(path) -> _Source:
    source = _Source()
    wb = openpyxl.load_workbook(Path(path), data_only=True)
    try:
        for ws in wb.worksheets:
            for row in ws.iter_rows(values_only=True):
                for value in row:
                    _add_cell(source, value)
    finally:
        wb.close()
    return source


def _add_cell(source, value):
    if value is None or isinstance(value, bool):
        return
    if isinstance(value, str):
        _scan_text(source, value)
    elif isinstance(value, datetime):
        source.parts.append(value.isoformat(" "))
        _add(source.dates, value, value.isoformat(" "))
        if value.time() == datetime.min.time():
            _add(source.dates, value.date(), value.date().isoformat())
    elif isinstance(value, date):
        source.parts.append(value.isoformat())
        _add(source.dates, value, value.isoformat())
    else:
        number = _decimal(value)
        if number is not None:
            source.parts.append(_show(number))
            _add(source.numbers, number, _show(number))


def _scan_text(source, text):
    source.parts.append(text)
    for match in NUMBER.finditer(text):
        raw = match.group().rstrip(",")
        number = _decimal(raw.replace(",", ""))
        if number is None:
            continue
        _add(source.numbers, number, raw)
        if raw.startswith("-"):
            _add(source.numbers, -number, raw[1:])
    for match in DATE.finditer(text):
        raw = match.group(1)
        found = to_date(raw)
        if found is not None:
            _add(source.dates, found, raw)


def _add(bucket, key, raw):
    bucket.setdefault(key, [])
    if raw not in bucket[key]:
        bucket[key].append(raw)


# --- 1 · 2. 원문 대조와 소수 자리 ---

def _targets(result):
    """대조할 (항목 이름, 값). 숫자 · 날짜 값과 번호필드."""
    number_field = result.doc_type.number_field
    for name, value in result.fields.items():
        if value is not None and (name == number_field or _comparable(value)):
            yield name, value
    for n, row in enumerate(result.table or [], 1):
        for column, value in row.items():
            if _comparable(value):
                yield f"품목 {n}줄 {column}", value


def _comparable(value) -> bool:
    return isinstance(value, (Decimal, int, float, date)) and not isinstance(value, bool)


def _source_checks(result, source) -> list[Check]:
    checks = []
    for item, value in _targets(result):
        if isinstance(value, str):
            ok = bool(value) and value in source.text
            checks.append(Check(item, value if ok else NOT_FOUND, value, ok))
        elif isinstance(value, date):
            tokens = source.dates.get(value, [])
            checks.append(Check(item, tokens[0] if tokens else NOT_FOUND, _show(value), bool(tokens)))
        else:
            number = _decimal(value)
            tokens = source.numbers.get(number, []) if number is not None else []
            checks.append(Check(item, tokens[0] if tokens else NOT_FOUND, _show(value), bool(tokens)))
            if tokens and source.has_places and isinstance(value, Decimal):
                checks.append(_places_check(item, value, tokens))
    return checks


def _places_check(item, value, tokens) -> Check:
    """결과 소수 자리 수 = 같은 값을 준 원본 토큰의 소수 자리 수(그런 토큰이 하나라도 있으면 통과)."""
    want = _places(value)
    same = [t for t in tokens if _places(Decimal(t.replace(",", ""))) == want]
    return Check(f"{item} 소수 자리", (same or tokens)[0], _show(value), bool(same))


def _places(number: Decimal) -> int:
    return max(0, -number.as_tuple().exponent)


# --- 3. 줄 검산 ---

def _line_checks(table) -> list[Check]:
    checks = []
    for n, row in enumerate(table or [], 1):
        item = f"{n}줄 금액"
        quantity, price, amount = (_decimal(row.get(k)) for k in LINE_COLUMNS)
        if quantity is None or price is None or amount is None:
            checks.append(Check(item, "", _show(row.get("금액")), None))
            continue
        product = quantity * price
        checks.append(Check(item, f"{_show(quantity)} x {_show(price)} = {_show(product)}",
                            _show(amount), _rounds_to(product, amount)))
    return checks


def _rounds_to(product: Decimal, amount: Decimal) -> bool:
    """product 를 amount 의 소수 자리로 반올림(사사오입 · 은행원식 중 하나)하면 amount 인가."""
    step = Decimal(1).scaleb(amount.as_tuple().exponent)
    try:
        return any(product.quantize(step, rounding=mode) == amount for mode in (ROUND_HALF_UP, ROUND_HALF_EVEN))
    except InvalidOperation:
        return False


# --- 4. 합계 검산 ---

def _total_checks(result) -> list[Check]:
    spec = result.doc_type.check
    if spec is None:
        return []
    fields = result.fields
    item_sum = _item_sum(result.table)
    checks = []

    subtotal = None
    if spec.subtotal:
        subtotal = _decimal(fields.get(spec.subtotal))
        checks.append(_equal("품목 합 = 소계", item_sum, _show(item_sum), subtotal))

    if spec.total:
        total = _decimal(fields.get(spec.total))
        name = ("소계" if spec.subtotal else "품목 합") + (" + 부대비용" if spec.extras else "") + " = 합계"
        # 소계 값이 비면 품목 합으로 대신 계산한다(합계를 아무것도 대조하지 않는 일이 없게).
        base = subtotal if subtotal is not None else item_sum
        # 부대비용 칸이 비면 0 으로 본다. 실제로 있었다면 합계가 안 맞아 불일치로 잡힌다.
        extras = [_zero_if_none(_decimal(fields.get(n))) for n in spec.extras]
        computed = None if base is None else base + sum(extras, Decimal(0))
        parts = " + ".join(_show(v) for v in [base, *extras]) if base is not None else ""
        checks.append(_equal(name, computed, f"{parts} = {_show(computed)}" if parts else "", total))
    return checks


def _zero_if_none(value) -> Decimal:
    return Decimal(0) if value is None else value


def _item_sum(table) -> Decimal | None:
    amounts = [_decimal(row.get("금액")) for row in table or []]
    if not amounts or any(a is None for a in amounts):
        return None
    return sum(amounts, Decimal(0))


def _equal(item, expected, source_text, stated) -> Check:
    if expected is None or stated is None:
        return Check(item, source_text, _show(stated), None)
    return Check(item, source_text, _show(stated), expected == stated)


# --- 공통 ---

def _decimal(value) -> Decimal | None:
    """숫자로 볼 수 있는 값을 Decimal 로. float 은 repr 그대로(12.3 → 12.3). 못 바꾸면 None."""
    if value is None or isinstance(value, bool):
        return None
    try:
        if isinstance(value, Decimal):
            number = value
        elif isinstance(value, int):
            number = Decimal(value)
        elif isinstance(value, float):
            number = Decimal(repr(value))
        elif isinstance(value, str) and value.strip():
            number = Decimal(value.replace(",", "").replace("$", "").strip())
        else:
            return None
    except InvalidOperation:
        return None
    return number if number.is_finite() else None


def _show(value) -> str:
    if value is None:
        return ""
    if isinstance(value, Decimal):
        return format(value, "f")
    if isinstance(value, date):
        return value.isoformat()
    return str(value)
