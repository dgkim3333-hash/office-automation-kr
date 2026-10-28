"""규칙.yaml 을 읽고, 쓰기 전에 구조를 검증한다."""
import re
import string
from dataclasses import dataclass, field
from pathlib import Path

import yaml

FORMATS = ("날짜", "금액")


class RuleError(Exception):
    """규칙.yaml 이 잘못됐을 때. 메시지는 사용자에게 그대로 보여 준다."""


@dataclass
class Field:
    name: str
    pattern: re.Pattern
    fmt: str | None = None


@dataclass
class Table:
    start: str
    end: str
    row_pattern: re.Pattern
    columns: list[str]
    amount_columns: list[str]
    quantity_columns: list[str]


@dataclass
class CheckSpec:
    """합계 검산에 쓸 필드 이름(모두 그 문서종류의 '형식: 금액' 필드)과 줄 검산 · 품목 합에 쓸 품목 열 이름."""
    subtotal: str | None
    extras: list[str]
    total: str | None
    quantity: str = "수량"
    price: str = "단가"
    amount: str = "금액"
    named: frozenset = field(default=frozenset(), compare=False, repr=False)  # 규칙에 직접 적은 열 키


@dataclass
class DocType:
    name: str
    keywords: list[str]
    number_field: str
    fields: dict[str, Field]
    table: Table | None
    filename_tpl: str
    check: CheckSpec | None = None


@dataclass
class HeadCell:
    name: str
    where: str  # "B4" 또는 "오른쪽:글자"
    fmt: str | None = None


@dataclass
class ExcelItems:
    header_text: str           # 이 글자가 있는 줄 바로 아래부터 품목
    columns: dict[str, str]    # 열이름 → 엑셀 열 문자
    amount_columns: list[str]
    quantity_columns: list[str]


@dataclass
class ExcelForm:
    name: str
    doc_type: str
    keywords: list[str]
    head: dict[str, HeadCell]
    items: ExcelItems | None


@dataclass
class CompanyColumn:
    name: str
    source: str  # 머리.X · 품목.X · 고객마스터.X · 줄번호 · 수량x단가 · 고정:글자 · 빈칸


@dataclass
class CompanyForm:
    filename_tpl: str
    sheet: str
    columns: list[CompanyColumn]


@dataclass
class Rules:
    doc_types: list[DocType]  # 판별 순서
    sheet_order: list[str] = field(default_factory=list)  # 장부 탭 순서(판별 순서와 따로)
    excel_forms: list[ExcelForm] = field(default_factory=list)  # 고객별 엑셀 PO 읽는 법(판별 순서)
    company_form: CompanyForm | None = None  # 엑셀 PO 를 바꿔 넣을 회사 양식

    def get(self, name):
        return next((d for d in self.doc_types if d.name == name), None)


def load_rules(path) -> Rules:
    path = Path(path)
    try:
        data = yaml.safe_load(path.read_text(encoding="utf-8"))
    except FileNotFoundError:
        raise RuleError(f"규칙 파일이 없습니다: {path}") from None
    except yaml.YAMLError as e:
        raise RuleError(f"규칙 파일 문법 오류(들여쓰기·따옴표 확인): {e}") from None

    if not isinstance(data, dict) or not isinstance(data.get("문서종류"), dict) or not data["문서종류"]:
        raise RuleError("규칙 파일 맨 위에 '문서종류:' 와 그 아래 종류가 하나 이상 있어야 합니다.")

    doc_types = [_doc_type(str(name), body) for name, body in data["문서종류"].items()]
    by_name = {d.name: d for d in doc_types}
    excel = data.get("엑셀양식") or {}
    if not isinstance(excel, dict):
        raise RuleError("엑셀양식 아래에는 고객별 이름: 내용 을 적어야 합니다.")
    for name in excel:
        if str(name) in by_name:
            raise RuleError(
                f"엑셀양식 이름 '{name}' 이 문서종류 이름과 같습니다. 승인 수를 양식 이름으로 세므로 "
                f"엑셀양식 이름을 다르게 지으세요(예: 고객 이름).")
    return Rules(
        doc_types,
        _sheet_order(data.get("시트순서"), [d.name for d in doc_types]),
        [_excel_form(str(name), body, by_name) for name, body in excel.items()],
        _company_form(data.get("회사양식")),
    )


CELL = re.compile(r"^[A-Z]{1,3}[1-9][0-9]*$")
COLUMN = re.compile(r"^[A-Z]{1,3}$")
RIGHT_OF = "오른쪽:"
COMPANY_VALUE = re.compile(r"^((머리|품목|고객마스터)\.\S+|줄번호|수량x단가|빈칸|고정:.*)$", re.S)


def _excel_form(name, body, doc_types) -> ExcelForm:
    where = f"[엑셀양식 {name}]"
    if not isinstance(body, dict):
        raise RuleError(f"{where} 내용이 비었습니다.")
    doc_type = doc_types.get(str(body.get("문서종류")))
    if doc_type is None:
        raise RuleError(f"{where} 문서종류 '{body.get('문서종류')}' 가 문서종류 목록에 없습니다.")
    keywords = body.get("판별키워드")
    if not isinstance(keywords, list) or not keywords:
        raise RuleError(f"{where} 판별키워드를 하나 이상 적어야 합니다.")

    head = {}
    for key, spec in (body.get("머리") or {}).items():
        cell, fmt = (spec.get("칸"), spec.get("형식")) if isinstance(spec, dict) else (spec, None)
        cell = str(cell or "")
        if not (CELL.match(cell) or (cell.startswith(RIGHT_OF) and cell[len(RIGHT_OF):].strip())):
            raise RuleError(f"{where} 머리 '{key}' 의 칸 '{cell}' 은 'B4' 나 '오른쪽:글자' 처럼 적어야 합니다.")
        if fmt is not None and fmt not in FORMATS:
            raise RuleError(f"{where} 머리 '{key}' 형식 '{fmt}' 는 쓸 수 없습니다. 날짜 또는 금액만 됩니다.")
        head[str(key)] = HeadCell(str(key), cell, fmt)
    if doc_type.number_field not in head:
        raise RuleError(f"{where} 머리에 번호필드 '{doc_type.number_field}' 가 있어야 합니다.")
    items = _excel_items(where, body.get("품목"))
    _check_columns(f"{where} 문서종류 {doc_type.name}", doc_type.check, items and items.columns, "품목")
    return ExcelForm(name, doc_type.name, [str(k) for k in keywords], head, items)


def _excel_items(where, spec) -> ExcelItems | None:
    if spec is None:
        return None
    if not isinstance(spec, dict) or not spec.get("머리글"):
        raise RuleError(f"{where} 품목에는 머리글(품목 표 제목 줄에 있는 글자)이 있어야 합니다.")
    columns = {str(k): str(v) for k, v in (spec.get("열") or {}).items()}
    if not columns:
        raise RuleError(f"{where} 품목 열을 {{품번: B, 수량: D}} 처럼 적어야 합니다.")
    for key, letter in columns.items():
        if not COLUMN.match(letter):
            raise RuleError(f"{where} 품목 열 '{key}' 의 '{letter}' 는 엑셀 열 문자(A, B …)가 아닙니다.")
    typed = {}
    for key in ("금액열", "수량열"):
        names = [str(c) for c in spec.get(key) or []]
        for n in names:
            if n not in columns:
                raise RuleError(f"{where} 품목 {key} '{n}' 가 열 목록에 없습니다.")
        typed[key] = names
    return ExcelItems(str(spec["머리글"]), columns, typed["금액열"], typed["수량열"])


def _company_form(spec) -> CompanyForm | None:
    if spec is None:
        return None
    where = "[회사양식]"
    if not isinstance(spec, dict) or not spec.get("파일명") or not isinstance(spec.get("열"), list):
        raise RuleError(f"{where} 파일명과 열 목록이 있어야 합니다.")
    columns = []
    for col in spec["열"]:
        name = str((col or {}).get("이름") or "")
        source = str((col or {}).get("값") or "")
        if not name:
            raise RuleError(f"{where} 이름이 없는 열이 있습니다.")
        if not COMPANY_VALUE.match(source):
            raise RuleError(
                f"{where} 열 '{name}' 의 값 '{source}' 를 알 수 없습니다. "
                "머리.이름 · 품목.이름 · 고객마스터.열 · 줄번호 · 수량x단가 · 고정:글자 · 빈칸 중 하나로 적으세요."
            )
        columns.append(CompanyColumn(name, source))
    return CompanyForm(str(spec["파일명"]), str(spec.get("시트") or "Sheet1"), columns)


def _sheet_order(spec, names) -> list[str]:
    """장부 탭 순서. 적힌 종류 먼저, 빠진 종류는 판별 순서대로 뒤에 붙인다."""
    if spec is None:
        return list(names)
    if not isinstance(spec, list):
        raise RuleError("시트순서는 [Invoice, DM] 처럼 적어야 합니다.")
    order = [str(n) for n in spec]
    for n in order:
        if n not in names:
            raise RuleError(f"시트순서의 '{n}' 가 문서종류에 없습니다.")
    return order + [n for n in names if n not in order]


def _doc_type(name, body) -> DocType:
    where = f"[{name}]"
    if not isinstance(body, dict):
        raise RuleError(f"{where} 내용이 비었습니다.")

    keywords = body.get("판별키워드")
    if not isinstance(keywords, list) or not keywords:
        raise RuleError(f"{where} 판별키워드는 [\"글자1\", \"글자2\"] 처럼 하나 이상 적어야 합니다.")

    raw_fields = body.get("필드")
    if not isinstance(raw_fields, dict) or not raw_fields:
        raise RuleError(f"{where} 필드가 하나 이상 있어야 합니다.")
    fields = {str(f): _field(where, str(f), spec) for f, spec in raw_fields.items()}

    number_field = body.get("번호필드")
    if number_field not in fields:
        raise RuleError(f"{where} 번호필드 '{number_field}' 가 필드 목록에 없습니다.")

    filename_tpl = body.get("파일명")
    if not isinstance(filename_tpl, str) or not filename_tpl:
        raise RuleError(f"{where} 파일명이 없습니다.")
    try:
        parts = [(n, spec) for _, n, spec, _ in string.Formatter().parse(filename_tpl) if n is not None]
    except ValueError as e:
        raise RuleError(f"{where} 파일명 중괄호 오류: {e}") from None
    for n, spec in parts:
        if n not in fields:
            raise RuleError(f"{where} 파일명의 {{{n}}} 가 필드 목록에 없습니다.")
        if spec and fields[n].fmt != "날짜":
            raise RuleError(f"{where} 파일명의 {{{n}:{spec}}} 처럼 ':' 서식은 '형식: 날짜' 필드에만 쓸 수 있습니다.")

    table = _table(where, body.get("표"))
    if "검산" in body and body["검산"] is None:
        raise RuleError(f"{where} 검산에 적은 것이 없습니다. 소계 · 부대비용 · 합계 중 하나 이상을 적거나 검산 줄을 지우세요.")
    check = _check_spec(where, body.get("검산"), fields)
    _check_columns(where, check, table and table.columns, "표의")

    return DocType(
        name=name,
        keywords=[str(k) for k in keywords],
        number_field=number_field,
        fields=fields,
        table=table,
        filename_tpl=filename_tpl,
        check=check,
    )


CHECK_KEYS = ("소계", "부대비용", "합계")
CHECK_COLUMN_KEYS = (("수량열", "quantity"), ("단가열", "price"), ("금액열", "amount"))


def _check_columns(where, check, columns, what):
    """검산에 쓰는 품목 열이 열 목록에 있어야 한다.

    금액열은 소계 · 합계 검산이 있으면(기본 이름이어도) 꼭 있어야 한다. 수량열 · 단가열은 규칙에 직접 적었을 때만
    검사한다. 기본 이름이 표에 없으면 줄 검산이 해당 없음이 될 뿐이다(단가 없는 표도 합계 검산은 할 수 있게).
    """
    if check is None or columns is None:
        return
    for key, attr in CHECK_COLUMN_KEYS:
        name = getattr(check, attr)
        needed = key in check.named or (attr == "amount" and (check.subtotal or check.total))
        if needed and name not in columns:
            raise RuleError(
                f"{where} 검산의 {key} '{name}' 가 {what} 열 목록에 없습니다. "
                f"검산 아래에 {key}: 열이름 처럼 실제 열 이름을 적으세요."
            )


def _check_spec(where, spec, fields) -> CheckSpec | None:
    if spec is None:
        return None
    if not isinstance(spec, dict):
        raise RuleError(f"{where} 검산은 {{소계: 필드, 부대비용: [필드], 합계: 필드}} 처럼 적어야 합니다.")
    allowed = CHECK_KEYS + tuple(k for k, _ in CHECK_COLUMN_KEYS)
    for key in spec:
        if key not in allowed:
            raise RuleError(
                f"{where} 검산의 '{key}' 는 쓸 수 없습니다. 소계 · 부대비용 · 합계 · 수량열 · 단가열 · 금액열만 됩니다."
            )
    if not spec:
        raise RuleError(f"{where} 검산에 적은 것이 없습니다. 소계 · 부대비용 · 합계 중 하나 이상을 적거나 검산 줄을 지우세요.")
    for key, value in spec.items():
        if _blank(value):
            raise RuleError(f"{where} 검산에 적은 것이 없습니다. '{key}' 에 이름을 적거나 그 줄을 지우세요.")
    extras = spec.get("부대비용") or []
    if not isinstance(extras, list):
        extras = [extras]
    subtotal = spec.get("소계")
    total = spec.get("합계")
    names = [str(n) for n in extras] + [str(n) for n in (subtotal, total) if n is not None]
    for n in names:
        if n not in fields:
            raise RuleError(f"{where} 검산의 '{n}' 은 필드 목록에 없습니다.")
        if fields[n].fmt != "금액":
            raise RuleError(f"{where} 검산의 '{n}' 은 '형식: 금액' 필드여야 합니다.")
    if extras and total is None:
        raise RuleError(f"{where} 검산의 부대비용은 합계와 함께 적어야 합니다.")
    columns = {attr: str(spec.get(key, getattr(CheckSpec, attr))) for key, attr in CHECK_COLUMN_KEYS}
    return CheckSpec(
        subtotal=None if subtotal is None else str(subtotal),
        extras=[str(n) for n in extras],
        total=None if total is None else str(total),
        **columns,
        named=frozenset(k for k, _ in CHECK_COLUMN_KEYS if k in spec),
    )


def _blank(value) -> bool:
    if value is None:
        return True
    if isinstance(value, list):
        return not value or any(_blank(v) for v in value)
    return not str(value).strip()


def _field(where, name, spec) -> Field:
    if not isinstance(spec, dict) or not spec.get("패턴"):
        raise RuleError(f"{where} 필드 '{name}' 에 패턴이 없습니다.")
    try:
        pattern = re.compile(spec["패턴"])
    except re.error as e:
        raise RuleError(f"{where} 필드 '{name}' 정규식 오류: {e}") from None
    if pattern.groups < 1:
        raise RuleError(f"{where} 필드 '{name}' 패턴에 값을 감쌀 괄호 ( ) 가 없습니다.")
    fmt = spec.get("형식")
    if fmt is not None and fmt not in FORMATS:
        raise RuleError(f"{where} 필드 '{name}' 형식 '{fmt}' 는 쓸 수 없습니다. 날짜 또는 금액만 됩니다.")
    return Field(name=name, pattern=pattern, fmt=fmt)


def _table(where, spec) -> Table | None:
    if spec is None:
        return None
    if not isinstance(spec, dict):
        raise RuleError(f"{where} 표 설정이 잘못됐습니다.")
    if "여러단어열" in spec:
        raise RuleError(f"{where} 표의 여러단어열은 더 이상 쓰지 않습니다. 행패턴으로 줄을 나누세요.")
    start, end, columns = spec.get("시작키워드"), spec.get("끝키워드"), spec.get("열")
    if not start or not end:
        raise RuleError(f"{where} 표에는 시작키워드와 끝키워드가 모두 있어야 합니다.")
    if not isinstance(columns, list) or not columns:
        raise RuleError(f"{where} 표의 열은 [열1, 열2] 처럼 적어야 합니다.")
    columns = [str(c) for c in columns]
    if not spec.get("행패턴"):
        raise RuleError(f"{where} 표에 행패턴(품목 한 줄을 나누는 정규식)이 없습니다.")
    try:
        row_pattern = re.compile(spec["행패턴"])
    except re.error as e:
        raise RuleError(f"{where} 표 행패턴 정규식 오류: {e}") from None
    if row_pattern.groups != len(columns):
        raise RuleError(
            f"{where} 표 행패턴의 괄호 수({row_pattern.groups})와 열 수({len(columns)})가 다릅니다."
        )
    typed = {}
    for key in ("금액열", "수량열"):
        names = [str(c) for c in spec.get(key) or []]
        for name in names:
            if name not in columns:
                raise RuleError(f"{where} 표의 {key} '{name}' 가 열 목록에 없습니다.")
        typed[key] = names
    return Table(start=str(start), end=str(end), row_pattern=row_pattern, columns=columns,
                 amount_columns=typed["금액열"], quantity_columns=typed["수량열"])
