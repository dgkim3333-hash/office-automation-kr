"""규칙.yaml 을 읽고, 쓰기 전에 구조를 검증한다."""
import re
import string
from dataclasses import dataclass
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
class DocType:
    name: str
    keywords: list[str]
    number_field: str
    fields: dict[str, Field]
    table: Table | None
    filename_tpl: str


@dataclass
class Rules:
    doc_types: list[DocType]

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

    return Rules([_doc_type(str(name), body) for name, body in data["문서종류"].items()])


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
        names = [n for _, n, _, _ in string.Formatter().parse(filename_tpl) if n is not None]
    except ValueError as e:
        raise RuleError(f"{where} 파일명 중괄호 오류: {e}") from None
    for n in names:
        if n not in fields:
            raise RuleError(f"{where} 파일명의 {{{n}}} 가 필드 목록에 없습니다.")

    return DocType(
        name=name,
        keywords=[str(k) for k in keywords],
        number_field=number_field,
        fields=fields,
        table=_table(where, body.get("표")),
        filename_tpl=filename_tpl,
    )


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
