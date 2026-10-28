"""양식(승인 수를 세는 단위): 엑셀 PO 는 엑셀양식 이름, PDF 는 문서종류 이름. 그 양식의 규칙 지문."""
import hashlib
from pathlib import Path

import yaml


def form_key(result) -> str:
    """승인 수를 세는 양식 이름. 엑셀이면 엑셀양식 이름, 아니면 문서종류 이름."""
    if result.source_form:
        return str(result.source_form)
    return result.doc_type.name


def _named(section, key: str):
    if not isinstance(section, dict):
        return None
    for name, body in section.items():
        if str(name) == key:
            return body
    return None


def _detection(section) -> list:
    """판별에 쓰는 부분: 위에서부터의 순서와 각 이름의 판별키워드. 다른 양식의 판별키워드를 고쳐도
    이 양식으로 판별되는 문서가 달라질 수 있으므로 모든 양식의 지문에 넣는다."""
    if not isinstance(section, dict):
        return []
    return [[str(name), body.get("판별키워드") if isinstance(body, dict) else None] for name, body in section.items()]


def form_fingerprint(rules_path: Path, key: str) -> str:
    """규칙.yaml 중 그 양식에 해당하는 부분 + 회사양식 + 판별 순서 · 판별키워드의 sha256 앞 16자.

    엑셀양식이면 그 엑셀양식과 연결된 문서종류 + 모든 엑셀양식의 판별 부분,
    아니면 그 문서종류 + 모든 문서종류의 판별 부분. 다른 양식의 추출 규칙을 고쳐도 값이 바뀌지 않는다.
    """
    return form_fingerprint_data(yaml.safe_load(Path(rules_path).read_text(encoding="utf-8")), key)


def rules_form_fingerprint(rules, key: str) -> str:
    """이미 읽은 규칙(Rules.raw)에서 계산한 양식 지문. 실행 중 규칙.yaml 이 바뀌어도 추출에 쓴 규칙과 같다."""
    return form_fingerprint_data(rules.raw, key)


def form_fingerprint_data(data, key: str) -> str:
    """form_fingerprint 와 같되 규칙.yaml 을 yaml 로 읽은 값에서 계산한다."""
    data = data if isinstance(data, dict) else {}
    part = {"회사양식": data.get("회사양식")}
    excel = _named(data.get("엑셀양식"), key)
    if excel is not None:
        part["엑셀양식"] = excel
        doc = excel.get("문서종류") if isinstance(excel, dict) else None
        part["문서종류"] = _named(data.get("문서종류"), str(doc))
        part["판별"] = _detection(data.get("엑셀양식"))
    else:
        part["문서종류"] = _named(data.get("문서종류"), key)
        part["판별"] = _detection(data.get("문서종류"))
    text = yaml.safe_dump(part, sort_keys=True)
    return hashlib.sha256(text.encode("utf-8")).hexdigest()[:16]
