"""PDF 텍스트 추출 → 문서종류 판별 → 필드·표 추출. 파일을 옮기거나 쓰지 않는다."""
from pathlib import Path

import pdfplumber


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
