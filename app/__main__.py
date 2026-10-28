"""명령 입구.

  uv run python -m app                받은문서/ 처리
  uv run python -m app --덮어쓰기      같은 번호 문서를 새 내용으로 바꿔 기록
  uv run python -m app init           작업 폴더 선택·생성 (설치.bat)
  uv run python -m app check <PDF>    추출 결과만 표로 보기 (파일 이동·장부 기록 없음)
"""
import argparse
import sys
from datetime import date
from pathlib import Path

from app import config
from app.extract import extract
from app.init import InitError, init
from app.organize import render_filename
from app.rules import RuleError, load_rules


def _parser():
    p = argparse.ArgumentParser(prog="python -m app", description="PDF → 엑셀 장부 사무 자동화")
    p.add_argument("command", nargs="?", choices=["init", "check"], help="init 또는 check (없으면 받은문서 처리)")
    p.add_argument("target", nargs="?", help="check 할 PDF 경로")
    p.add_argument("--덮어쓰기", dest="overwrite", action="store_true", help="이미 있는 문서를 새 내용으로 바꿈")
    return p


def main(argv=None, repo_root=config.REPO_ROOT) -> int:
    args = _parser().parse_args(argv)
    try:
        if args.command == "init":
            init(repo_root=repo_root)
            return 0
        work = config.work_dir(repo_root)
        rules = load_rules(work / config.RULES_FILE)
        if args.command == "check":
            if not args.target:
                print("검사할 PDF 경로를 적어 주세요. 예: uv run python -m app check 받은문서\\샘플.pdf")
                return 1
            return check(Path(args.target), rules)
        print("받은문서 처리는 아직 준비 중입니다.")
        return 1
    except (config.ConfigError, InitError) as e:
        print(e)
        return 1
    except RuleError as e:
        print(f"규칙.yaml 오류: {e}")
        return 1


def _show(value):
    if value is None:
        return "(비어 있음)"
    if isinstance(value, date):
        return value.isoformat()
    return str(value)


def _cell(text):
    return str(text).replace("|", "\\|")


def check(pdf: Path, rules) -> int:
    if not pdf.is_file():
        print(f"파일이 없습니다: {pdf}")
        return 1
    result = extract(pdf, rules)
    print(f"파일: {pdf.name}")
    print(f"문서종류: {result.doc_type.name if result.doc_type else '(판별 실패)'}")

    if result.review_reason and (result.doc_type is None or not result.has_text):
        print(f"\n확인필요: {result.review_reason}")
        if result.text_head:
            print("\n--- 1쪽 텍스트 앞 30줄 ---")
            print(result.text_head)
        return 0

    print(f"번호필드: {result.doc_type.number_field}")
    print("\n| 필드 | 값 | 비고 |\n|---|---|---|")
    for name, value in result.fields.items():
        notes = []
        if value is None:
            notes.append("못 찾음")
        notes += [w for w in result.warnings if f"'{name}' 값" in w]
        print(f"| {_cell(name)} | {_cell(_show(value))} | {_cell(', '.join(notes))} |")

    table = result.doc_type.table
    if table is not None:
        if result.table is None:
            print("\n품목 표: 읽지 못함 (아래 경고 참고)")
        else:
            print(f"\n품목 {len(result.table)}줄")
            print("| " + " | ".join(table.columns) + " |")
            print("|" + "---|" * len(table.columns))
            for row in result.table:
                print("| " + " | ".join(_cell(row.get(c, "")) for c in table.columns) + " |")

    others = [w for w in result.warnings if "' 못 찾음" not in w and "' 값 " not in w]
    for w in others:
        print(f"경고: {w}")
    if result.review_reason:
        print(f"\n확인필요: {result.review_reason}")
    else:
        print(f"\n정리될 경로: {config.ORGANIZED}\\{render_filename(result)}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
