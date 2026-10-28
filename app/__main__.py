"""명령 입구.

  uv run python -m app                받은문서/ 처리
  uv run python -m app --덮어쓰기      같은 번호 문서를 새 내용으로 바꿔 기록
  uv run python -m app init [경로]    작업 폴더 선택·생성 (설치.bat). 경로를 주면 묻지 않음
  uv run python -m app check <PDF>    추출 결과만 표로 보기 (파일 이동·장부 기록 없음)
  uv run python -m app outlook        (선택) Outlook '자동화' 폴더 첨부 → 받은문서 (수집.bat)
"""
import argparse
import sys
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path

from app import config
from app.excel_extract import read_excel
from app.extract import ExtractResult, extract
from app.init import InitError, init
from app.ledger import Ledger, LedgerLocked, is_locked
from app.local_guard import BANNER, LocalOnlyError, block_network, check_environment, check_work_dir
from app.log import append_log, summary
from app.master import MasterError, load_master, lookup
from app.outlook import collect_from_outlook
from app.outputs import write_company_form, write_extract_result
from app.organize import (
    free_name, move_duplicate, move_review, move_success, render_filename, render_text, target_path,
)
from app.rules import RuleError, load_rules

LOCKED_MESSAGE = (
    "누적장부.xlsx 가 엑셀에서 열려 있습니다. 엑셀을 닫고 다시 실행하세요.\n"
    "(파일은 하나도 옮기지 않았습니다. 받은문서에 그대로 있습니다.)"
)
LOG_LOCKED_MESSAGE = (
    "처리로그.csv 가 다른 프로그램(엑셀 등)에서 열려 있습니다. 닫고 다시 실행하세요.\n"
    "(파일은 하나도 옮기지 않았습니다.)"
)
HELD_REASON = "다른 프로그램(PDF 뷰어·엑셀 등)에서 열려 있어 옮길 수 없음. 닫고 다시 실행하세요"
INPUT_SUFFIXES = (".pdf", ".xlsx", ".xls")
XLS_REASON = "옛 엑셀 형식(.xls)은 읽지 못합니다. 엑셀에서 열어 「다른 이름으로 저장」 → .xlsx 로 저장한 뒤 받은문서에 다시 넣으세요"


def _now():
    return datetime.now().replace(microsecond=0)


@dataclass
class _Item:
    """받은문서 파일(PDF·엑셀) 한 건의 처리 계획."""
    src: Path
    result: ExtractResult
    outcome: str = ""      # 성공 · 중복 · 확인필요 · 보류 · 이동 실패
    reason: str = ""
    dst: Path | None = None  # 성공일 때 옮길 곳
    added: bool = False      # 이번에 장부에 새로 적었나(출력파일을 만들 대상)
    master_row: dict | None = None  # 엑셀 PO 의 고객마스터 줄


def run(work: Path, rules, overwrite=False) -> int:
    """받은문서 처리. 순서: 전부 추출 → 장부 저장 1회 → 저장이 된 뒤에만 파일 이동."""
    inbox = work / config.INBOX
    log_file = work / config.LOG_FILE
    pdfs = sorted(p for p in inbox.iterdir() if p.is_file() and p.suffix.lower() in INPUT_SUFFIXES
                  and not p.name.startswith("~$"))  # ~$ 는 엑셀이 열어 둔 임시 파일
    if not pdfs:
        print(f"받은문서에 PDF·엑셀 파일이 없습니다: {inbox}")
        print(summary([]))
        return 0
    if is_locked(log_file):
        print(LOG_LOCKED_MESSAGE)
        return 2

    items = []
    for pdf in pdfs:
        if is_locked(pdf):  # PDF 뷰어 등이 잡고 있으면 옮길 수 없으므로 장부에도 넣지 않는다
            items.append(_Item(pdf, ExtractResult(doc_type=None), "보류", HELD_REASON))
            continue
        if pdf.suffix.lower() == ".xls":
            items.append(_Item(pdf, ExtractResult(doc_type=None), "확인필요", XLS_REASON))
            continue
        try:
            result = read_excel(pdf, rules) if pdf.suffix.lower() == ".xlsx" else extract(pdf, rules)
        except Exception as e:  # 깨진 파일 한 건이 나머지를 막지 않게
            result = ExtractResult(doc_type=None, has_text=False)
            items.append(_Item(pdf, result, "확인필요", f"파일을 읽지 못함: {e}"))
            continue
        items.append(_Item(pdf, result))

    when = _now()
    try:
        ledger = Ledger.open(work / config.LEDGER_FILE, rules)
        _plan(items, ledger, work, overwrite, when, rules)
        for item in items:
            ledger.add_log(when, item.src.name, item.outcome, item.reason)
        ledger.save()
    except LedgerLocked:
        for item in items:
            _log(log_file, when, item.src.name, "보류", "장부가 엑셀에서 열려 있음")
        print(LOCKED_MESSAGE)
        return 2

    for item in items:
        _move(item, work)
        _log(log_file, when, item.src.name, item.outcome, item.reason)
    for item in items:
        if item.outcome == "성공" and item.added:
            _write_output(item, work, rules)

    print()
    print(summary(i.outcome for i in items))
    if any(i.outcome == "확인필요" for i in items):
        print(f"확인필요 폴더의 .txt 를 열어 사유를 보세요: {work / config.REVIEW}")
    if any(i.outcome in ("보류", "이동 실패") for i in items):
        print("보류·이동 실패 파일은 받은문서에 그대로 있습니다. 열린 프로그램을 닫고 다시 실행하세요.")
    return 0


def _log(log_file, when, name, outcome, reason):
    try:
        append_log(log_file, when, name, outcome, reason)
    except OSError as e:
        print(f"경고: 처리로그.csv 에 쓰지 못함 ({e}). 장부의 처리로그 시트에는 기록됨.")


def _master_loader(work):
    """고객마스터.xlsx 를 처음 필요할 때 한 번만 읽는다."""
    cache = {}

    def load():
        if "master" not in cache:
            cache["master"] = load_master(work / config.MASTER_FILE)
        return cache["master"]

    return load


def _plan(items, ledger, work, overwrite, when, rules):
    """각 건을 성공·중복·확인필요로 나누고 장부에 기록한다(아직 저장·이동은 안 함)."""
    seen, reserved = set(), set()
    master = _master_loader(work)
    for item in items:
        if item.outcome:  # 이미 정해짐(보류·읽기 실패)
            continue
        result = item.result
        if result.review_reason:
            item.outcome, item.reason = "확인필요", result.review_reason
            continue
        key = (result.doc_type.name, str(result.number))
        added = False
        try:
            if key in seen or (ledger.is_duplicate(*key) and not overwrite):
                _plan_duplicate(item, ledger, work, key, seen)
                continue
            if result.source_form and rules.company_form is not None:  # 엑셀 PO: 회사 양식에 고객마스터가 필요
                customer = result.fields.get("고객명")
                try:
                    item.master_row = lookup(master(), customer)
                except MasterError as e:
                    item.outcome, item.reason = "확인필요", str(e)
                    continue
                if item.master_row is None:
                    item.outcome, item.reason = "확인필요", (
                        f"고객마스터에 '{customer}' 가 없습니다. 고객마스터.xlsx 에 이 고객 줄을 더한 뒤 다시 넣으세요")
                    continue
            notes = []
            if ledger.is_duplicate(*key):
                ledger.remove(*key)
                notes.append("덮어쓰기")
            item.dst = free_name(target_path(result, work).with_suffix(item.src.suffix.lower()), reserved)
            ledger.add(result, item.src.name, item.dst.relative_to(work), processed_at=when)
            added = item.added = True
            seen.add(key)
            reserved.add(item.dst)
            item.outcome, item.reason = "성공", "; ".join(notes + result.warnings)
        except Exception as e:  # 한 건의 오류가 나머지를 막지 않게: 기록을 되돌리고 확인필요로
            if added or not ledger.is_duplicate(*key):
                ledger.remove(*key)
            item.dst = None
            item.added = False
            item.outcome, item.reason = "확인필요", f"처리 중 오류: {e}"


def _plan_duplicate(item, ledger, work, key, seen):
    recorded = ledger.organized_path(*key)
    if key not in seen and recorded and not (work / recorded).exists():
        # 지난번 이동이 실패해 장부에만 있는 문서: 중복으로 보내지 않고 장부에 적힌 자리로 옮긴다
        item.dst = work / recorded
        item.outcome, item.reason = "성공", "장부에는 있으나 정리문서에 파일이 없어 그 자리로 옮김"
        seen.add(key)
        return
    item.outcome, item.reason = "중복", f"{key[0]} {key[1]} 이미 장부에 있음"


def _write_output(item, work, rules):
    """문서별 결과 엑셀(출력파일). 실패해도 장부·이동은 이미 끝났으므로 알리기만 한다."""
    folder = work / config.OUTPUT
    try:
        if item.result.source_form and rules.company_form is not None:
            path = write_company_form(folder, item.result, item.master_row, rules.company_form)
        else:
            path = write_extract_result(folder, item.result, item.src.name)
        print(f"출력      {path.name}")
    except Exception as e:
        print(f"경고: 출력파일을 만들지 못함 ({item.src.name}): {e}")


def _move(item, work):
    try:
        if item.outcome == "성공":
            final = move_success(item.src, item.dst)
            print(f"성공      {item.src.name} -> {final.relative_to(work)}")
        elif item.outcome == "중복":
            move_duplicate(item.src, work)
            print(f"중복      {item.src.name} -> {config.INBOX}\\{config.DUPLICATES}\\")
        elif item.outcome == "확인필요":
            move_review(item.src, work, item.reason, item.result.text_head)
            print(f"확인필요  {item.src.name} : {item.reason}")
        else:
            print(f"보류      {item.src.name} : {item.reason}")
    except OSError as e:
        recorded = " (장부에는 있음)" if item.outcome == "성공" else ""
        item.outcome, item.reason = "이동 실패", f"이동 실패{recorded}: {e}"
        print(f"이동 실패 {item.src.name} : 받은문서에 그대로 둠{recorded}")

def _parser():
    p = argparse.ArgumentParser(prog="python -m app", description="PDF → 엑셀 장부 사무 자동화")
    p.add_argument("command", nargs="?", choices=["init", "check", "outlook"],
                   help="init · check · outlook (없으면 받은문서 처리)")
    p.add_argument("target", nargs="?", help="check 할 PDF 경로, 또는 init 할 작업 폴더 경로")
    p.add_argument("--덮어쓰기", dest="overwrite", action="store_true", help="이미 있는 문서를 새 내용으로 바꿈")
    return p


def _utf8_console():
    """출력이 파이프여도(Claude Code 가 실행할 때) 한글이 깨지거나 죽지 않게 utf-8 로 고정."""
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8", errors="replace")
        except (AttributeError, ValueError):
            pass


def main(argv=None, repo_root=config.REPO_ROOT) -> int:
    _utf8_console()
    args = _parser().parse_args(argv)
    try:
        check_environment()
    except LocalOnlyError as e:
        print(e)
        return 1
    print(BANNER)
    print()
    restore_network = block_network()
    try:
        return _dispatch(args, repo_root)
    finally:
        restore_network()


def _dispatch(args, repo_root) -> int:
    try:
        if args.command == "init":
            # 경로를 주면 묻지 않고 그 자리에 만든다(설치 스크립트용)
            init(work_dir=args.target, repo_root=repo_root)
            return 0
        work = config.work_dir(repo_root)
        check_work_dir(work)
        if args.command == "outlook":
            collect_from_outlook(work / config.INBOX)
            return 0
        rules = load_rules(work / config.RULES_FILE)
        if args.command == "check":
            if not args.target:
                print("검사할 PDF 경로를 적어 주세요. 예: uv run python -m app check 받은문서\\샘플.pdf")
                return 1
            return check(Path(args.target), rules, work)
        return run(work, rules, overwrite=args.overwrite)
    except (config.ConfigError, InitError, LocalOnlyError) as e:
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


def check(pdf: Path, rules, work: Path | None = None) -> int:
    if not pdf.is_file():
        print(f"파일이 없습니다: {pdf}")
        return 1
    if pdf.suffix.lower() == ".xls":
        print(XLS_REASON)
        return 0
    is_excel = pdf.suffix.lower() == ".xlsx"
    result = read_excel(pdf, rules) if is_excel else extract(pdf, rules)
    print(f"파일: {pdf.name}")
    print(f"문서종류: {result.doc_type.name if result.doc_type else '(판별 실패)'}")
    if result.source_form:
        print(f"엑셀양식: {result.source_form}")

    if result.review_reason and (result.doc_type is None or not result.has_text):
        print(f"\n확인필요: {result.review_reason}")
        if result.text_head:
            print("\n--- 앞부분 글자(30줄) ---")
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

    columns = _item_columns(result, rules)
    if columns:
        if result.table is None:
            print("\n품목 표: 읽지 못함 (아래 경고 참고)")
        else:
            print(f"\n품목 {len(result.table)}줄")
            print("| " + " | ".join(columns) + " |")
            print("|" + "---|" * len(columns))
            for row in result.table:
                print("| " + " | ".join(_cell(_show(row.get(c, ""))) for c in columns) + " |")

    others = [w for w in result.warnings if "' 못 찾음" not in w and "' 값 " not in w]
    for w in others:
        print(f"경고: {w}")
    if result.review_reason:
        print(f"\n확인필요: {result.review_reason}")
        return 0
    organized = render_filename(result).with_suffix(pdf.suffix.lower())
    print(f"\n정리될 경로: {config.ORGANIZED}\\{organized}")
    if result.source_form and rules.company_form is not None:
        _check_company_form(result, rules, work)
    return 0


def _item_columns(result, rules):
    if result.source_form:
        form = next(f for f in rules.excel_forms if f.name == result.source_form)
        return list(form.items.columns) if form.items else []
    return result.doc_type.table.columns if result.doc_type.table else []


def _check_company_form(result, rules, work):
    """엑셀 PO: 고객마스터에서 고객을 찾았는지, 회사 양식이 몇 줄로 만들어질지."""
    customer = result.fields.get("고객명")
    row = None
    try:
        row = lookup(load_master(work / config.MASTER_FILE), customer) if work else None
    except MasterError as e:
        print(f"고객마스터: {e}")
    if row is not None:
        print(f"고객마스터: 찾음 (고객코드 {row.get('고객코드')})")
    else:
        print(f"고객마스터: 없음 ('{customer}' 줄을 고객마스터.xlsx 에 더해야 회사 양식이 만들어집니다)")
    name = render_text(rules.company_form.filename_tpl, result.fields)
    print(f"회사 양식: {name} ({len(result.table or [])}줄)")


if __name__ == "__main__":
    sys.exit(main())
