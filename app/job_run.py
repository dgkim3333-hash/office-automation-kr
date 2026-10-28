"""작업 하나 실행: 입력 폴더의 PDF · 엑셀을 읽어 결과(회사 양식 · 누적장부)를 만든다.

입력 폴더의 원본은 읽기만 한다(옮기기 · 이름 바꾸기 · 지우기 없음). 어떤 파일을 처리했는지는
작업 폴더의 처리기록.csv 에 파일 내용 지문(sha256)으로 남긴다. 이 파일은 줄을 더하기만 하며,
한 지문의 지금 상태는 그 지문의 마지막 줄이다.
"""
import csv
import hashlib
import os
import shutil
import sys
from contextlib import contextmanager
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path

import openpyxl
from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE

from app.excel_extract import read_excel
from app.extract import extract
from app.jobs import Job, _folder_title, list_jobs, save_job
from app.ledger import ITEMS_SHEET, Ledger, LedgerLocked, is_locked
from app.local_guard import LocalOnlyError, check_work_dir
from app.log import append_log
from app.master import MasterError, load_master, lookup
from app.organize import free_name
from app.outputs import _same_value, _text, company_rows, read_back, write_check_file, write_company_form
from app.rules import RuleError, load_rules
from app.verify import failed, verify

LEDGER_NAME = "누적장부_{제목}.xlsx"  # 출력 폴더에 둔다
RECORD_FILE = "처리기록.csv"  # 작업 폴더
RECORD_COLUMNS = ["지문", "파일", "문서종류", "번호", "상태", "규칙지문", "날짜"]
LOG_FILE = "처리로그.csv"  # 작업 폴더. 아무도 안 볼 때(자동 실행)를 위한 건별 기록
LOCK_FILE = "실행중.lock"
MASTER_FILE = "고객마스터.xlsx"  # 작업 폴더들의 위(work) 에 있는 공용 파일
WAITING = "승인대기"
WAITING_CHECKS = "대조표"
REVIEW = "확인필요"
CHECKS = "검산"

INPUT_SUFFIXES = (".pdf", ".xlsx")  # .xls 는 판별하지 못하므로 건너뛴다
UNMATCHED_SUFFIXES = (".pdf", ".xlsx", ".xls")
SETTLE_SECONDS = 60  # 이보다 최근에 바뀐 파일은 아직 저장 중일 수 있어 다음 실행으로 미룬다
UNMATCHED_SECONDS = 24 * 3600
COUNT_KEYS = ("성공", "중복", "확인필요", "승인대기", "보류")
NEWS_KEYS = ("성공", "중복", "확인필요", "승인대기")
LEDGER_MODES = ("누적장부", "둘다")
COMPANY_MODES = ("회사양식", "둘다")
MASTER_PREFIX = "고객마스터."

RETRY_NOTE = "규칙.yaml 을 고치면 다음 실행 때 다시 시도합니다. 규칙을 그대로 두어도 다음 날 한 번 더 시도합니다."
HELD_REASON = "다른 프로그램(PDF 뷰어 · 엑셀 등)에서 열려 있음. 닫으면 다음 실행 때 처리합니다"
SETTLING_REASON = "방금 바뀐 파일(1분 안). 저장이 끝난 뒤 다음 실행 때 처리합니다"
LEDGER_LOCKED_REASON = "누적장부가 엑셀에서 열려 있음. 엑셀을 닫으면 다음 실행 때 처리합니다"


class JobLocked(Exception):
    """같은 작업이 이미 실행 중일 때(자동 실행 중에 지금 실행을 누른 경우 등)."""


@dataclass
class RunReport:
    counts: dict[str, int] = field(default_factory=lambda: dict.fromkeys(COUNT_KEYS, 0))
    unmatched_old: int = 0
    message: str = ""
    output_files: list[Path] = field(default_factory=list)

    @property
    def has_news(self) -> bool:
        """알릴 것이 있나. 보류만 있으면 알리지 않는다(다음 실행 때 처리됨)."""
        return any(self.counts.get(k, 0) > 0 for k in NEWS_KEYS) or self.unmatched_old > 0


# --- 같은 작업 두 번 실행 막기 ---

def _lock(f):
    f.seek(0)
    if sys.platform == "win32":
        import msvcrt

        msvcrt.locking(f.fileno(), msvcrt.LK_NBLCK, 1)
    else:
        import fcntl

        fcntl.flock(f.fileno(), fcntl.LOCK_EX | fcntl.LOCK_NB)


def _unlock(f):
    f.seek(0)
    if sys.platform == "win32":
        import msvcrt

        msvcrt.locking(f.fileno(), msvcrt.LK_UNLCK, 1)
    else:
        import fcntl

        fcntl.flock(f.fileno(), fcntl.LOCK_UN)


@contextmanager
def job_lock(job: Job):
    """작업 폴더의 실행중.lock 을 잡는다. 이미 누가 잡고 있으면 JobLocked."""
    f = open(Path(job.folder) / LOCK_FILE, "a+b")
    try:
        try:
            _lock(f)
        except OSError:
            raise JobLocked(f"이 작업이 이미 실행 중입니다({job.name}). 끝난 뒤 다시 실행하세요.") from None
        try:
            yield
        finally:
            try:
                _unlock(f)
            except OSError:
                pass  # 파일을 닫으면 잠금도 풀린다
    finally:
        f.close()


# --- 지문 · 처리기록 ---

def file_fingerprint(path) -> str:
    """파일 내용의 sha256(16진수). 이름이 바뀌어도 같은 파일이면 같은 값."""
    digest = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def rules_fingerprint(job: Job) -> str:
    """규칙.yaml 바이트의 sha256 앞 16자. 규칙이 바뀌었는지 볼 때 쓴다."""
    return hashlib.sha256(Path(job.rules_path).read_bytes()).hexdigest()[:16]


def load_records(job: Job) -> list[dict]:
    path = Path(job.folder) / RECORD_FILE
    if not path.exists():
        return []
    with open(path, encoding="utf-8-sig", newline="") as f:
        return [dict(row) for row in csv.DictReader(f)]


def _append_record(job: Job, row: dict):
    path = Path(job.folder) / RECORD_FILE
    new = not path.exists() or path.stat().st_size == 0
    # 새 파일만 utf-8-sig(BOM 포함), 이어 쓸 때는 BOM 없이 utf-8
    with open(path, "a", encoding="utf-8-sig" if new else "utf-8", newline="") as f:
        writer = csv.writer(f)
        if new:
            writer.writerow(RECORD_COLUMNS)
        writer.writerow([row.get(c, "") for c in RECORD_COLUMNS])


def _current_state(records) -> dict[str, dict]:
    """지문 → 그 지문의 마지막 줄."""
    state = {}
    for row in records:
        state[row.get("지문", "")] = row
    return state


# --- 바깥에서 부르는 것 ---

def run_job(job: Job, work: Path, *, now=None, source_dir: Path | None = None, gate=None,
            on_success=None) -> RunReport:
    """작업 하나를 실행한다. source_dir 기본 = 입력 폴더(승인 실행은 작업 폴더의 승인대기\\ 를 넘긴다).

    gate(job, rules, result) 가 True 면 결과를 내지 않고 승인대기로 보낸다.
    on_success(job, result) 는 출력까지 끝난 건마다 부른다.
    """
    now = (now or datetime.now()).replace(microsecond=0)
    source = Path(source_dir) if source_dir is not None else Path(job.input_dir)
    with job_lock(job):
        return _Run(job, Path(work), now, source, gate, on_success).run()


def unmatched_old(job: Job, work: Path, now: datetime) -> int:
    """입력 폴더에 하루 넘게 남아 있고, 같은 입력 폴더를 쓰는 어느 작업의 처리기록에도 없는 파일 수."""
    folder = Path(job.input_dir)
    if not folder.is_dir():
        return 0
    jobs = [j for j in list_jobs(work) if _same_dir(j.input_dir, folder)]
    if not any(_same_dir(j.folder, job.folder) for j in jobs):
        jobs.append(job)
    known = set()
    for j in jobs:
        try:
            known |= {r.get("지문") for r in load_records(j)}
        except (OSError, UnicodeDecodeError, csv.Error):
            pass
    count = 0
    for path in folder.iterdir():
        if not (path.is_file() and path.suffix.lower() in UNMATCHED_SUFFIXES and not path.name.startswith("~$")):
            continue
        try:
            if now.timestamp() - path.stat().st_mtime <= UNMATCHED_SECONDS:
                continue
            if file_fingerprint(path) not in known:
                count += 1
        except OSError:
            continue
    return count


def _same_dir(a, b) -> bool:
    return os.path.normcase(os.path.abspath(a)) == os.path.normcase(os.path.abspath(b))


def ledger_path(job: Job) -> Path:
    return Path(job.output_dir) / LEDGER_NAME.format(제목=_folder_title(job.title))


# --- 저장한 누적장부 다시 읽기 ---

def _sheet_rows(ws):
    rows = [list(r) for r in ws.iter_rows(values_only=True)]
    return (rows[0] if rows else []), rows[1:]


def _row_problems(where, header, row, want: dict) -> list[str]:
    problems = []
    for name, value in want.items():
        if name not in header:
            continue  # 장부에 열이 없는 값은 쓰지 않았다
        got = row[header.index(name)] if header.index(name) < len(row) else None
        if isinstance(value, str):
            value = ILLEGAL_CHARACTERS_RE.sub("", value)
        if not _same_value(value, got):
            problems.append(f"{where} {name}: 쓴 값 {_text(value)}, 읽은 값 {_text(got)}")
    return problems


def _ledger_problems(path, result, doc_values: dict) -> list[str]:
    """저장한 누적장부에서 이번 문서의 줄(문서종류 시트 · 품목 시트)을 찾아 쓴 값과 같은지 본다."""
    doc = result.doc_type
    number = str(result.number)
    wb = openpyxl.load_workbook(Path(path), data_only=True)
    try:
        if doc.name not in wb.sheetnames or ITEMS_SHEET not in wb.sheetnames:
            return [f"누적장부에 '{doc.name}' 또는 '{ITEMS_SHEET}' 시트가 없습니다"]
        header, rows = _sheet_rows(wb[doc.name])
        if doc.number_field not in header:
            return [f"누적장부 {doc.name} 시트에 '{doc.number_field}' 열이 없습니다"]
        col = header.index(doc.number_field)
        mine = [r for r in rows if col < len(r) and r[col] is not None and str(r[col]) == number]
        if len(mine) != 1:
            return [f"누적장부 {doc.name} 시트에 번호 {number} 줄이 {len(mine)}개입니다"]
        problems = _row_problems(f"누적장부 {doc.name}", header, mine[0], doc_values)

        header, rows = _sheet_rows(wb[ITEMS_SHEET])
        kind, num = header.index("문서종류"), header.index("문서번호")
        mine = [r for r in rows if r[kind] == doc.name and r[num] is not None and str(r[num]) == number]
        want_items = [{"문서종류": doc.name, "문서번호": result.number, **row} for row in result.table or []]
        if len(mine) != len(want_items):
            return problems + [f"누적장부 품목 시트에 {number} 줄이 {len(mine)}개입니다(쓴 줄 {len(want_items)}개)"]
        for i, (got, want) in enumerate(zip(mine, want_items)):
            problems += _row_problems(f"누적장부 품목 {i + 1}줄", header, got, want)
        return problems
    finally:
        wb.close()


# --- 실행 한 번 ---

class _Run:
    def __init__(self, job, work, now, source, gate, on_success):
        self.job = job
        self.work = work
        self.now = now
        self.today = now.date().isoformat()
        self.source = source
        self.gate = gate
        self.on_success = on_success
        self.report = RunReport()
        self.warnings: list[str] = []
        self.ledger = None
        self.ledger_locked = False
        self.master = None
        # 승인 실행(승인대기\ 를 읽음)이면 승인대기 기록을 건너뛰지 않는다
        self.approving = _same_dir(source, Path(job.folder) / WAITING)

    # 준비

    def _preflight(self) -> str | None:
        job = self.job
        for label, path in (("입력", self.source), ("출력", job.output_dir)):
            try:
                check_work_dir(path)
            except LocalOnlyError as e:
                return f"{label} 폴더를 쓸 수 없습니다. {e}"
            if not Path(path).is_dir():
                return f"{label} 폴더가 없습니다: {path}"
        if is_locked(Path(job.folder) / RECORD_FILE):
            return (f"{RECORD_FILE} 가 다른 프로그램(엑셀 등)에서 열려 있습니다. 닫고 다시 실행하세요.\n"
                    "(아무 파일도 처리하지 않았습니다.)")
        return None

    def run(self) -> RunReport:
        stop = self._preflight()
        if stop:
            return RunReport(message=stop)
        try:
            self.rules = load_rules(self.job.rules_path)
            self.rules_fp = rules_fingerprint(self.job)
        except (RuleError, OSError) as e:
            return RunReport(message=f"규칙.yaml 오류: {e}")
        if self.job.output_mode in COMPANY_MODES and self.rules.company_form is None:
            return RunReport(message=(
                f"출력 방식이 '{self.job.output_mode}' 인데 규칙.yaml 에 회사양식이 없습니다. "
                "규칙.yaml 에 회사양식을 적거나 출력 방식을 '누적장부' 로 바꾸세요."))
        try:
            self.state = _current_state(load_records(self.job))
        except (OSError, UnicodeDecodeError, csv.Error) as e:
            return RunReport(message=f"{RECORD_FILE} 를 읽지 못했습니다: {e}")
        self.success_keys = {(r.get("문서종류"), r.get("번호")) for r in self.state.values() if r.get("상태") == "성공"}

        for path in self._candidates():
            self._one(path)

        self.job.last_run = self.today
        try:
            save_job(self.job)
        except OSError as e:
            self.warnings.append(f"경고: 작업설정.yaml 에 마지막 실행 날짜를 쓰지 못함 ({e})")
        self.report.unmatched_old = unmatched_old(self.job, self.work, self.now)
        self.report.message = self._message()
        return self.report

    def _candidates(self):
        return sorted(p for p in self.source.iterdir()
                      if p.is_file() and p.suffix.lower() in INPUT_SUFFIXES and not p.name.startswith("~$"))

    # 한 건

    def _one(self, path: Path):
        try:
            age = self.now.timestamp() - path.stat().st_mtime
        except OSError:
            return
        if age < SETTLE_SECONDS:
            self._done(path, "보류", SETTLING_REASON)
            return
        if os.access(path, os.W_OK) and is_locked(path):  # 읽기 전용 파일은 잠김 검사가 늘 실패하므로 건너뜀
            self._done(path, "보류", HELD_REASON)
            return
        try:
            fp = file_fingerprint(path)
        except OSError:
            self._done(path, "보류", HELD_REASON)
            return
        if self._skip(fp):
            return
        try:
            result = read_excel(path, self.rules) if path.suffix.lower() == ".xlsx" else extract(path, self.rules)
        except Exception as e:  # 깨진 파일 한 건이 나머지를 막지 않게
            self._review(path, fp, None, f"파일을 읽지 못함: {e}")
            return
        if result.doc_type is None and result.has_text:
            return  # 판별 실패: 다른 작업 몫일 수 있으므로 기록하지 않고 그대로 둔다
        if result.review_reason:
            self._review(path, fp, result, result.review_reason)
            return
        key = (result.doc_type.name, str(result.number))
        if key in self.success_keys:
            self._record(path, fp, result, "중복")
            self._done(path, "중복", f"{key[0]} {key[1]} 이미 처리됨")
            return
        try:
            self._process(path, fp, result, key)
        except Exception as e:  # 한 건의 오류가 나머지를 막지 않게
            self._review(path, fp, result, f"처리 중 오류: {e}")

    def _skip(self, fp) -> bool:
        row = self.state.get(fp)
        if row is None:
            return False
        status = row.get("상태")
        if status in ("성공", "중복"):
            return True
        if status == "승인대기":
            return not self.approving
        if status == "확인필요":
            return row.get("날짜") == self.today and row.get("규칙지문") == self.rules_fp
        return False

    def _process(self, path, fp, result, key):
        checks = verify(result, path)
        bad = failed(checks)
        if bad:
            self._review(path, fp, result, "검산 불일치: " + "; ".join(
                f"{c.item} (원본 {c.source}, 결과 {c.output})" for c in bad))
            return
        master_row = None
        form = self.rules.company_form
        if self.job.output_mode in COMPANY_MODES and any(c.source.startswith(MASTER_PREFIX) for c in form.columns):
            customer = result.fields.get("고객명")
            try:
                master_row = lookup(self._master(), customer)
            except MasterError as e:
                self._review(path, fp, result, str(e))
                return
            if master_row is None:
                self._review(path, fp, result, (
                    f"고객마스터에 '{customer}' 가 없습니다. 고객마스터.xlsx 에 이 고객 줄을 더한 뒤 다시 넣으세요"))
                return
        if self.gate is not None and self.gate(self.job, self.rules, result):
            self._to_waiting(path, fp, result, checks)
            return
        self._write(path, fp, result, key, checks, master_row)

    def _master(self):
        if self.master is None:
            self.master = load_master(self.work / MASTER_FILE)
        return self.master

    def _to_waiting(self, path, fp, result, checks):
        waiting = Path(self.job.folder) / WAITING
        waiting.mkdir(parents=True, exist_ok=True)
        copy = free_name(waiting / path.name)
        try:
            shutil.copy2(path, copy)
            write_check_file(waiting / WAITING_CHECKS, result, checks, path.name)
        except Exception as e:
            if copy.exists():
                copy.unlink()
            self._review(path, fp, result, f"승인대기로 보내지 못함: {e}")
            return
        self._record(path, fp, result, "승인대기")
        self._done(path, "승인대기", "")

    def _open_ledger(self):
        if self.ledger is None:
            self.ledger = Ledger.open(ledger_path(self.job), self.rules)
        return self.ledger

    def _write(self, path, fp, result, key, checks, master_row):
        """출력: 누적장부 · 회사 양식. 저장한 뒤 다시 읽어 다르면 이번 건만 되돌리고 확인필요."""
        mode = self.job.output_mode
        if mode in LEDGER_MODES and self.ledger_locked:
            self._done(path, "보류", LEDGER_LOCKED_REASON)
            return
        ledger, added, saved, company = None, False, False, None
        problems = []
        try:
            if mode in LEDGER_MODES:
                ledger = self._open_ledger()
                if ledger.is_duplicate(*key):
                    self._record(path, fp, result, "중복")
                    self._done(path, "중복", f"{key[0]} {key[1]} 이미 누적장부에 있음")
                    return
                doc_values = {"처리일시": self.now, "원본파일": path.name, "정리된경로": str(path), **result.fields}
                ledger.add(result, path.name, path, processed_at=self.now)
                added = True
                ledger.save()
                saved = True
                try:
                    problems += _ledger_problems(ledger.path, result, doc_values)
                except Exception as e:
                    problems.append(f"저장한 누적장부를 다시 읽지 못함: {e}")
            if mode in COMPANY_MODES:
                form = self.rules.company_form
                company = write_company_form(self.job.output_dir, result, master_row, form)
                header, rows = company_rows(result, master_row, form)
                try:
                    problems += read_back(company, form.sheet, header, rows)
                except Exception as e:
                    problems.append(f"저장한 회사 양식 파일을 다시 읽지 못함: {e}")
        except LedgerLocked:
            self.ledger_locked = True
            self._rollback(ledger, added, saved, key, company)
            self._done(path, "보류", LEDGER_LOCKED_REASON)
            return
        except Exception as e:
            undo = self._rollback(ledger, added, saved, key, company)
            self._review(path, fp, result, f"결과를 쓰는 중 오류: {e}{undo}")
            return
        if problems:
            undo = self._rollback(ledger, added, saved, key, company)
            shown = "; ".join(problems[:5]) + (f" 외 {len(problems) - 5}건" if len(problems) > 5 else "")
            self._review(path, fp, result, f"저장한 파일을 다시 읽어 보니 쓴 값과 다릅니다: {shown}{undo}")
            return

        try:
            write_check_file(Path(self.job.folder) / CHECKS, result, checks, path.name)
        except Exception as e:  # 결과는 이미 나갔으므로 알리기만 한다
            self.warnings.append(f"경고: 검산 파일을 만들지 못함 ({path.name}): {e}")
        self._record(path, fp, result, "성공")
        self.success_keys.add(key)
        if company is not None:
            self.report.output_files.append(company)
        if ledger is not None and ledger.path not in self.report.output_files:
            self.report.output_files.append(ledger.path)
        self._done(path, "성공", "; ".join(result.warnings))
        if self.on_success is not None:
            try:
                self.on_success(self.job, result)
            except Exception as e:
                self.warnings.append(f"경고: 성공 처리 뒤 작업을 하지 못함 ({path.name}): {e}")

    def _rollback(self, ledger, added, saved, key, company) -> str:
        """이번 건이 쓴 것을 되돌린다. 되돌리지 못한 것이 있으면 사유에 붙일 글자를 돌려준다."""
        notes = []
        if added:
            ledger.remove(*key)
            if saved:
                try:
                    ledger.save()
                except Exception as e:
                    notes.append(f"누적장부에서 이 문서 줄을 지우지 못함({e}). 누적장부를 열어 직접 지우세요")
        if company is not None and company.exists():
            try:
                company.unlink()
            except OSError as e:
                notes.append(f"회사 양식 파일을 지우지 못함({company.name}: {e}). 직접 지우세요")
        return "".join(f". {n}" for n in notes)

    # 기록

    def _record(self, path, fp, result, status):
        row = {
            "지문": fp,
            "파일": path.name,
            "문서종류": result.doc_type.name if result is not None and result.doc_type else "",
            "번호": str(result.number) if result is not None and result.number is not None else "",
            "상태": status,
            "규칙지문": self.rules_fp,
            "날짜": self.today,
        }
        try:
            _append_record(self.job, row)
        except OSError as e:
            self.warnings.append(f"경고: {RECORD_FILE} 에 쓰지 못함 ({path.name}, {status}): {e}")
        self.state[fp] = row

    def _review(self, path, fp, result, reason):
        folder = Path(self.job.folder) / REVIEW
        try:
            folder.mkdir(parents=True, exist_ok=True)
            note = f"원본: {path}\n사유: {reason}\n\n{RETRY_NOTE}\n"
            (folder / f"{path.stem}.txt").write_text(note, encoding="utf-8")
        except OSError as e:
            self.warnings.append(f"경고: 확인필요 사유 파일을 쓰지 못함 ({path.name}): {e}")
        self._record(path, fp, result, "확인필요")
        self._done(path, "확인필요", reason)

    def _done(self, path, outcome, reason):
        self.report.counts[outcome] += 1
        try:
            append_log(Path(self.job.folder) / LOG_FILE, self.now, path.name, outcome, reason)
        except OSError as e:
            self.warnings.append(f"경고: {LOG_FILE} 에 쓰지 못함 ({e})")

    def _message(self) -> str:
        counts = self.report.counts
        parts = [f"{k} {counts[k]}" for k in NEWS_KEYS]
        if counts["보류"]:
            parts.append(f"보류 {counts['보류']}")
        lines = [f"{self.job.title}: " + " · ".join(parts)]
        folder = Path(self.job.folder)
        if counts["확인필요"]:
            lines.append(f"확인필요 폴더의 .txt 를 열어 사유를 보세요: {folder / REVIEW}")
        if counts["승인대기"]:
            lines.append(f"승인대기 폴더의 대조표를 확인한 뒤 승인하세요: {folder / WAITING / WAITING_CHECKS}")
        if counts["보류"]:
            lines.append("보류 파일은 입력 폴더에 그대로 있고 다음 실행 때 처리합니다.")
        if self.ledger_locked:
            lines.append(f"누적장부가 엑셀에서 열려 있습니다. 엑셀을 닫아 두세요: {ledger_path(self.job)}")
        if self.report.unmatched_old:
            lines.append(f"어느 작업에도 맞지 않는 파일 {self.report.unmatched_old}개가 입력 폴더에 하루 넘게 "
                         f"남아 있습니다: {self.job.input_dir}")
        return "\n".join(lines + self.warnings)
