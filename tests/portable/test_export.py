"""무설치판용 규칙 내보내기: 규칙.yaml → JSON(무설치판이 ConvertFrom-Json 으로 읽음)."""
import json
import os
import re
import shutil
import subprocess
import time
from pathlib import Path

import pytest

from app import config, local_guard
from app.__main__ import main
from app.job_run import rules_fingerprint
from app.jobs import find_job
from app.portable import (
    ONEDRIVE_ENV, BundleError, _check_regex, export_rules, make_bundle, rules_to_portable,
)
from app.rules import RuleError, rules_from_text

from .conftest import decode_output

REPO = Path(__file__).resolve().parent.parent.parent
EXAMPLE_RULES = REPO / "rules" / "규칙.예시.yaml"


def _example_text() -> str:
    return EXAMPLE_RULES.read_text(encoding="utf-8")


def _without_master(text: str) -> str:
    """예시 규칙에서 고객마스터.* 열 줄을 뺀다(무설치판은 고객마스터를 지원하지 않음)."""
    return "\n".join(line for line in text.splitlines() if "값: 고객마스터." not in line) + "\n"


def _example_rules():
    return rules_from_text(_without_master(_example_text()))


def test_example_rules_export_round_trip():
    """예시 규칙(고객마스터 열 제외)을 내보내면 JSON 이 되고, 문서종류 · 회사양식 열 순서가 그대로다."""
    rules = _example_rules()
    data = rules_to_portable(rules)

    assert json.loads(json.dumps(data, ensure_ascii=False)) == data
    assert [d["name"] for d in data["doc_types"]] == ["DM", "Invoice", "PO"]
    assert data["sheet_order"] == ["Invoice", "DM", "PO"]
    assert [c["name"] for c in data["company_form"]["columns"]] == [c.name for c in rules.company_form.columns]
    assert data["version"] == 1

    invoice = next(d for d in data["doc_types"] if d["name"] == "Invoice")
    assert invoice["number_field"] == "인보이스번호"
    assert invoice["filename"] == "{고객명}/{발행일:%Y-%m}/Invoice_{인보이스번호}.pdf"
    assert {"name": "발행일", "pattern": r"Date[.:]?\s*(\d{1,2}/\d{1,2}/\d{4})", "format": "날짜"} in invoice["fields"]
    assert invoice["table"]["columns"] == ["줄", "품번", "품명", "수량", "단가", "금액"]
    assert invoice["table"]["amount_columns"] == ["단가", "금액"]
    assert invoice["table"]["quantity_columns"] == ["수량"]
    assert invoice["check"] == {"subtotal": "소계", "extras": ["운임"], "total": "합계금액",
                                "quantity_column": "수량", "price_column": "단가", "amount_column": "금액"}
    po = next(d for d in data["doc_types"] if d["name"] == "PO")
    assert po["table"] is None
    assert next(f for f in po["fields"] if f["name"] == "고객PO번호")["format"] is None

    acme = next(f for f in data["excel_forms"] if f["name"] == "ACME")
    assert acme["doc_type"] == "PO" and acme["keywords"] == ["Buyer:"]
    assert {"name": "발주일", "where": "오른쪽:PO Date:", "format": "날짜"} in acme["head"]
    assert acme["items"]["header"] == "Supplier Part No"
    assert acme["items"]["columns"]["품번"] == "B"
    assert acme["items"]["amount_columns"] == ["단가", "금액"]
    assert data["company_form"]["filename"] == "회사양식_SO입력_{고객PO번호}.xlsx"
    assert data["company_form"]["sheet"] == "SO_Upload"


def test_master_columns_rejected():
    """고객마스터.* 값은 무설치판이 지원하지 않으므로 내보내기에서 거절한다."""
    rules = rules_from_text(_example_text())
    with pytest.raises(RuleError) as e:
        rules_to_portable(rules)
    assert "고객마스터" in str(e.value)
    assert "무설치판" in str(e.value)


@pytest.mark.parametrize("token, bad", [
    ("(?P<", r"(?P<x>\d+)"),
    ("(?P=", r"(?P<x>\d+)(?P=x)"),
    ("\\Z", r"(\d+)\Z"),
])
def test_python_only_regex_rejected(token, bad):
    """.NET 과 문법이 다른 파이썬 전용 정규식은 필드 패턴에서 거절하고, 메시지에 그 글자가 들어 있다."""
    rules = _example_rules()
    rules.doc_types[0].fields["DM번호"].pattern = re.compile(bad)
    with pytest.raises(RuleError) as e:
        rules_to_portable(rules)
    assert token in str(e.value)


def test_named_backreference_alone_rejected():
    """(?P= 는 이름 붙은 묶음 없이도 목록에 있어야 하므로 검사 함수에 글자만 직접 넣어 본다."""
    with pytest.raises(RuleError) as e:
        _check_regex("[시험]", r"(a)(?P=x)")
    assert "(?P=" in str(e.value)


def test_python_only_regex_rejected_in_table_row_pattern():
    """표 행패턴도 같은 목록으로 거절한다."""
    rules = _example_rules()
    table = rules.get("Invoice").table
    table.row_pattern = re.compile(table.row_pattern.pattern[:-1] + r"\Z")
    with pytest.raises(RuleError) as e:
        rules_to_portable(rules)
    assert "\\Z" in str(e.value)


def test_markup_columns_exported():
    """마크업.* 열은 지원한다. 마크업 열 이름은 기본값(품번 · 단가 · 수량)으로 내보낸다."""
    rules = _example_rules()
    assert any(c.source == "마크업.단가" for c in rules.company_form.columns)
    data = rules_to_portable(rules)
    assert data["company_form"]["markup_columns"] == {"part": "품번", "price": "단가", "quantity": "수량"}
    assert {"name": "HQ Unit Price", "value": "마크업.단가"} in data["company_form"]["columns"]


def test_markup_column_names_customized():
    """회사양식.마크업 으로 바꾼 품목 열 이름이 그대로 나간다."""
    text = _without_master(_example_text()).replace(
        "  시트: SO_Upload\n", "  시트: SO_Upload\n  마크업: {품번열: 품명, 수량열: 수량}\n", 1)
    data = rules_to_portable(rules_from_text(text))
    assert data["company_form"]["markup_columns"] == {"part": "품명", "price": "단가", "quantity": "수량"}


def test_no_company_form_exports_null():
    """회사양식이 없으면 company_form 은 null, 엑셀양식이 없으면 빈 목록이다."""
    text = _example_text().split("# 고객이 보낸 엑셀 PO 읽는 법")[0]
    data = rules_to_portable(rules_from_text(text))
    assert data["company_form"] is None
    assert data["excel_forms"] == []


@pytest.fixture
def env(tmp_path, capsys):
    """가짜 저장소(rules 만) + init 한 작업 폴더."""
    repo = tmp_path / "repo"
    shutil.copytree(REPO / "rules", repo / "rules")
    work = tmp_path / "work"
    assert main(["init", str(work)], repo_root=repo) == 0
    (tmp_path / "in").mkdir()
    (tmp_path / "out").mkdir()
    capsys.readouterr()
    return repo, work.resolve(), tmp_path


def _new_job(repo, tmp_path, title, mode, rules_text):
    assert main(["job", "new", "--제목", title, "--입력", str(tmp_path / "in"),
                 "--출력", str(tmp_path / "out"), "--방식", mode], repo_root=repo) == 0
    job = find_job(config.work_dir(repo), title)
    job.rules_path.write_text(rules_text, encoding="utf-8")
    return job


def _make_job(repo, tmp_path, rules_text):
    return _new_job(repo, tmp_path, "시험 작업", "누적장부", rules_text)


def test_cli_export_writes_json(env, capsys):
    """export-portable 01 --out FILE: JSON 파일이 생기고 fingerprint 가 job_run.rules_fingerprint 와 같다."""
    repo, work, tmp_path = env
    job = _make_job(repo, tmp_path, _without_master(_example_text()))
    out = tmp_path / "내보냄" / "규칙.json"
    capsys.readouterr()

    assert main(["export-portable", f"{job.number:02d}", "--out", str(out)], repo_root=repo) == 0

    data = json.loads(out.read_text(encoding="utf-8"))
    assert data["fingerprint"] == rules_fingerprint(job)
    assert [d["name"] for d in data["doc_types"]] == ["DM", "Invoice", "PO"]
    assert not out.read_bytes().startswith(b"\xef\xbb\xbf")  # BOM 없는 utf-8
    assert str(out) in capsys.readouterr().out


def test_cli_export_rejects_master_columns(env, capsys):
    """고객마스터 열이 있는 작업은 파일을 만들지 않고 오류를 보여 준다."""
    repo, work, tmp_path = env
    job = _make_job(repo, tmp_path, _example_text())
    out = tmp_path / "규칙.json"
    capsys.readouterr()

    assert main(["export-portable", f"{job.number:02d}", "--out", str(out)], repo_root=repo) == 1

    shown = capsys.readouterr().out
    assert "고객마스터" in shown and "무설치판" in shown
    assert not out.exists()


def test_export_rules_function(env):
    """export_rules(번호, 경로, 작업폴더) 는 쓴 경로를 돌려준다."""
    repo, work, tmp_path = env
    job = _make_job(repo, tmp_path, _without_master(_example_text()))
    out = tmp_path / "규칙.json"
    assert export_rules(job.number, out, work) == out
    assert json.loads(out.read_text(encoding="utf-8"))["version"] == 1


# --- 묶음 만들기(bundle-portable): 친구 PC 로 넘길 폴더 하나 ---

BUNDLE_SAMPLES = (Path("samples") / "03_Invoice_DM_PDF" / "INV-2026-10231.pdf",
                  Path("samples") / "01_고객오더_엑셀" / "고객오더_ACME_PO-ACME-2026-0417.xlsx")
RESULT_KEYS = {"doc_type", "form", "is_excel", "fields", "table", "warnings", "review_reason"}


def _portable_code_files(root: Path) -> list[str]:
    """저장소 무설치판\\ 의 코드(.ps1 · .bat) 상대 경로."""
    return sorted(p.relative_to(root).as_posix() for p in (root / "무설치판").rglob("*")
                  if p.is_file() and p.suffix.lower() in (".ps1", ".bat"))


def _bundle_files(out: Path) -> set[str]:
    return {p.relative_to(out).as_posix() for p in out.rglob("*") if p.is_file()}


@pytest.fixture
def bundle_env(env):
    """env + 가짜 저장소에 무설치판\\ 과 자체 시험용 가짜 샘플 두 개(tmp_path 안의 복사본)."""
    repo, work, tmp_path = env
    shutil.copytree(REPO / "무설치판", repo / "무설치판")
    for sample in BUNDLE_SAMPLES:
        (repo / sample).parent.mkdir(parents=True, exist_ok=True)
        shutil.copyfile(REPO / sample, repo / sample)
    return env


def _plant_private_files(tmp_path: Path, jobs):
    """묶음에 들어가면 안 되는 것들: 입력 폴더 파일 · 처리기록 · 처리로그 · 누적장부 · 확인필요 · 검산 · 규칙.yaml.bak."""
    (tmp_path / "in" / "메일_Order ACME-PO-0001.pdf").write_bytes(b"%PDF-1.4 fake")
    (tmp_path / "in" / "고객오더_ACME.xlsx").write_bytes(b"fake")
    for job in jobs:
        (job.folder / "처리기록.csv").write_text("fake\n", encoding="utf-8")
        (job.folder / "처리로그.csv").write_text("fake\n", encoding="utf-8")
        (job.folder / "규칙.yaml.bak").write_text("fake\n", encoding="utf-8")
        for sub in ("확인필요", "검산", "승인대기"):
            (job.folder / sub).mkdir(exist_ok=True)
            (job.folder / sub / "INV-0001.pdf").write_bytes(b"%PDF-1.4 fake")
    (tmp_path / "out" / "누적장부_ACME.xlsx").write_bytes(b"fake")


def test_bundle_layout(bundle_env):
    """묶음 = 무설치판\\(코드) + 작업\\NN_제목\\규칙.json · 작업.json + 자체시험\\(샘플 PDF · 엑셀 · 규칙.json · 정답.json).
    고객마스터 · 마크업표 · 입력 폴더 파일 · 처리기록 · 확인필요 · 검산 · 다른 PDF · xlsx 는 없다."""
    repo, work, tmp_path = bundle_env
    text = _without_master(_example_text())
    first = _new_job(repo, tmp_path, "시험 작업", "누적장부", text)
    second = _new_job(repo, tmp_path, "두번째 작업", "둘다", text)
    _plant_private_files(tmp_path, [first, second])
    (repo / "무설치판" / "메모.txt").write_text("코드 아님\n", encoding="utf-8")  # 무설치판\ 은 코드만 옮긴다
    assert (work / "고객마스터.xlsx").exists() and (work / "마크업표.xlsx").exists()
    out = tmp_path / "묶음"

    assert make_bundle(work, [1, 2], out, repo_root=repo) == out

    expected = set(_portable_code_files(repo))
    for job in (first, second):
        expected |= {f"작업/{job.name}/규칙.json", f"작업/{job.name}/작업.json"}
    expected |= {f"자체시험/{s.name}" for s in BUNDLE_SAMPLES} | {"자체시험/규칙.json", "자체시험/정답.json"}
    files = _bundle_files(out)
    assert files == expected
    assert sorted(p.name for p in out.iterdir()) == sorted(["무설치판", "작업", "자체시험"])

    for job in (first, second):
        info = json.loads((out / "작업" / job.name / "작업.json").read_text(encoding="utf-8"))
        assert info == {"number": job.number, "title": job.title, "output_mode": job.output_mode}
        rules = json.loads((out / "작업" / job.name / "규칙.json").read_text(encoding="utf-8"))
        assert rules == {**rules_to_portable(_example_rules()), "fingerprint": rules_fingerprint(job)}
    assert json.loads((out / "작업" / second.name / "작업.json").read_text(encoding="utf-8"))["output_mode"] == "둘다"

    answer = json.loads((out / "자체시험" / "정답.json").read_text(encoding="utf-8"))
    assert answer["pdf"]["file"] == "INV-2026-10231.pdf" and answer["excel"]["file"] == BUNDLE_SAMPLES[1].name
    for key in ("pdf", "excel"):
        assert set(answer[key]["result"]) == RESULT_KEYS
    assert answer["pdf"]["result"]["doc_type"] == "Invoice" and answer["excel"]["result"]["is_excel"] is True

    allowed = {f"자체시험/{s.name}" for s in BUNDLE_SAMPLES}
    for name in files:
        assert not any(word in name for word in ("고객마스터", "마크업표", "처리기록", "처리로그", "누적장부",
                                                 "확인필요", "검산", "승인대기", "메일_Order", ".bak"))
        if name.lower().endswith((".pdf", ".xlsx", ".xls", ".csv")):
            assert name in allowed


def test_bundle_copies_portable_code_bytes(bundle_env):
    """무설치판\\ 의 .ps1 · .bat 은 바이트 그대로(UTF-8 BOM, .bat 은 CRLF) 들어간다. 저장소 진짜 파일과 견준다."""
    repo, work, tmp_path = bundle_env
    _new_job(repo, tmp_path, "시험 작업", "누적장부", _without_master(_example_text()))
    out = tmp_path / "묶음"
    make_bundle(work, [1], out, repo_root=REPO)

    names = _portable_code_files(REPO)
    assert "무설치판/준비.bat" in names and "무설치판/lib/Setup.ps1" in names
    assert {n for n in _bundle_files(out) if n.startswith("무설치판/")} == set(names)
    for name in names:
        data = (out / name).read_bytes()
        assert data == (REPO / name).read_bytes(), name
        assert data.startswith(b"\xef\xbb\xbf"), name
        if name.endswith(".bat"):
            assert data.count(b"\n") == data.count(b"\r\n") > 0, name


def test_bundle_rules_written_fresh(bundle_env):
    """준비는 같은 번호면 규칙.json 수정 시각이 새것을 고르므로, 묶음의 규칙.json · 작업.json 은 묶을 때 새로 쓴다."""
    repo, work, tmp_path = bundle_env
    job = _new_job(repo, tmp_path, "시험 작업", "누적장부", _without_master(_example_text()))
    os.utime(job.rules_path, (946684800, 946684800))  # 2000-01-01
    before = time.time() - 2
    out = tmp_path / "묶음"
    make_bundle(work, [1], out, repo_root=repo)
    for name in ("규칙.json", "작업.json"):
        assert (out / "작업" / job.name / name).stat().st_mtime >= before


@pytest.mark.parametrize("case", ["repo", "repo_itself", "onedrive_segment", "onedrive_env", "onedrive_commercial_env",
                                  "onedrive_consumer_env", "unc", "network_drive"])
def test_bundle_refuses_repo_and_onedrive(bundle_env, monkeypatch, case):
    """묶음 폴더가 저장소 안 · OneDrive(환경 변수 경로 또는 OneDrive 로 시작하는 폴더) · 네트워크면 아무것도 만들지 않고 거부."""
    repo, work, tmp_path = bundle_env
    _new_job(repo, tmp_path, "시험 작업", "누적장부", _without_master(_example_text()))
    for name in ONEDRIVE_ENV:
        monkeypatch.delenv(name, raising=False)
    if case == "repo":
        out, words = repo / "묶음", ["저장소 안"]
    elif case == "repo_itself":
        out, words = repo, ["저장소 안"]
    elif case == "onedrive_segment":
        out, words = tmp_path / "OneDrive - ACME" / "묶음", ["OneDrive - ACME", "동기화"]
    elif case.endswith("_env"):
        synced = tmp_path / "회사 동기화"
        name = {"onedrive_env": "OneDrive", "onedrive_commercial_env": "OneDriveCommercial",
                "onedrive_consumer_env": "OneDriveConsumer"}[case]
        monkeypatch.setenv(name, str(synced))
        out, words = synced / "묶음", ["OneDrive", str(synced)]
    elif case == "unc":
        out, words = Path(r"\\fileserver\share\묶음"), ["네트워크 위치"]
    else:
        monkeypatch.setattr(local_guard, "_drive_type", lambda root: local_guard.DRIVE_REMOTE)
        out, words = tmp_path / "묶음", ["네트워크 드라이브"]
    before = _bundle_files(repo)

    with pytest.raises(BundleError) as e:
        make_bundle(work, [1], out, repo_root=repo)

    for word in words:
        assert word in str(e.value)
    assert "묶음 폴더" in str(e.value)
    assert _bundle_files(repo) == before
    if case not in ("repo_itself", "unc"):
        assert not out.exists()


def test_bundle_refuses_real_repo(bundle_env):
    """진짜 저장소 안(기본 repo_root)도 거부한다. 저장소에 아무것도 만들지 않는다."""
    repo, work, tmp_path = bundle_env
    _new_job(repo, tmp_path, "시험 작업", "누적장부", _without_master(_example_text()))
    out = REPO / "묶음_시험_지우세요"
    with pytest.raises(BundleError, match="저장소 안"):
        make_bundle(work, [1], out)
    assert not out.exists()


@pytest.mark.parametrize("stale", ["무설치판", "작업", "자체시험"])
def test_bundle_refuses_existing_bundle(bundle_env, stale):
    """묶음 폴더에 지난 묶음(무설치판 · 작업 · 자체시험)이 있으면 섞지 않고 거부하며, 있던 것은 그대로 둔다."""
    repo, work, tmp_path = bundle_env
    _new_job(repo, tmp_path, "시험 작업", "누적장부", _without_master(_example_text()))
    out = tmp_path / "묶음"
    (out / stale).mkdir(parents=True)
    (out / stale / "옛것.txt").write_text("옛 묶음\n", encoding="utf-8")

    with pytest.raises(BundleError) as e:
        make_bundle(work, [1], out, repo_root=repo)

    assert stale in str(e.value) and "새 폴더" in str(e.value)
    assert _bundle_files(out) == {f"{stale}/옛것.txt"}


def test_bundle_refuses_work_folder_as_out(bundle_env):
    """작업 폴더 자체를 묶음 폴더로 주면(작업\\ 이 있음) 거부하고 작업은 그대로다."""
    repo, work, tmp_path = bundle_env
    job = _new_job(repo, tmp_path, "시험 작업", "누적장부", _without_master(_example_text()))
    before = _bundle_files(work)
    with pytest.raises(BundleError, match="작업"):
        make_bundle(work, [1], work, repo_root=repo)
    assert _bundle_files(work) == before and job.rules_path.exists()


def test_bundle_refuses_unsupported_job_before_writing(bundle_env):
    """한 작업이라도 무설치판이 지원하지 않는 규칙(고객마스터 열)이면 묶음 폴더를 만들지 않는다."""
    repo, work, tmp_path = bundle_env
    _new_job(repo, tmp_path, "시험 작업", "누적장부", _without_master(_example_text()))
    _new_job(repo, tmp_path, "마스터 작업", "누적장부", _example_text())
    out = tmp_path / "묶음"
    with pytest.raises(RuleError, match="고객마스터"):
        make_bundle(work, [1, 2], out, repo_root=repo)
    assert not out.exists()


def test_bundle_without_jobs_refused(bundle_env):
    """묶을 작업이 없으면 거부한다."""
    repo, work, tmp_path = bundle_env
    with pytest.raises(BundleError, match="작업이 없습니다"):
        make_bundle(work, [], tmp_path / "묶음", repo_root=repo)
    assert not (tmp_path / "묶음").exists()


def test_bundle_setup_runs_on_friend_side(bundle_env, powershell_exe):
    """만든 묶음을 친구 PC 처럼(작업 폴더 = 묶음, 묶음 안의 무설치판.ps1) 준비하면 작업설정.json 이 묶음의 제목 ·
    출력 방식으로 생긴다. 자체 시험은 끼우기로 통과(오피스 없이), 예약 · 바탕화면 · 폴더 고르기도 가짜."""
    repo, work, tmp_path = bundle_env
    job = _new_job(repo, tmp_path, "시험 작업", "둘다", _without_master(_example_text()))
    out = make_bundle(work, [1], tmp_path / "묶음", repo_root=repo)
    (out / "마크업표.xlsx").write_bytes(b"x")  # 친구 PC 에 이미 있다고 침(없으면 준비가 엑셀로 예시를 만든다)
    friend_in, friend_out = tmp_path / "친구입력", tmp_path / "친구출력"
    for folder in (friend_in, friend_out, tmp_path / "desktop", tmp_path / "startup"):
        folder.mkdir()
    hooks = tmp_path / "hooks.json"
    hooks.write_text(json.dumps({"schedule": "fail", "desktop": str(tmp_path / "desktop"),
                                 "startup_dir": str(tmp_path / "startup"), "notice_file": str(tmp_path / "알림.jsonl"),
                                 "self_test": "pass", "pick": [str(friend_in), str(friend_out)]}, ensure_ascii=False),
                     encoding="utf-8")

    done = subprocess.run([powershell_exe, "-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass",
                           "-File", str(out / "무설치판" / "무설치판.ps1"), "-Command", "준비", "-Work", str(out),
                           "-NoDialog", "-TestHooks", str(hooks), "-TestPdfWarningOff", "1"],
                          capture_output=True, timeout=600)
    text = decode_output(done.stdout) + decode_output(done.stderr)

    assert done.returncode == 0, text
    data = json.loads((out / "작업" / job.name / "작업설정.json").read_text(encoding="utf-8-sig"))
    assert data["number"] == 1 and data["title"] == "시험 작업" and data["output_mode"] == "둘다"
    assert data["input_dir"] == str(friend_in) and data["output_dir"] == str(friend_out)


def test_cli_bundle(bundle_env, capsys):
    """bundle-portable --out DIR [--작업 NN …]: 고른 작업만(멈춘 작업도 고를 수 있음), 생략하면 멈추지 않은 모든 작업.
    친구 PC 로 복사할 것을 알려 준다."""
    repo, work, tmp_path = bundle_env
    text = _without_master(_example_text())
    first = _new_job(repo, tmp_path, "시험 작업", "누적장부", text)
    second = _new_job(repo, tmp_path, "두번째 작업", "회사양식", text)
    third = _new_job(repo, tmp_path, "세번째 작업", "둘다", text)
    for job in (first, third):  # job new 는 멈춤으로 만든다. 둘째만 멈춘 채로 둔다
        assert main(["job", "resume", f"{job.number:02d}"], repo_root=repo) == 0
    capsys.readouterr()

    out = tmp_path / "묶음1"
    assert main(["bundle-portable", "--out", str(out), "--작업", "02"], repo_root=repo) == 0
    shown = capsys.readouterr().out
    assert str(out) in shown and second.name in shown and "준비.bat" in shown
    assert sorted(p.name for p in (out / "작업").iterdir()) == [second.name]

    out = tmp_path / "묶음2"
    assert main(["bundle-portable", "--out", str(out)], repo_root=repo) == 0
    assert sorted(p.name for p in (out / "작업").iterdir()) == sorted([first.name, third.name])


def test_cli_bundle_all_paused(bundle_env, capsys):
    """--작업 을 생략했는데 멈추지 않은 작업이 없으면 만들지 않고 --작업 으로 고르라고 알려 준다."""
    repo, work, tmp_path = bundle_env
    _new_job(repo, tmp_path, "시험 작업", "누적장부", _without_master(_example_text()))
    capsys.readouterr()
    assert main(["bundle-portable", "--out", str(tmp_path / "묶음")], repo_root=repo) == 1
    assert "--작업" in capsys.readouterr().out
    assert not (tmp_path / "묶음").exists()


def test_cli_bundle_errors(bundle_env, capsys):
    """잘못된 묶음 폴더 · 없는 작업 · 숫자가 아닌 번호 · 지원하지 않는 규칙은 한국어 오류와 반환 1."""
    repo, work, tmp_path = bundle_env
    _new_job(repo, tmp_path, "시험 작업", "누적장부", _without_master(_example_text()))
    _new_job(repo, tmp_path, "마스터 작업", "누적장부", _example_text())
    capsys.readouterr()

    assert main(["bundle-portable", "--out", str(repo / "묶음")], repo_root=repo) == 1
    assert "저장소 안" in capsys.readouterr().out
    assert not (repo / "묶음").exists()

    out = tmp_path / "묶음"
    assert main(["bundle-portable", "--out", str(out), "--작업", "07"], repo_root=repo) == 1
    assert "찾지 못했습니다" in capsys.readouterr().out
    assert main(["bundle-portable", "--out", str(out), "--작업", "가"], repo_root=repo) == 1
    assert "숫자" in capsys.readouterr().out
    assert main(["bundle-portable", "--out", str(out), "--작업", "01", "02"], repo_root=repo) == 1
    shown = capsys.readouterr().out
    assert "고객마스터" in shown and "02_마스터 작업" in shown
    assert not out.exists()
