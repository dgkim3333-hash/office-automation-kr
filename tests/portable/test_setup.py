"""무설치판 준비 · 예약 · 자체 시험(작업 10).

시험은 진짜 작업 스케줄러 · 바탕화면 · 시작프로그램 폴더를 건드리지 않는다: 예약 등록은 늘 가짜(-TestHooks 의 schedule),
바로가기는 tmp_path 의 가짜 바탕화면 · 시작프로그램 폴더에 만든다. 폴더 고르기 창 · 메시지 창은 띄우지 않는다(-NoDialog, 고를 폴더는 hooks).
엑셀 · 워드가 필요한 시험은 @office(시험마다 /Automation 오피스 · %TEMP%\\무설치판_* 이 남지 않았는지 본다).
"""
import contextlib
import io
import json
import msvcrt
import os
import re
import shutil
import subprocess
import threading
import time
from datetime import datetime
from pathlib import Path

import openpyxl
import pytest

from app.__main__ import check
from app.local_guard import LocalOnlyError, check_work_dir
from app.markup import EXAMPLE_ROWS, load_markup
from app.portable import SELF_TEST_ANSWER, SELF_TEST_DIR, SELF_TEST_RULES, build_self_test, rules_to_portable
from app.rules import rules_from_text

from .conftest import DISPATCHER, REPO, decode_output, example_rules, without_master
from .test_office_read import no_leftover, office, temp_dir  # noqa: F401 (픽스처)
from .test_run import ACME, INVOICE, PO_RULES, Env, clean_temp, drop, example_text, files, split_row_po  # noqa: F401 (픽스처)

PORTABLE = REPO / "무설치판"
RUN_BAT = PORTABLE / "지금실행.bat"
SETUP_BAT = PORTABLE / "준비.bat"
TIMEOUT = 600
PDF_DIFFERENT = "이 PC 의 워드는 PDF 를 다르게 읽습니다. PDF 작업은 켜지 않았습니다"
TELL_BOSS = "대표에게 알려 주세요"
PDF_GUIDE = "워드에서 PDF 를 한 번 열고 '이 메시지를 다시 표시 안 함'에 체크하세요. 그 뒤 준비.bat 을 다시 두 번 클릭하면"


# --- 도우미 ---

def powershell(*args: str, timeout: int = TIMEOUT) -> tuple[int, str]:
    """무설치판.ps1 을 돌려 (반환 코드, 화면 글자)."""
    done = subprocess.run(["powershell.exe", "-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass",
                           "-File", str(DISPATCHER), *args], capture_output=True, timeout=timeout)
    return done.returncode, decode_output(done.stdout) + decode_output(done.stderr)


def write_hooks(root: Path, **kw) -> Path:
    """시험용 끼우기. 예약 등록은 늘 가짜, 바탕화면 · 시작프로그램은 tmp 폴더(빠뜨리면 무설치판이 멈춘다)."""
    hooks = {"schedule": "fail", "desktop": str(root / "desktop"), "startup_dir": str(root / "startup"),
             "notice_file": str(root / "알림.jsonl"), "self_test": "pass"}
    hooks.update(kw)
    (root / "desktop").mkdir(exist_ok=True)
    (root / "startup").mkdir(exist_ok=True)
    path = root / f"hooks_{time.monotonic_ns()}.json"
    path.write_text(json.dumps({k: v for k, v in hooks.items() if v is not None}, ensure_ascii=False), encoding="utf-8")
    return path


def notices(root: Path) -> list[dict]:
    path = root / "알림.jsonl"
    if not path.exists():
        return []
    return [json.loads(line) for line in path.read_text(encoding="utf-8-sig").splitlines() if line.strip()]


def make_bundle(root: Path, rules_text: str, number: int = 1, title: str = "일일 정리", mode: str = "누적장부") -> Path:
    """대표가 넘기는 묶음 모양(작업 11): 작업\\NN_제목\\규칙.json · 작업.json."""
    folder = root / "작업" / f"{number:02d}_{title}"
    folder.mkdir(parents=True, exist_ok=True)
    data = rules_to_portable(rules_from_text(rules_text))
    data["fingerprint"] = "f" * 16
    (folder / "규칙.json").write_text(json.dumps(data, ensure_ascii=False), encoding="utf-8")
    (folder / "작업.json").write_text(json.dumps({"number": number, "title": title, "output_mode": mode},
                                                 ensure_ascii=False), encoding="utf-8")
    return folder


def run_setup(bundle: Path, work: Path, hooks: Path, warning_off: str = "1") -> tuple[int, str]:
    return powershell("-Command", "준비", "-Bundle", str(bundle), "-Work", str(work), "-NoDialog",
                      "-TestHooks", str(hooks), "-TestPdfWarningOff", warning_off)


def settings(folder: Path) -> dict:
    return json.loads((folder / "작업설정.json").read_text(encoding="utf-8-sig"))


def read_shortcut(run_engine, path: Path) -> dict:
    return run_engine("setup", {"action": "read_shortcut", "path": str(path)})


def hold_lock(folder: Path):
    """지금 판 job_lock 과 같은 1바이트 잠금(실행중.lock)을 잡는다. 푸는 함수를 돌려준다."""
    f = open(folder / "실행중.lock", "a+b")
    f.seek(0)
    msvcrt.locking(f.fileno(), msvcrt.LK_NBLCK, 1)

    def release():
        if f.closed:
            return
        f.seek(0)
        msvcrt.locking(f.fileno(), msvcrt.LK_UNLCK, 1)
        f.close()

    return release


def catch_up(env: Env, now: datetime, hooks: Path | None = None, notify: bool = False) -> dict:
    out = env.work.parent / f"놓친실행_{time.monotonic_ns()}.json"
    args = ["-Command", "놓친실행", "-Work", str(env.work), "-Now", now.strftime("%Y-%m-%dT%H:%M:%S"),
            "-Output", str(out), "-TestPdfWarningOff", "1"]
    if hooks is not None:
        args += ["-TestHooks", str(hooks)]
    if not notify:
        args.append("-NoNotify")
    code, text = powershell(*args)
    assert out.exists(), text
    data = json.loads(out.read_text(encoding="utf-8-sig"))
    data["exit_code"] = code
    return data


# --- 자체 시험 자료(작업 11 이 묶음에 넣는 것) ---

def test_self_test_folder_shape(tmp_path):
    """자체시험\\ = 가짜 PDF 1 · 가짜 엑셀 1 · 규칙.json · 정답.json, 정답은 지금 판 엔진 결과(엔진 JSON 표기)."""
    folder = build_self_test(tmp_path)

    assert folder == tmp_path / SELF_TEST_DIR
    answer = json.loads((folder / SELF_TEST_ANSWER).read_text(encoding="utf-8"))
    assert sorted(p.name for p in folder.iterdir()) == sorted(
        [SELF_TEST_ANSWER, SELF_TEST_RULES, answer["pdf"]["file"], answer["excel"]["file"]])
    assert answer["version"] == 1
    assert (folder / answer["pdf"]["file"]).read_bytes() == INVOICE.read_bytes()
    assert (folder / answer["excel"]["file"]).read_bytes() == ACME.read_bytes()
    pdf = answer["pdf"]["result"]
    assert set(pdf) == {"doc_type", "form", "is_excel", "fields", "table", "warnings", "review_reason"}
    assert pdf["doc_type"] == "Invoice" and pdf["review_reason"] is None and pdf["warnings"] == []
    # 워드가 「Bill To / Ship To」 칸 문단 여러 개를 높이로 묶어야 나오는 값(묶음이 다르면 이 값이 달라진다)
    assert pdf["fields"]["고객명"] == "ACME Semiconductor Korea Co., Ltd."
    assert pdf["fields"]["합계금액"] == "115440.00" and pdf["fields"]["발행일"] == "2026-10-14"
    assert len(pdf["table"]) == 4 and pdf["table"][0]["수량"] == 2
    excel = answer["excel"]["result"]
    assert excel["form"] == "ACME" and excel["is_excel"] is True and len(excel["table"]) == 4
    rules = json.loads((folder / SELF_TEST_RULES).read_text(encoding="utf-8"))
    assert rules["company_form"] is None and [d["name"] for d in rules["doc_types"]] == ["DM", "Invoice", "PO"]
    assert "고객마스터" not in (folder / SELF_TEST_RULES).read_text(encoding="utf-8")


# --- 예약 · 바로가기 (오피스 없이) ---

def test_schedule_whatif_time(tmp_path, run_engine):
    """작업 3 → 매일 16:40, 현재 사용자 · 로그인했을 때만, 놓치면 켜자마자(StartWhenAvailable). 아무것도 등록하지 않는다."""
    env = Env(tmp_path, "누적장부", example_text(), number=3, title="셋째 일")
    data = settings(env.folder)
    del data["time"]   # 시각이 없으면 번호로: 16:00 + 20분 × 2
    (env.folder / "작업설정.json").write_text(json.dumps(data, ensure_ascii=False), encoding="utf-8")
    out = run_engine("setup", {"action": "schedule", "folder": str(env.folder), "work": str(env.work), "whatif": True})

    text = out["text"]
    assert list(out) == ["text"]
    assert "-Daily -At '16:40'" in text
    user = subprocess.run(["whoami"], capture_output=True, text=True).stdout.strip()
    assert f"-UserId '{user}'".casefold() in text.casefold()
    assert "-LogonType Interactive" in text and "-RunLevel Limited" in text
    assert "-StartWhenAvailable" in text
    assert f'-File "{DISPATCHER}" -Command 실행 -Job 03 -Auto -Work "{env.work}"' in text
    assert "-TaskPath '\\사무자동화 무설치판\\' -TaskName '03_셋째 일'" in text
    assert settings(env.folder)["method"] == "없음"   # -WhatIf 는 설정도 바꾸지 않는다

    # 마지막 줄(Register-ScheduledTask)만 빼고 돌려 본다: 작업 객체만 만들고 아무것도 등록하지 않는다(매개변수 · 따옴표 검사)
    lines = text.split("\n")
    assert lines[-1].startswith("Register-ScheduledTask ") and len(lines) == 5
    script = tmp_path / "whatif.ps1"
    script.write_bytes(b"\xef\xbb\xbf" + ("$ErrorActionPreference = 'Stop'\n" + "\n".join(lines[:-1]) +
                                         "\n$action.Arguments\n([datetime]$trigger.StartBoundary).ToLocalTime().ToString('HH:mm')"
                                         "\n$settings.StartWhenAvailable\n").encode("utf-8"))
    done = subprocess.run(["powershell.exe", "-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass", "-File", str(script)],
                          capture_output=True, timeout=120)
    shown = decode_output(done.stdout).splitlines()
    assert done.returncode == 0, decode_output(done.stderr)
    assert shown[0].endswith(f'-Command 실행 -Job 03 -Auto -Work "{env.work}"')
    assert shown[1] == "16:40" and shown[2] == "True"


def test_schedule_falls_back(tmp_path, run_engine):
    """작업 스케줄러가 되면 작업스케줄러 / 실패하면 시작프로그램 「놓친 실행」 바로가기 / 그것도 실패하면 없음."""
    env = Env(tmp_path, "누적장부", example_text())
    startup = tmp_path / "startup"
    startup.mkdir()
    base = {"action": "schedule", "folder": str(env.folder), "work": str(env.work), "startup_dir": str(startup)}

    ok = run_engine("setup", {**base, "task": "ok"})
    assert ok["method"] == "작업스케줄러" and list(startup.iterdir()) == []
    assert settings(env.folder)["method"] == "작업스케줄러" and settings(env.folder)["time"] == "16:00"

    fallback = run_engine("setup", {**base, "task": "fail"})
    assert fallback["method"] == "시작프로그램"
    assert any("작업 스케줄러에 등록하지 못함(시험: 작업 스케줄러 막힘)" in n for n in fallback["notes"])
    link = startup / "사무자동화 무설치판 놓친 작업 실행.lnk"
    made = read_shortcut(run_engine, link)
    assert made["target"].lower().endswith("powershell.exe")
    assert f'-File "{DISPATCHER}" -Command 놓친실행 -Work "{env.work}"' in made["arguments"]
    assert settings(env.folder)["method"] == "시작프로그램"

    blocked = tmp_path / "startup_is_a_file"
    blocked.write_text("x", encoding="utf-8")
    none = run_engine("setup", {**base, "task": "fail", "startup_dir": str(blocked)})
    assert none["method"] == "없음"
    assert any("시작프로그램 바로가기를 만들지 못함" in n for n in none["notes"])
    assert settings(env.folder)["method"] == "없음"


def test_run_shortcut(tmp_path, run_engine):
    """바탕화면 「지금 실행 - 제목」 = 무설치판\\지금실행.bat NN (작업 폴더가 묶음 폴더가 아니면 -Work 를 붙임)."""
    env = Env(tmp_path, "누적장부", example_text(), number=2, title="둘째 일")
    data = settings(env.folder)
    data["title"] = "둘째: 일"   # 파일 이름에 못 쓰는 글자는 _ (지금 판 _folder_title)
    (env.folder / "작업설정.json").write_text(json.dumps(data, ensure_ascii=False), encoding="utf-8")
    desktop = tmp_path / "desktop"
    desktop.mkdir()
    out = run_engine("setup", {"action": "shortcut", "folder": str(env.folder), "work": str(env.work),
                               "desktop": str(desktop)})
    link = desktop / "지금 실행 - 둘째_ 일.lnk"
    assert out["path"] == str(link) and link.exists()
    made = read_shortcut(run_engine, link)
    assert Path(made["target"]) == RUN_BAT
    assert made["arguments"] == f'02 -Work "{env.work}"'

    same = run_engine("setup", {"action": "shortcut", "folder": str(env.folder), "work": str(REPO),
                                "desktop": str(desktop)})
    assert read_shortcut(run_engine, Path(same["path"]))["arguments"] == "02"


def test_office_version_not_fixed():
    """설계 위험 1: 레지스트리 경로에 오피스 버전(16.0)을 박지 않는다(설치된 버전을 HKCR\\…\\CurVer 에서 찾음)."""
    for path in PORTABLE.rglob("*.ps1"):
        text = path.read_text(encoding="utf-8-sig")
        assert not re.search(r"Office\\\d+\.\d", text), path.name
    office_ps1 = (PORTABLE / "lib" / "Office.ps1").read_text(encoding="utf-8-sig")
    assert "'HKCU:\\Software\\Microsoft\\Office\\' + $version + '\\Word\\Options'" in office_ps1


def test_bat_files():
    """준비.bat · 지금실행.bat: BOM 없이 첫 줄 @echo off, 다음 chcp 65001, 같은 폴더의 무설치판.ps1 을 부른다.
    (BOM 이 있으면 cmd 가 첫 줄을 「'BOM@echo'은(는) … 배치 파일이 아닙니다」 로 실패하고 echo 가 켜진 채 남는다.)"""
    setup = SETUP_BAT.read_bytes()
    text = setup.decode("utf-8")
    assert text.splitlines()[:2] == ["@echo off", "chcp 65001 > nul"]
    assert 'powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0무설치판.ps1" -Command 준비' in text
    run = RUN_BAT.read_bytes().decode("utf-8")
    assert '-Command 실행 -Job %1 %2 %3' in run
    bats = list(PORTABLE.rglob("*.bat"))
    assert sorted(b.name for b in bats) == ["준비.bat", "지금실행.bat"]
    for bat in bats:   # 작업 트리의 바이트 그대로(묶음은 이 파일들을 복사한다): BOM 없음 · 첫 줄 @echo off · 줄끝 모두 CRLF
        data = bat.read_bytes()
        assert data.startswith(b"@echo off\r\n") and data.endswith(b"\r\n"), bat.name
        assert data.count(b"\n") == data.count(b"\r\n"), f"LF 만 있는 줄: {bat.name}"


FAKE_DISPATCHER = """\
param([string]$Command, [string]$Job)
$out = Join-Path $PSScriptRoot '받음.json'
$data = [ordered]@{ root = $PSScriptRoot; command = $Command; job = $Job }
[IO.File]::WriteAllText($out, (ConvertTo-Json -InputObject $data -Depth 20), (New-Object Text.UTF8Encoding $false))
"""


@pytest.mark.parametrize("bat, args, command", [(SETUP_BAT, [], "준비"), (RUN_BAT, ["01"], "실행")],
                         ids=["setup", "run"])
def test_bat_runs_in_korean_folder(bat, args, command, tmp_path):
    """진짜 .bat 바이트를 한글 · 띄어쓰기 폴더에 복사해 cmd 로 돌린다. 옆의 무설치판.ps1 은 받은 것만 적는 가짜라
    예약 등록 · 바탕화면 · 폴더 고르기는 일어나지 않는다. 첫 줄 오류 · 명령 되풀이(echo) 없이 %~dp0 가 한글 경로로 풀린다."""
    folder = tmp_path / "무설치 묶음 폴더" / "무설치판"
    folder.mkdir(parents=True)
    shutil.copyfile(bat, folder / bat.name)
    (folder / "무설치판.ps1").write_bytes(b"\xef\xbb\xbf" + FAKE_DISPATCHER.encode("utf-8"))

    # 전체 경로로 부른다(NoDefaultCurrentDirectoryInExePath 가 켜진 환경에서는 cmd 가 현재 폴더의 이름만으로 찾지 않음)
    done = subprocess.run(["cmd.exe", "/c", str(folder / bat.name), *args], cwd=folder, input=b"\r\n",
                          capture_output=True, timeout=120)

    raw = done.stdout + done.stderr
    shown = raw.decode("utf-8", errors="replace") + raw.decode("mbcs", errors="replace")
    assert b"@echo" not in raw and "배치 파일이 아닙니다" not in shown, shown
    assert b"chcp" not in raw and b"powershell.exe" not in raw, shown   # echo off: 명령 줄이 화면에 찍히지 않음
    got = json.loads((folder / "받음.json").read_text(encoding="utf-8"))
    assert got["root"] == str(folder)
    assert got["command"] == command and got["job"] == ("01" if args else "")


# --- 놓친 실행 · 자동 실행 실패 알림 (오피스 없이: 입력 폴더가 비어 있음) ---

def test_catchup_runs_only_due(tmp_path):
    """15:00 → 없음, 17:00 · last_run 어제 → 실행, last_run 오늘 → 없음."""
    env = Env(tmp_path, "누적장부", example_text())   # time 16:00, last_run 없음

    early = catch_up(env, datetime(2026, 10, 30, 15, 0))
    assert early["ran"] == [] and early["exit_code"] == 0
    assert settings(env.folder)["last_run"] is None

    data = settings(env.folder)
    data["last_run"] = "2026-10-29"
    (env.folder / "작업설정.json").write_text(json.dumps(data, ensure_ascii=False), encoding="utf-8")
    late = catch_up(env, datetime(2026, 10, 30, 17, 0))
    assert late["ran"] == ["01_일일 정리"] and late["exit_code"] == 0
    assert settings(env.folder)["last_run"] == "2026-10-30"
    assert "] 놓친 실행\n" in (env.folder / "실행기록.txt").read_text(encoding="utf-8")

    again = catch_up(env, datetime(2026, 10, 30, 18, 0))
    assert again["ran"] == []


def test_catchup_time_from_settings_and_slot(tmp_path):
    """작업설정.json 의 time 이 있으면 그 시각, 없으면 16:00 + 20분 × (번호 − 1)."""
    two = Env(tmp_path, "누적장부", example_text(), number=2, title="둘째", work=tmp_path / "work")
    data = settings(two.folder)
    del data["time"]
    (two.folder / "작업설정.json").write_text(json.dumps(data, ensure_ascii=False), encoding="utf-8")
    one = Env(tmp_path, "누적장부", example_text(), number=1, title="첫째", work=tmp_path / "work")
    data = settings(one.folder)
    data["time"] = "09:05"
    (one.folder / "작업설정.json").write_text(json.dumps(data, ensure_ascii=False), encoding="utf-8")

    assert catch_up(one, datetime(2026, 10, 30, 10, 0))["ran"] == ["01_첫째"]
    assert catch_up(one, datetime(2026, 10, 30, 16, 19))["ran"] == []
    assert catch_up(one, datetime(2026, 10, 30, 16, 20))["ran"] == ["02_둘째"]


def test_auto_run_waits_for_lock_then_runs(tmp_path):
    """자동 실행이 잠긴 작업을 만나면 기다렸다 다시 해 본다(지금 판 retry_locked). 그사이 풀리면 정상 실행."""
    env = Env(tmp_path, "누적장부", example_text())
    hooks = write_hooks(tmp_path, lock_retry=[1, 20])
    release = hold_lock(env.folder)
    timer = threading.Timer(2.5, release)
    timer.start()
    try:
        code, text = powershell("-Command", "실행", "-Job", "1", "-Work", str(env.work), "-Auto",
                                "-TestHooks", str(hooks), "-TestPdfWarningOff", "1")
    finally:
        timer.cancel()
        release()
    assert code == 0, text
    assert settings(env.folder)["last_run"] is not None
    assert notices(tmp_path) == []


def test_auto_run_locked_notifies(tmp_path):
    """끝내 잠겨 있으면 반환 2, 실행기록에 남기고 실패 알림(아무도 안 보는 자동 실행이 조용히 끝나지 않게)."""
    env = Env(tmp_path, "누적장부", example_text())
    hooks = write_hooks(tmp_path, lock_retry=[1, 2])
    release = hold_lock(env.folder)
    try:
        started = time.monotonic()
        code, text = powershell("-Command", "실행", "-Job", "1", "-Work", str(env.work), "-Auto",
                                "-TestHooks", str(hooks), "-TestPdfWarningOff", "1")
        waited = time.monotonic() - started
    finally:
        release()
    assert code == 2, text
    assert waited >= 2
    locked = ("일일 정리 작업이 다른 일로 잠겨 있어 오늘 자동 실행을 못 했습니다. 지금 실행을 눌러 주세요"
              "(바탕화면 「지금 실행 - 일일 정리」).")
    record = (env.folder / "실행기록.txt").read_text(encoding="utf-8")
    assert "] 자동 실행 실패\n" + locked in record
    sent = notices(tmp_path)
    assert len(sent) == 1
    assert sent[0]["title"] == "사무자동화: 자동 실행 실패"
    assert sent[0]["body"] == f"작업 01 자동 실행: {locked}\n기록: {env.folder / '실행기록.txt'}"
    assert sent[0]["folder"] == str(env.folder)
    assert settings(env.folder)["last_run"] is None


def test_auto_run_unexpected_error_notifies(tmp_path):
    """자동 실행이 작업을 찾지 못함(반환 3) → 작업 폴더 실행기록 · 실패 알림."""
    env = Env(tmp_path, "누적장부", example_text())
    hooks = write_hooks(tmp_path)
    code, text = powershell("-Command", "실행", "-Job", "7", "-Work", str(env.work), "-Auto", "-TestHooks", str(hooks))
    assert code == 3, text
    sent = notices(tmp_path)
    assert len(sent) == 1 and sent[0]["title"] == "사무자동화: 자동 실행 실패"
    assert sent[0]["body"].startswith("작업 07 자동 실행: 작업 7 이 없습니다")
    assert TELL_BOSS in sent[0]["body"]
    assert f"기록: {env.work / '실행기록.txt'}" in sent[0]["body"]
    assert "] 자동 실행 실패\n작업 7 이 없습니다" in (env.work / "실행기록.txt").read_text(encoding="utf-8")


def test_manual_run_locked_does_not_notify(tmp_path):
    """「지금 실행」(사람이 화면을 봄)은 잠겨도 기다리지 않고 알림도 띄우지 않는다(작업 9 그대로)."""
    env = Env(tmp_path, "누적장부", example_text())
    hooks = write_hooks(tmp_path, lock_retry=[5, 60])
    release = hold_lock(env.folder)
    try:
        started = time.monotonic()
        code, _ = powershell("-Command", "실행", "-Job", "1", "-Work", str(env.work), "-TestHooks", str(hooks))
        assert time.monotonic() - started < 5
    finally:
        release()
    assert code == 2 and notices(tmp_path) == []


def test_catchup_locked_notifies(tmp_path):
    """놓친 실행도 잠긴 작업은 기다렸다 다시 해 보고, 끝내 잠겨 있으면 그 작업에 기록하고 모아서 알린다."""
    env = Env(tmp_path, "누적장부", example_text())
    hooks = write_hooks(tmp_path, lock_retry=[1, 2])
    release = hold_lock(env.folder)
    try:
        out = catch_up(env, datetime(2026, 10, 30, 17, 0), hooks, notify=True)
    finally:
        release()
    assert out["ran"] == [] and out["locked"] == ["01_일일 정리"] and out["exit_code"] == 1
    assert "] 놓친 실행 실패\n일일 정리 작업이 다른 일로 잠겨 있어" in (env.folder / "실행기록.txt").read_text(encoding="utf-8")
    sent = notices(tmp_path)
    assert len(sent) == 1 and sent[0]["title"] == "사무자동화: 자동 실행 실패"
    assert sent[0]["body"].startswith("놓친 작업 실행: 01_일일 정리: 일일 정리 작업이 다른 일로 잠겨 있어")
    assert sent[0]["folder"] == str(env.work)


# --- 준비 (오피스 없이: 자체 시험은 끼운 결과, 마크업표는 미리 둠) ---

def _setup_dirs(tmp_path):
    bundle = tmp_path / "묶음"
    work = tmp_path / "작업 폴더"
    inp, out = tmp_path / "입력", tmp_path / "출력"
    for d in (bundle, work, inp, out):
        d.mkdir()
    (work / "마크업표.xlsx").write_bytes(b"friend's table")   # 있으면 건드리지 않는다
    return bundle, work, inp, out


def test_setup_update_keeps_records(tmp_path, run_engine):
    """같은 묶음으로 준비 두 번 → 두 번째는 폴더를 묻지 않고, 처리기록 · 장부 그대로, 규칙.json 만 새 것."""
    bundle, work, inp, out = _setup_dirs(tmp_path)
    source = make_bundle(bundle, example_text())

    code, text = run_setup(bundle, work, write_hooks(tmp_path, pick=[str(inp), str(out)]))
    assert code == 0, text
    assert PDF_GUIDE not in text
    folder = work / "작업" / "01_일일 정리"
    first = settings(folder)
    assert first["input_dir"] == str(inp) and first["output_dir"] == str(out)
    assert first["title"] == "일일 정리" and first["output_mode"] == "누적장부"
    assert first["number"] == 1 and first["time"] == "16:00" and first["method"] == "시작프로그램"
    assert (folder / "규칙.json").read_bytes() == (source / "규칙.json").read_bytes()
    link = tmp_path / "desktop" / "지금 실행 - 일일 정리.lnk"
    assert read_shortcut(run_engine, link)["arguments"] == f'01 -Work "{work}"'
    assert (tmp_path / "startup" / "사무자동화 무설치판 놓친 작업 실행.lnk").exists()
    assert (work / "마크업표.xlsx").read_bytes() == b"friend's table"
    assert json.loads((work / "자체시험결과.json").read_text(encoding="utf-8"))["pdf_ok"] is True

    (folder / "처리기록.csv").write_bytes("﻿지문,파일\r\nabc,x.pdf\r\n".encode("utf-8"))
    (out / "누적장부_일일 정리.xlsx").write_bytes(b"ledger bytes")
    first["last_run"] = "2026-10-29"
    (folder / "작업설정.json").write_text(json.dumps(first, ensure_ascii=False), encoding="utf-8")
    new_rules = json.loads((source / "규칙.json").read_text(encoding="utf-8"))
    new_rules["fingerprint"] = "2" * 16
    (source / "규칙.json").write_text(json.dumps(new_rules, ensure_ascii=False), encoding="utf-8")

    code, text = run_setup(bundle, work, write_hooks(tmp_path, pick=[]))   # 폴더를 물으면 「취소」 가 되어 멈춘다
    assert code == 0, text
    assert "폴더 고르기" not in text
    assert (folder / "규칙.json").read_bytes() == (source / "규칙.json").read_bytes()
    assert (folder / "처리기록.csv").read_bytes() == "﻿지문,파일\r\nabc,x.pdf\r\n".encode("utf-8")
    assert (out / "누적장부_일일 정리.xlsx").read_bytes() == b"ledger bytes"
    second = settings(folder)
    assert (second["input_dir"], second["output_dir"], second["last_run"]) == (str(inp), str(out), "2026-10-29")


def test_setup_in_bundle_folder(tmp_path):
    """친구 PC 처럼 작업 폴더 = 묶음 폴더: 묶음의 작업 폴더에 작업설정.json 이 생기고 규칙.json 은 그대로."""
    bundle, _, inp, out = _setup_dirs(tmp_path)
    (bundle / "마크업표.xlsx").write_bytes(b"x")
    source = make_bundle(bundle, example_text(), number=2, title="둘째")
    before = (source / "규칙.json").read_bytes()
    code, text = run_setup(bundle, bundle, write_hooks(tmp_path, pick=[str(inp), str(out)], schedule="ok"), warning_off="0")
    assert code == 0, text
    assert PDF_GUIDE in text   # 워드의 PDF 변환 안내가 켜져 있으면 끄는 법을 알려 준다(설정은 바꾸지 않음)
    data = settings(source)
    assert data["input_dir"] == str(inp) and data["time"] == "16:20" and data["method"] == "작업스케줄러"
    assert (source / "규칙.json").read_bytes() == before
    assert "자동 실행: Windows 작업 스케줄러 (매일 16:20)" in text


def _rename_in_bundle(bundle: Path, rules_text: str, fingerprint: str) -> Path:
    """대표가 같은 번호 작업의 제목을 바꾼 새 묶음(01_새 제목, 출력 방식 둘다)을 친구가 덮어쓴 모양: 옛 폴더는 그대로 남는다."""
    new = make_bundle(bundle, rules_text, title="새 제목", mode="둘다")
    data = json.loads((new / "규칙.json").read_text(encoding="utf-8"))
    data["fingerprint"] = fingerprint
    (new / "규칙.json").write_text(json.dumps(data, ensure_ascii=False), encoding="utf-8")
    return new


def _job_folders(work: Path) -> list[str]:
    return sorted(p.name for p in (work / "작업").iterdir() if p.is_dir())


@pytest.mark.parametrize("layout", ["work_is_bundle", "work_apart"])
def test_setup_renamed_job_keeps_one_folder(layout, tmp_path):
    """같은 번호 · 다른 제목의 새 묶음으로 다시 준비: 번호로 맞춰 친구 PC 의 옛 폴더(이름 그대로)를 쓰고, 규칙 · 대표 값(제목 ·
    출력 방식)만 새것으로. 입력 · 출력 · 시각 · 처리기록 · last_run 은 그대로, 폴더를 묻지 않고, 폴더 · 예약 · 바로가기는 하나."""
    bundle, work, inp, out = _setup_dirs(tmp_path)
    if layout == "work_is_bundle":
        work = bundle   # 친구 PC 기본 모양: 묶음 폴더가 곧 작업 폴더
        (work / "마크업표.xlsx").write_bytes(b"x")
    old = make_bundle(bundle, example_text(), title="옛 제목")
    log = tmp_path / "예약.txt"
    code, text = run_setup(bundle, work, write_hooks(tmp_path, pick=[str(inp), str(out)], schedule="ok", schedule_log=str(log)))
    assert code == 0, text
    folder = work / "작업" / "01_옛 제목"
    (folder / "처리기록.csv").write_bytes("\ufeff지문,파일\r\nabc,x.pdf\r\n".encode("utf-8"))
    first = settings(folder)
    first["last_run"] = "2026-10-29"
    (folder / "작업설정.json").write_text(json.dumps(first, ensure_ascii=False), encoding="utf-8")
    os.utime(old / "규칙.json", (time.time() - 3600, time.time() - 3600))   # 옛 묶음 파일이 더 오래됨
    new = _rename_in_bundle(bundle, example_text(), "9" * 16)
    new_rules = (new / "규칙.json").read_bytes()

    code, text = run_setup(bundle, work, write_hooks(tmp_path, pick=[], schedule="ok", schedule_log=str(log)))
    assert code == 0, text
    assert "폴더 고르기" not in text
    assert _job_folders(work) == ["01_옛 제목"]
    assert (folder / "규칙.json").read_bytes() == new_rules
    second = settings(folder)
    assert (second["title"], second["output_mode"]) == ("새 제목", "둘다")
    for key in ("input_dir", "output_dir", "number", "time", "method", "last_run"):
        assert second[key] == first[key], key
    assert (folder / "처리기록.csv").read_bytes() == "\ufeff지문,파일\r\nabc,x.pdf\r\n".encode("utf-8")
    assert sorted(set(log.read_text(encoding="utf-8-sig").splitlines())) == ["01_옛 제목"]   # 예약 하나(같은 작업 이름을 다시 씀)
    assert sorted(p.name for p in (tmp_path / "desktop").iterdir()) == ["지금 실행 - 새 제목.lnk"]

    code, text = run_setup(bundle, work, write_hooks(tmp_path, pick=[], schedule="ok", schedule_log=str(log)))
    assert code == 0, text   # 한 번 더 준비해도 옛 제목 · 옛 규칙으로 되돌아가지 않는다
    assert _job_folders(work) == ["01_옛 제목"] and (folder / "규칙.json").read_bytes() == new_rules
    assert settings(folder)["title"] == "새 제목"

    report = tmp_path / "보고.json"
    code, text = powershell("-Command", "실행", "-Job", "1", "-Work", str(work), "-Output", str(report),
                            "-NoNotify", "-TestPdfWarningOff", "1")
    assert code == 0, text


def test_setup_update_failure_keeps_new_bundle(tmp_path):
    """작업 폴더 = 묶음 폴더에서 제목이 바뀐 새 묶음으로 다시 준비하는데 기존 작업 갱신이 실패하면(작업설정.json 이 깨짐),
    새 묶음 폴더(새 규칙의 유일한 사본)를 치우지 않고 「넣었습니다」 같은 거짓 줄도 찍지 않는다. 고친 뒤 다시 준비하면 새 규칙이 들어간다."""
    bundle, _, inp, out = _setup_dirs(tmp_path)
    work = bundle
    (work / "마크업표.xlsx").write_bytes(b"x")
    old = make_bundle(bundle, example_text(), title="옛 제목")
    code, text = run_setup(bundle, work, write_hooks(tmp_path, pick=[str(inp), str(out)], schedule="ok"))
    assert code == 0, text
    folder = work / "작업" / "01_옛 제목"
    good = (folder / "작업설정.json").read_bytes()
    (folder / "작업설정.json").write_text("{ 깨진 json", encoding="utf-8")
    os.utime(old / "규칙.json", (time.time() - 3600, time.time() - 3600))
    new = _rename_in_bundle(bundle, example_text(), "9" * 16)
    new_rules = (new / "규칙.json").read_bytes()

    code, text = run_setup(bundle, work, write_hooks(tmp_path, pick=[], schedule="ok"))

    assert code == 1, text
    assert "01_옛 제목: 바꾸지 못했습니다" in text
    assert new.is_dir() and (new / "규칙.json").read_bytes() == new_rules
    assert "정리:" not in text and "넣었습니다" not in text
    assert _job_folders(work) == ["01_새 제목", "01_옛 제목"]

    (folder / "작업설정.json").write_bytes(good)   # 고친 뒤 다시 준비 → 새 규칙이 들어가고 새 묶음 폴더는 치워진다
    code, text = run_setup(bundle, work, write_hooks(tmp_path, pick=[], schedule="ok"))
    assert code == 0, text
    assert (folder / "규칙.json").read_bytes() == new_rules
    assert _job_folders(work) == ["01_옛 제목"]
    assert "새 규칙은 01_옛 제목 에 넣었습니다" in text


def test_rejects_onedrive_folder(tmp_path):
    """고른 폴더가 OneDrive 아래 → 지금 판 check_work_dir 와 같은 거부 문구, 다시 고르게 하고, 취소하면 준비를 멈춘다."""
    bundle, work, _, _ = _setup_dirs(tmp_path)
    make_bundle(bundle, example_text())
    synced = tmp_path / "OneDrive - ACME" / "입력"
    synced.mkdir(parents=True)
    with pytest.raises(LocalOnlyError) as refused:
        check_work_dir(synced)

    code, text = run_setup(bundle, work, write_hooks(tmp_path, pick=[str(synced)]))
    assert code == 1
    assert str(refused.value) in text
    assert "이 폴더는 쓸 수 없습니다. 내 PC 의 로컬 디스크 폴더를 다시 고르세요." in text
    assert "폴더를 고르지 않아 준비를 멈췄습니다" in text
    assert not (work / "작업" / "01_일일 정리" / "작업설정.json").exists()
    assert list((tmp_path / "desktop").iterdir()) == []


def test_setup_stops_when_excel_self_test_fails(tmp_path):
    """엑셀 자체 시험이 틀리면 예약 · 바로가기 없이 멈추고 대표에게 알리라고 한다."""
    bundle, work, inp, out = _setup_dirs(tmp_path)
    make_bundle(bundle, example_text())
    code, text = run_setup(bundle, work, write_hooks(tmp_path, pick=[str(inp), str(out)], self_test="excel_fail"))
    assert code == 1
    assert TELL_BOSS in text and "준비를 멈췄습니다" in text
    assert not (work / "작업" / "01_일일 정리" / "작업설정.json").exists()
    assert list((tmp_path / "desktop").iterdir()) == [] and list((tmp_path / "startup").iterdir()) == []
    saved = json.loads((work / "자체시험결과.json").read_text(encoding="utf-8"))
    assert saved["excel_ok"] is False


def test_pdf_self_test_failure_is_kept_and_blocks_pdfs(tmp_path):
    """PDF 자체 시험이 틀리면 준비는 계속하되 자체시험결과.json 에 남기고, 실행은 PDF 를 기록 없이 두고 알린다."""
    bundle, work, inp, out = _setup_dirs(tmp_path)
    make_bundle(bundle, example_text())
    code, text = run_setup(bundle, work, write_hooks(tmp_path, pick=[str(inp), str(out)], self_test="pdf_fail"))
    assert code == 0, text
    assert PDF_DIFFERENT in text
    saved = json.loads((work / "자체시험결과.json").read_text(encoding="utf-8"))
    assert saved["pdf_ok"] is False and saved["excel_ok"] is True
    assert set(saved) >= {"pdf_ok", "excel_ok", "date", "office_version"}
    assert saved["date"] == datetime.now().strftime("%Y-%m-%d")

    folder = work / "작업" / "01_일일 정리"
    drop(INVOICE, inp)
    report_path = tmp_path / "보고.json"
    code, text = powershell("-Command", "실행", "-Job", "1", "-Work", str(work), "-Output", str(report_path),
                            "-NoNotify", "-TestPdfWarningOff", "1")
    assert code == 0, text
    report = json.loads(report_path.read_text(encoding="utf-8-sig"))
    assert report["counts"] == {"성공": 0, "중복": 0, "확인필요": 0, "보류": 0}
    assert not (folder / "처리기록.csv").exists() and not (folder / "처리로그.csv").exists()
    assert report["has_news"] is True
    assert any("PDF 1개를 처리하지 않았습니다" in a and PDF_DIFFERENT in a for a in report["alerts"])
    assert files(out) == []


def test_excel_self_test_failure_blocks_excel_pos(tmp_path):
    """자체시험결과.json 의 excel_ok 가 거짓이면(다시 준비하다 엑셀 시험이 틀림) 매일 실행도 엑셀 PO 를 기록 없이 두고 알린다."""
    env = Env(tmp_path, "누적장부", example_text())
    (env.work / "자체시험결과.json").write_text(json.dumps(
        {"pdf_ok": True, "excel_ok": False, "date": "2026-10-30", "office_version": "16.0",
         "message": "이 PC 의 엑셀은 엑셀 PO 를 다르게 읽습니다", "detail": ""}, ensure_ascii=False), encoding="utf-8")
    drop(ACME, env.inp)
    report_path = tmp_path / "보고.json"
    code, text = powershell("-Command", "실행", "-Job", "1", "-Work", str(env.work), "-Output", str(report_path),
                            "-NoNotify", "-TestPdfWarningOff", "1")
    assert code == 0, text
    report = json.loads(report_path.read_text(encoding="utf-8-sig"))
    assert report["counts"] == {"성공": 0, "중복": 0, "확인필요": 0, "보류": 0}
    assert env.records() == [] and env.log() == [] and files(env.out) == []
    assert report["has_news"] is True
    assert any("엑셀 PO 1개를 처리하지 않았습니다" in a and "이 PC 의 엑셀은 엑셀 PO 를 다르게 읽습니다" in a
               for a in report["alerts"]), report["alerts"]


# --- 오피스 ---

@office
@pytest.mark.office
def test_self_test_passes_here(tmp_path, run_engine, clean_temp, no_leftover):
    """이 PC(대표 PC)의 워드 · 엑셀로 자체시험\\ 을 읽으면 정답과 같다."""
    build_self_test(tmp_path)
    started = time.monotonic()
    out = run_engine("setup", {"action": "self_test", "bundle": str(tmp_path)})
    print(f"자체 시험 시간: {time.monotonic() - started:.1f}초")
    assert out == {**out, "ok": True, "pdf_ok": True, "excel_ok": True}, out


@office
@pytest.mark.office
def test_self_test_catches_word_row_grouping(tmp_path, run_engine, clean_temp, no_leftover):
    """워드 표 행을 묶는 높이 기준(3pt)이 이 PC 와 다르게 동작하면(여기서는 기준을 일부러 크게) PDF 만 틀리고 그 문구를 낸다."""
    build_self_test(tmp_path)
    out = run_engine("setup", {"action": "self_test", "bundle": str(tmp_path), "row_tolerance": 1000})
    assert out["ok"] is False and out["pdf_ok"] is False and out["excel_ok"] is True
    assert out["message"] == PDF_DIFFERENT
    assert out["detail"].startswith("PDF.fields.고객명: 정답 'ACME Semiconductor Korea Co., Ltd.', 이 PC null")
    assert "PDF.warnings: 정답 0개, 이 PC 1개" in out["detail"]


@office
@pytest.mark.office
def test_self_test_skips_pdf_when_word_warning_on(tmp_path, run_engine, clean_temp, no_leftover):
    """워드의 PDF 변환 안내가 켜져 있으면(숨은 확인 창에서 멈춤) PDF 는 열지 않고 통과로 보지 않는다. 엑셀은 시험한다."""
    build_self_test(tmp_path)
    out = run_engine("setup", {"action": "self_test", "bundle": str(tmp_path), "pdf_warning_off": False})
    assert (out["ok"], out["pdf_ok"], out["excel_ok"]) == (False, False, True)
    assert out["message"] == "워드의 PDF 변환 안내가 켜져 있어 PDF 자체 시험을 하지 못했습니다. PDF 작업은 켜지 않았습니다"


@office
@pytest.mark.office
@pytest.mark.parametrize("source", [INVOICE, ACME], ids=["pdf", "xlsx"])
def test_check_command_matches_python_check(source, tmp_path, clean_temp, no_leftover):
    """대표 PC 의 점검: 콘솔 표(항목 · 값 · 비고 / 품목 / 검산 · 원본 · 결과 · 일치)가 지금 판 check 출력과 같다."""
    path = tmp_path / source.name
    shutil.copyfile(source, path)
    rules = example_rules()
    rules_json = tmp_path / "규칙.json"
    rules_json.write_text(json.dumps(rules_to_portable(rules), ensure_ascii=False), encoding="utf-8")
    printed = io.StringIO()
    with contextlib.redirect_stdout(printed):
        assert check(path, rules) == 0
    expected = printed.getvalue().splitlines()

    started = time.monotonic()
    code, text = powershell("-Command", "점검", "-File", str(path), "-Rules", str(rules_json))
    print(f"점검 시간: {time.monotonic() - started:.1f}초")
    assert code == 0, text
    got = text.splitlines()
    cut = next(i for i, line in enumerate(expected) if line.startswith("마크업표:"))
    assert got[:cut] == expected[:cut]
    assert "| 검산 | 원본 | 결과 | 일치 |" in got and "|---|---|---|---|" in got
    assert expected[-1].startswith("회사 양식: ") and got[-1] == expected[-1]


@pytest.mark.office
def test_check_command_unreadable_table_is_review(tmp_path, clean_temp, no_leftover):
    """점검도 실행과 같게: 규칙에 품목 표가 있는데 줄이 어긋나 표를 버리면 「확인필요: 품목 표를 읽지 못함: <경고>」 로 끝난다."""
    path = split_row_po(tmp_path / "쪼갠줄.pdf", "PO-77002")
    rules_json = tmp_path / "규칙.json"
    rules_json.write_text(json.dumps(rules_to_portable(rules_from_text(without_master(PO_RULES))), ensure_ascii=False),
                          encoding="utf-8")

    code, text = powershell("-Command", "점검", "-File", str(path), "-Rules", str(rules_json))

    assert code == 0, text
    assert "품목 표: 읽지 못함 (아래 경고 참고)" in text
    assert text.splitlines()[-1] == "확인필요: 품목 표를 읽지 못함: 표 행패턴과 맞지 않는 줄이 있어 품목 건너뜀: Assembly"
    assert "| 검산 | 원본 | 결과 | 일치 |" not in text


@office
@pytest.mark.office
def test_setup_with_office(tmp_path, run_engine, clean_temp, no_leftover):
    """오피스로 끝까지: 진짜 자체 시험 → 작업 · 바로가기 → 마크업표가 없으면 가짜 예시(지금 판 EXAMPLE_ROWS 와 같음)."""
    bundle, work, inp, out = _setup_dirs(tmp_path)
    (work / "마크업표.xlsx").unlink()
    build_self_test(bundle)
    make_bundle(bundle, example_text())
    started = time.monotonic()
    code, text = run_setup(bundle, work, write_hooks(tmp_path, pick=[str(inp), str(out)], self_test=None))
    print(f"준비 시간: {time.monotonic() - started:.1f}초")
    assert code == 0, text
    saved = json.loads((work / "자체시험결과.json").read_text(encoding="utf-8"))
    assert saved["pdf_ok"] is True and saved["excel_ok"] is True and saved["office_version"] == "16.0"
    table = load_markup(work / "마크업표.xlsx")
    assert [[r.part, r.rate, r.places] for r in table.values()] == [[p, r, d] for p, r, d in EXAMPLE_ROWS]
    assert openpyxl.load_workbook(work / "마크업표.xlsx").active.title == "마크업표"
    assert (tmp_path / "desktop" / "지금 실행 - 일일 정리.lnk").exists()
