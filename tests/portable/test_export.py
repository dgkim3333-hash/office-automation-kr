"""무설치판용 규칙 내보내기: 규칙.yaml → JSON(무설치판이 ConvertFrom-Json 으로 읽음)."""
import json
import re
import shutil
from pathlib import Path

import pytest

from app import config
from app.__main__ import main
from app.job_run import rules_fingerprint
from app.jobs import find_job
from app.portable import _check_regex, export_rules, rules_to_portable
from app.rules import RuleError, rules_from_text

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


def _make_job(repo, tmp_path, rules_text):
    assert main(["job", "new", "--제목", "시험 작업", "--입력", str(tmp_path / "in"),
                 "--출력", str(tmp_path / "out"), "--방식", "누적장부"], repo_root=repo) == 0
    job = find_job(config.work_dir(repo), "시험 작업")
    job.rules_path.write_text(rules_text, encoding="utf-8")
    return job


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
