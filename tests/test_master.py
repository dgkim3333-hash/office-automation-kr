"""고객마스터.xlsx — 고객명 → 고객코드·결제조건·기본 인도조건."""
import shutil
from pathlib import Path

import openpyxl
import pytest

from app.init import init
from app.master import MasterError, load_master, lookup, write_example

REPO = Path(__file__).resolve().parent.parent


def test_example_master_round_trip(tmp_path):
    path = tmp_path / "고객마스터.xlsx"
    write_example(path)
    master = load_master(path)
    row = lookup(master, "ACME Semiconductor Korea Co., Ltd.")
    assert row["고객코드"] == "C-KR-0417" and row["결제조건"] == "Net 60"
    assert lookup(master, "Hanil DataCenter Co., Ltd.")["고객코드"] == "C-KR-0522"


def test_lookup_ignores_case_and_extra_spaces(tmp_path):
    path = tmp_path / "고객마스터.xlsx"
    write_example(path)
    assert lookup(load_master(path), "  acme   semiconductor korea co., ltd. ")["고객코드"] == "C-KR-0417"


def test_unknown_customer_is_none(tmp_path):
    path = tmp_path / "고객마스터.xlsx"
    write_example(path)
    assert lookup(load_master(path), "없는 회사") is None


def test_master_without_name_column_raises(tmp_path):
    path = tmp_path / "고객마스터.xlsx"
    wb = openpyxl.Workbook()
    wb.active.append(["회사", "코드"])
    wb.save(path)
    with pytest.raises(MasterError, match="고객명"):
        load_master(path)


@pytest.fixture
def repo(tmp_path):
    root = tmp_path / "repo"
    (root / "rules").mkdir(parents=True)
    shutil.copy(REPO / "rules" / "규칙.예시.yaml", root / "rules" / "규칙.예시.yaml")
    return root


def test_init_creates_master_and_jobs_folder(tmp_path, repo):
    work = init(work_dir=tmp_path / "사무자동화", repo_root=repo)
    assert (work / "작업").is_dir()
    assert lookup(load_master(work / "고객마스터.xlsx"), "Hanil DataCenter Co., Ltd.") is not None


def test_init_keeps_existing_master(tmp_path, repo):
    work = tmp_path / "사무자동화"
    work.mkdir()
    (work / "고객마스터.xlsx").write_bytes(b"my master")
    init(work_dir=work, repo_root=repo)
    assert (work / "고객마스터.xlsx").read_bytes() == b"my master"
