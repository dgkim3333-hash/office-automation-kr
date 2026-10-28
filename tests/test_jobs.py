import pytest

from app.jobs import (
    JOB_SUBFOLDERS,
    JobError,
    create_job,
    find_job,
    list_jobs,
    load_job,
    save_job,
    slot_time,
)


@pytest.fixture
def dirs(tmp_path):
    work, inp, out = tmp_path / "w", tmp_path / "in", tmp_path / "out"
    for d in (work, inp, out):
        d.mkdir()
    return work, inp, out


def test_slot_times():
    assert [slot_time(n) for n in (1, 2, 3, 4, 5)] == ["16:00", "16:20", "16:40", "17:00", "17:20"]


def test_create_job_makes_folder_and_settings(dirs):
    work, inp, out = dirs
    job = create_job(work, "일일 PO 정리", inp, out, "둘다")
    assert job.number == 1 and job.run_time == "16:00"
    assert job.folder == work / "작업" / "01_일일 PO 정리"
    assert all((job.folder / s).is_dir() for s in JOB_SUBFOLDERS)
    assert job.rules_path == job.folder / "규칙.yaml"
    assert job.name == "01_일일 PO 정리"
    again = load_job(job.folder)
    assert (again.title, again.input_dir, again.output_mode, again.approvals) == ("일일 PO 정리", inp, "둘다", {})
    assert again.output_dir == out and again.paused is False and again.scheduler == "" and again.last_run == ""


def test_second_job_gets_next_slot(dirs):
    work, inp, out = dirs
    create_job(work, "첫째", inp, out, "누적장부")
    second = create_job(work, "둘째", inp, out, "회사양식")
    assert second.number == 2 and second.run_time == "16:20"


def test_title_with_forbidden_characters_and_duplicates(dirs):
    work, inp, out = dirs
    a = create_job(work, "PO/DM 정리?", inp, out, "누적장부")
    b = create_job(work, "PO/DM 정리?", inp, out, "누적장부")
    assert a.folder.name == "01_PO_DM 정리_" and b.folder.name == "02_PO_DM 정리_"
    assert find_job(work, "2").folder == b.folder and find_job(work, "02_PO_DM 정리_").folder == b.folder


def test_title_trimmed_and_cut_to_30(dirs):
    work, inp, out = dirs
    job = create_job(work, "  " + "가" * 40 + ".  ", inp, out, "누적장부")
    assert job.folder.name == "01_" + "가" * 30
    dotted = create_job(work, "보고서.", inp, out, "누적장부")
    assert dotted.folder.name == "02_보고서"


def test_deleted_number_is_not_reused_while_higher_exists(dirs):
    work, inp, out = dirs
    create_job(work, "하나", inp, out, "누적장부")
    two = create_job(work, "둘", inp, out, "누적장부")
    three = create_job(work, "셋", inp, out, "누적장부")
    import shutil

    shutil.rmtree(three.folder)
    shutil.rmtree(two.folder)
    nxt = create_job(work, "넷", inp, out, "누적장부")
    assert nxt.number == 2  # 최대 번호(1) + 1


def test_rejects_unknown_output_mode(dirs):
    work, inp, out = dirs
    with pytest.raises(JobError, match="엑셀"):
        create_job(work, "작업", inp, out, "엑셀")


def test_rejects_missing_input_and_output_folder(dirs, tmp_path):
    work, inp, out = dirs
    with pytest.raises(JobError, match="입력 폴더가 없습니다"):
        create_job(work, "작업", tmp_path / "nope", out, "누적장부")
    with pytest.raises(JobError, match="출력 폴더가 없습니다"):
        create_job(work, "작업", inp, tmp_path / "nope", "누적장부")
    assert list_jobs(work) == []


def test_rejects_sync_folder(dirs, tmp_path):
    work, inp, out = dirs
    sync = tmp_path / "OneDrive - ACME" / "in"
    sync.mkdir(parents=True)
    with pytest.raises(JobError, match="OneDrive"):
        create_job(work, "작업", sync, out, "누적장부")
    with pytest.raises(JobError, match="OneDrive"):
        create_job(work, "작업", inp, sync, "누적장부")


def test_find_job_by_title_and_unknown_lists_existing(dirs):
    work, inp, out = dirs
    job = create_job(work, "일일 PO 정리", inp, out, "둘다")
    assert find_job(work, "일일 PO 정리").folder == job.folder
    assert find_job(work, "01").folder == job.folder
    with pytest.raises(JobError, match="01_일일 PO 정리"):
        find_job(work, "없는작업")


def test_find_job_with_no_jobs(dirs):
    work, _, _ = dirs
    with pytest.raises(JobError, match="없는작업"):
        find_job(work, "없는작업")


def test_save_and_load_roundtrip_and_list_order(dirs):
    work, inp, out = dirs
    a = create_job(work, "가", inp, out, "누적장부")
    create_job(work, "나", inp, out, "회사양식")
    a.paused = True
    a.scheduler = "작업스케줄러-1"
    a.approvals = {"Hanil SO": 10}
    a.fingerprints = {"Hanil SO": "abc123"}
    a.last_run = "2026-10-28"
    save_job(a)
    back = load_job(a.folder)
    assert back.paused is True and back.scheduler == "작업스케줄러-1"
    assert back.approvals == {"Hanil SO": 10} and back.fingerprints == {"Hanil SO": "abc123"}
    assert back.last_run == "2026-10-28"
    assert [j.number for j in list_jobs(work)] == [1, 2]


def test_list_jobs_empty_when_no_jobs_dir(dirs):
    work, _, _ = dirs
    assert list_jobs(work) == []
