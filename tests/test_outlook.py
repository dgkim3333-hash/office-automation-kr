"""(선택) Outlook 폴더 → 받은문서/. COM 은 가짜 객체로 대신한다."""
from pathlib import Path

from app.outlook import DONE_CATEGORY, collect, collect_from_outlook


class FakeAttachment:
    def __init__(self, name, data=b"%PDF-1.4 fake"):
        self.FileName = name
        self.data = data

    def SaveAsFile(self, path):
        Path(path).write_bytes(self.data)


class FakeMail:
    def __init__(self, attachments, categories=""):
        self.Attachments = attachments
        self.Categories = categories
        self.saved = False

    def Save(self):
        self.saved = True


def test_saves_only_pdf_attachments(tmp_path):
    mail = FakeMail([FakeAttachment("a.pdf"), FakeAttachment("logo.png"), FakeAttachment("B.PDF")])
    saved = collect([mail], tmp_path)
    assert sorted(p.name for p in saved) == ["B.PDF", "a.pdf"]
    assert sorted(p.name for p in tmp_path.iterdir()) == ["B.PDF", "a.pdf"]


def test_same_name_gets_suffix(tmp_path):
    (tmp_path / "a.pdf").write_bytes(b"old")
    collect([FakeMail([FakeAttachment("a.pdf")]), FakeMail([FakeAttachment("a.pdf")])], tmp_path)
    assert sorted(p.name for p in tmp_path.iterdir()) == ["a.pdf", "a_2.pdf", "a_3.pdf"]
    assert (tmp_path / "a.pdf").read_bytes() == b"old"


def test_marks_mail_done_and_skips_done(tmp_path):
    new = FakeMail([FakeAttachment("a.pdf")], categories="고객")
    done = FakeMail([FakeAttachment("old.pdf")], categories=f"고객, {DONE_CATEGORY}")
    collect([new, done], tmp_path)
    assert new.Categories == f"고객, {DONE_CATEGORY}"
    assert new.saved is True
    assert done.saved is False
    assert not (tmp_path / "old.pdf").exists()


class FakeFolders:
    """COM 컬렉션처럼 Count 와 1부터 세는 Item 만 있다."""

    def __init__(self, names):
        self.items = [type("Folder", (), {"Name": n})() for n in names]
        self.Count = len(self.items)

    def Item(self, i):
        return self.items[i - 1]


def test_find_subfolder_by_name():
    from app.outlook import find_subfolder

    assert find_subfolder(FakeFolders(["고객", "자동화"]), "자동화").Name == "자동화"


def test_find_subfolder_missing_raises():
    import pytest

    from app.outlook import find_subfolder

    with pytest.raises(LookupError):
        find_subfolder(FakeFolders(["고객"]), "자동화")


def test_without_outlook_prints_korean_message(tmp_path, capsys):
    def no_outlook():
        raise OSError("COM 없음")

    assert collect_from_outlook(tmp_path, connect=no_outlook) == []
    assert "Outlook" in capsys.readouterr().out


def test_main_outlook_command_saves_into_inbox(tmp_path, monkeypatch):
    import shutil

    import app.__main__ as cli
    from app.init import init

    repo = tmp_path / "repo"
    (repo / "rules").mkdir(parents=True)
    shutil.copy(Path(__file__).resolve().parent.parent / "rules" / "규칙.예시.yaml", repo / "rules")
    work = init(work_dir=tmp_path / "work", repo_root=repo)
    calls = []
    monkeypatch.setattr(cli, "collect_from_outlook", lambda inbox: calls.append(inbox) or [])
    assert cli.main(["outlook"], repo_root=repo) == 0
    assert calls == [work / "받은문서"]