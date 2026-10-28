"""처리로그.csv 와 콘솔 요약."""
import csv
from datetime import datetime

from app.log import append_log, summary


def test_append_log_writes_utf8_sig_with_header(tmp_path):
    path = tmp_path / "처리로그.csv"
    append_log(path, datetime(2026, 10, 28, 9, 5, 0), "a.pdf", "성공", "")
    append_log(path, datetime(2026, 10, 28, 9, 6, 0), "b.pdf", "확인필요", "판별 실패")

    raw = path.read_bytes()
    assert raw.startswith(b"\xef\xbb\xbf")  # 엑셀에서 한글이 안 깨지게
    assert raw.count(b"\xef\xbb\xbf") == 1
    with open(path, encoding="utf-8-sig", newline="") as f:
        rows = list(csv.reader(f))
    assert rows == [
        ["일시", "파일", "결과", "사유"],
        ["2026-10-28 09:05:00", "a.pdf", "성공", ""],
        ["2026-10-28 09:06:00", "b.pdf", "확인필요", "판별 실패"],
    ]


def test_summary_text():
    assert summary(["성공", "성공", "확인필요", "중복", "성공"]) == "성공 3 · 중복 1 · 확인필요 1"
    assert summary([]) == "성공 0 · 중복 0 · 확인필요 0"
