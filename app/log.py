"""처리로그.csv 기록과 콘솔 요약. 장부가 잠겨 있어도 항상 기록한다."""
import csv
from pathlib import Path

COLUMNS = ["일시", "파일", "결과", "사유"]
OUTCOMES = ("성공", "중복", "확인필요")
EXTRA_OUTCOMES = ("보류", "이동 실패")


def append_log(path, when, file_name, outcome, reason):
    path = Path(path)
    new = not path.exists() or path.stat().st_size == 0
    # 새 파일만 utf-8-sig(BOM 포함), 이어 쓸 때는 BOM 없이 utf-8
    with open(path, "a", encoding="utf-8-sig" if new else "utf-8", newline="") as f:
        writer = csv.writer(f)
        if new:
            writer.writerow(COLUMNS)
        writer.writerow([when.strftime("%Y-%m-%d %H:%M:%S"), file_name, outcome, reason or ""])


def summary(outcomes) -> str:
    """성공·중복·확인필요는 항상, 보류·이동 실패는 있을 때만."""
    outcomes = list(outcomes)
    names = list(OUTCOMES) + [o for o in EXTRA_OUTCOMES if o in outcomes]
    return " · ".join(f"{name} {outcomes.count(name)}" for name in names)
