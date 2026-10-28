"""고객마스터.xlsx: 고객명 → 고객코드·결제조건·기본 인도조건. 회사 정보라 사용자가 직접 채운다."""
from pathlib import Path

import openpyxl
from openpyxl.styles import Font
from openpyxl.worksheet.table import Table, TableStyleInfo

NAME_COLUMN = "고객명"
EXAMPLE_COLUMNS = ["고객명", "고객코드", "결제조건", "기본인도조건"]
# 가짜 예시(samples 의 가짜 고객). 실제로 쓸 때는 여러분 고객으로 바꾼다.
EXAMPLE_ROWS = [
    ["ACME Semiconductor Korea Co., Ltd.", "C-KR-0417", "Net 60", "DAP Cheongun"],
    ["Hanil DataCenter Co., Ltd.", "C-KR-0522", "Net 30", "CIF Haeun"],
]


class MasterError(Exception):
    """고객마스터.xlsx 를 읽을 수 없을 때."""


def _key(name) -> str:
    return " ".join(str(name or "").split()).lower()


def load_master(path) -> dict[str, dict]:
    path = Path(path)
    if not path.exists():
        raise MasterError(f"고객마스터.xlsx 가 없습니다: {path}")
    ws = openpyxl.load_workbook(path, data_only=True).worksheets[0]
    rows = list(ws.iter_rows(values_only=True))
    header = [str(h).strip() if h is not None else "" for h in (rows[0] if rows else [])]
    if NAME_COLUMN not in header:
        raise MasterError(f"고객마스터.xlsx 첫 줄에 '{NAME_COLUMN}' 열이 있어야 합니다.")
    master = {}
    for values in rows[1:]:
        row = {h: v for h, v in zip(header, values) if h}
        if _key(row.get(NAME_COLUMN)):
            master[_key(row[NAME_COLUMN])] = row
    return master


def lookup(master, customer_name):
    """고객명으로 찾는다(대소문자·띄어쓰기 차이는 무시). 없으면 None."""
    return master.get(_key(customer_name))


def write_example(path):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "고객마스터"
    ws.append(EXAMPLE_COLUMNS)
    for row in EXAMPLE_ROWS:
        ws.append(row)
    for cell in ws[1]:
        cell.font = Font(bold=True)
    for letter, width in zip("ABCD", (40, 14, 14, 18)):
        ws.column_dimensions[letter].width = width
    table = Table(displayName="Customer_Master", ref=f"A1:D{len(EXAMPLE_ROWS) + 1}")
    table.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)
    ws.add_table(table)
    ws.freeze_panes = "A2"
    wb.save(path)
