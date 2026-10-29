"""엑셀 파일 → 무설치판 엔진이 읽는 격자 JSON (시험 전용).

모양: {"active": 활성 시트 번호(0부터), "sheets": [{"name", "cells": [[행, 열, {"t", "v", "f"}], …]}]}
  t = n(숫자) · s(글자) · b(참거짓) · d(날짜, v 는 "YYYY-MM-DDTHH:MM:SS"). f = 칸 표시 형식. 빈 칸은 싣지 않는다.
  칸은 행 순서 다음 열 순서로 늘어놓는다. 숫자는 JSON 숫자(엑셀 Value2 와 같은 double).
무설치판 작업 7 의 엑셀 읽기(COM)도 같은 모양을 낸다.
"""
from datetime import date, datetime
from pathlib import Path

import openpyxl

DATE_FORMAT = "%Y-%m-%dT%H:%M:%S"


def _cell(value):
    """(t, v). 격자가 담지 못하는 값이면 TypeError."""
    if isinstance(value, bool):  # bool 은 int 의 하위 형식이라 먼저
        return "b", value
    if isinstance(value, (int, float)):
        return "n", value
    if isinstance(value, datetime):
        return "d", value.strftime(DATE_FORMAT)
    if isinstance(value, date):
        return "d", datetime(value.year, value.month, value.day).strftime(DATE_FORMAT)
    if isinstance(value, str):
        return "s", value
    raise TypeError(f"격자에 담을 수 없는 칸 값: {type(value).__name__}")


def grid_from_xlsx(path) -> dict:
    """xlsx 를 openpyxl(data_only=True, 수식은 저장된 값)로 읽어 격자 JSON 사전으로."""
    wb = openpyxl.load_workbook(Path(path), data_only=True)
    sheets = []
    for ws in wb.worksheets:
        cells = []
        for row in ws.iter_rows():
            for cell in row:
                if cell.value is None:
                    continue
                t, v = _cell(cell.value)
                cells.append([cell.row, cell.column, {"t": t, "v": v, "f": cell.number_format}])
        sheets.append({"name": ws.title, "cells": cells})
    return {"active": wb.worksheets.index(wb.active), "sheets": sheets}
