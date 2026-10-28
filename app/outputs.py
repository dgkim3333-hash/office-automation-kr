"""문서별 결과 엑셀(출력파일). 엑셀 PO → 회사 SO 양식, PDF → 추출결과."""
from datetime import date, datetime
from decimal import Decimal
from pathlib import Path

import openpyxl
from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE
from openpyxl.styles import Font

from app.excel_extract import excel_number
from app.ledger import AMOUNT_FORMAT, DATE_FORMAT, DATETIME_FORMAT, fit_columns
from app.organize import free_name, render_text

LINE_STEP = 10  # 줄번호: 10, 20, 30 …


def _number(value) -> Decimal:
    """엑셀 칸에서 온 float 은 엑셀처럼 유효 15자리로(0.30000000000000004 → 0.3)."""
    return excel_number(value) if isinstance(value, float) else Decimal(str(value))


def _value(source, result, item, master_row, index):
    if source == "줄번호":
        return (index + 1) * LINE_STEP
    if source == "빈칸":
        return None
    if source == "수량x단가":
        qty, price = (item or {}).get("수량"), (item or {}).get("단가")
        return None if qty is None or price is None else _number(qty) * _number(price)
    if source.startswith("고정:"):
        return source[len("고정:"):]
    group, _, name = source.partition(".")
    if group == "머리":
        return result.fields.get(name)
    if group == "품목":
        return (item or {}).get(name)
    if group == "고객마스터":
        return (master_row or {}).get(name)
    return None


def company_rows(result, master_row, form):
    """회사 양식의 머리글과 행(품목 한 줄 = 한 행)."""
    header = [c.name for c in form.columns]
    items = result.table or []
    rows = [[_value(c.source, result, item, master_row, i) for c in form.columns] for i, item in enumerate(items)]
    return header, rows


def _write_sheet(ws, header, rows):
    ws.append(header)
    for cell in ws[1]:
        cell.font = Font(bold=True)
    for row in rows:
        ws.append([ILLEGAL_CHARACTERS_RE.sub("", v) if isinstance(v, str) else v for v in row])
    for row in ws.iter_rows(min_row=2):
        for cell in row:
            if isinstance(cell.value, datetime):
                cell.number_format = DATETIME_FORMAT
            elif isinstance(cell.value, date):
                cell.number_format = DATE_FORMAT
            elif isinstance(cell.value, Decimal):
                cell.number_format = AMOUNT_FORMAT
    ws.freeze_panes = "A2"
    fit_columns(ws)


def _save(wb, folder, name) -> Path:
    folder = Path(folder)
    folder.mkdir(parents=True, exist_ok=True)
    path = free_name(folder / name)
    wb.save(path)
    return path


def write_company_form(folder, result, master_row, form) -> Path:
    header, rows = company_rows(result, master_row, form)
    wb = openpyxl.Workbook()
    wb.active.title = form.sheet
    _write_sheet(wb.active, header, rows)
    return _save(wb, folder, render_text(form.filename_tpl, result.fields))


def write_extract_result(folder, result, source_name) -> Path:
    """PDF 한 건의 추출 결과: 추출결과 시트(필드·값·비고) + 품목 시트."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "추출결과"
    rows = [["문서종류", result.doc_type.name, None], ["원본파일", source_name, None]]
    rows += [[name, value, None if value is not None else "못 찾음"] for name, value in result.fields.items()]
    _write_sheet(ws, ["필드", "값", "비고"], rows)
    if result.doc_type.table is not None:
        columns = result.doc_type.table.columns
        _write_sheet(wb.create_sheet("품목"), columns, [[r.get(c) for c in columns] for r in result.table or []])
    return _save(wb, folder, render_text(f"추출결과_{{{result.doc_type.number_field}}}.xlsx", result.fields))
