"""테스트 전용 가짜 PDF 생성기(ACME 가짜 5종). 사용자용 샘플은 samples/ 에 따로 있다.

테스트가 임시 폴더에 PDF 를 만들 때 build(out_dir) 를 부른다.
직접 실행: uv run python tests/fixtures/build_fixtures.py <출력폴더>
"""
from pathlib import Path

from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas

CUSTOMER = "ACME Semiconductor Korea"
DATE = "10/15/2026"
TOTAL = "12,345.67"

# (파일명, 제목, 번호 줄, 품목 줄들)
DOCUMENTS = [
    (
        "ACME_Invoice_INV-2026-0001.pdf",
        "INVOICE",
        "Invoice No: INV-2026-0001",
        [
            ("CP-200", "Chiller Pump Assembly", "2", "4,500.00", "9,000.00"),
            ("FT-310", "Filter Cartridge", "10", "120.00", "1,200.00"),
            ("VL-055", "Control Valve Kit", "1", "2,145.67", "2,145.67"),
        ],
    ),
    (
        "ACME_PO_PO-77001.pdf",
        "PURCHASE ORDER",
        "PO Number: PO-77001",
        [
            ("CP-200", "Chiller Pump Assembly", "1", "4,500.00", "4,500.00"),
            ("HX-900", "Heat Exchanger", "1", "7,845.67", "7,845.67"),
        ],
    ),
    (
        "ACME_DM_DM-5101.pdf",
        "DEBIT MEMO",
        "Debit Memo No: DM-5101",
        [
            ("SV-001", "Field Service Labor", "8", "150.00", "1,200.00"),
            ("FR-002", "Freight Charge", "1", "11,145.67", "11,145.67"),
        ],
    ),
]


def _new_canvas(path):
    # invariant=1: 만들 때마다 같은 바이트가 나오게 한다(커밋 차이 방지)
    return canvas.Canvas(str(path), pagesize=letter, invariant=1)


def _document(path, title, number_line, items):
    c = _new_canvas(path)
    c.setFont("Helvetica-Bold", 20)
    c.drawString(72, 720, title)
    c.setFont("Helvetica", 11)
    c.drawString(72, 690, "ACME Cooling Systems Inc.")
    c.drawString(72, 660, number_line)
    c.drawString(72, 645, f"Date: {DATE}")
    c.drawString(72, 630, f"Bill To: {CUSTOMER}")

    y = 590
    c.setFont("Helvetica-Bold", 10)
    for x, label in [(72, "Item"), (150, "Description")]:
        c.drawString(x, y, label)
    for x, label in [(390, "Qty"), (470, "Unit Price"), (540, "Amount")]:
        c.drawRightString(x, y, label)
    c.setFont("Helvetica", 10)
    for part, desc, qty, price, amount in items:
        y -= 18
        c.drawString(72, y, part)
        c.drawString(150, y, desc)
        c.drawRightString(390, y, qty)
        c.drawRightString(470, y, price)
        c.drawRightString(540, y, amount)

    y -= 30
    c.drawString(390, y, f"Subtotal: {TOTAL}")
    c.setFont("Helvetica-Bold", 11)
    c.drawString(390, y - 18, f"Total: ${TOTAL}")
    c.showPage()
    c.save()


def _letter(path):
    c = _new_canvas(path)
    c.setFont("Helvetica", 11)
    lines = [
        "ACME Cooling Systems Inc.",
        "",
        "Dear Customer,",
        "Thank you for visiting our booth at the trade show.",
        "Our new catalog will be mailed next month.",
        "",
        "Best regards,",
        "ACME Sales Team",
    ]
    y = 720
    for line in lines:
        c.drawString(72, y, line)
        y -= 16
    c.showPage()
    c.save()


def _blank(path):
    c = _new_canvas(path)
    c.showPage()  # 글자 없는 빈 쪽 = 스캔본 흉내
    c.save()


def build(out_dir):
    out_dir = Path(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    for name, title, number_line, items in DOCUMENTS:
        _document(out_dir / name, title, number_line, items)
    _letter(out_dir / "Unknown_letter.pdf")
    _blank(out_dir / "Scan_blank.pdf")
    return sorted(out_dir.glob("*.pdf"))


def _self_check(out_dir):
    import pdfplumber

    with pdfplumber.open(Path(out_dir) / DOCUMENTS[0][0]) as pdf:
        text = pdf.pages[0].extract_text() or ""
    ok = "Invoice No" in text
    print(f"자가 확인: Invoice 에서 'Invoice No' 글자 추출 — {'성공' if ok else '실패'}")
    return ok


if __name__ == "__main__":
    import sys

    target = Path(sys.argv[1])
    for f in build(target):
        print(f"만듦: {f.name}")
    raise SystemExit(0 if _self_check(target) else 1)
