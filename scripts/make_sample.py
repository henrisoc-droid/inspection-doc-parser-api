"""
make_sample.py — build samples/sample_inspection_report.pdf

A SYNTHETIC inspection report that exercises each thing the parser handles:
  * a confidentiality-style banner and header block repeated on every page
    (should be detected and stripped as boilerplate)
  * fillable AcroForm fields with values (read from the PDF's form structure)
  * a borderless Element / Rating / Comment table aligned with whitespace
  * a ruled table with grid lines
Every name, ID, and finding is invented.

Usage:  python scripts/make_sample.py
"""

from pathlib import Path

from reportlab.lib import colors
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas

OUT = Path(__file__).resolve().parent.parent / "samples" / "sample_inspection_report.pdf"
W, H = letter


def page_furniture(c, page):
    c.setFont("Helvetica-Bold", 8)
    c.drawCentredString(W / 2, H - 28, "SAMPLE DOCUMENT - SYNTHETIC DATA - NOT A REAL INSPECTION REPORT")
    c.setFont("Helvetica", 9)
    c.drawString(50, H - 46, "Structure No. 00-000-0000-00-001    Facility Carried: Sample Road over Example Creek")
    c.drawString(50, H - 58, "Inspection Type: Routine    Inspector: Inspector A (synthetic)")
    c.line(50, H - 64, W - 50, H - 64)
    c.setFont("Helvetica", 8)
    c.drawRightString(W - 50, 30, f"Page {page} of 3")


def field(c, name, label, value, x, y, width=60):
    c.setFont("Helvetica", 10)
    c.drawString(x, y + 4, label)
    c.acroForm.textfield(name=name, value=value, x=x + 150, y=y, width=width, height=16,
                         fontSize=10, borderColor=colors.grey, fillColor=colors.white)


def build():
    OUT.parent.mkdir(parents=True, exist_ok=True)
    c = canvas.Canvas(str(OUT), pagesize=letter)
    c.setTitle("Sample Inspection Report (synthetic)")

    # Page 1: summary and fillable condition-rating fields
    page_furniture(c, 1)
    c.setFont("Helvetica-Bold", 16)
    c.drawString(50, H - 100, "ROUTINE BRIDGE INSPECTION REPORT")
    c.setFont("Helvetica", 10)
    c.drawString(50, H - 120, "Inspection Date: 06/12/2025    Year Built: 1978    Spans: 3")
    c.setFont("Helvetica-Bold", 11)
    c.drawString(50, H - 150, "Condition Ratings")
    y = H - 175
    for name, label, value in [
        ("deck_rating", "Deck:", "6"),
        ("superstructure_rating", "Superstructure:", "6"),
        ("substructure_rating", "Substructure:", "5"),
        ("channel_rating", "Channel:", "6"),
        ("culvert_rating", "Culvert:", "N"),
    ]:
        field(c, name, label, value, 60, y)
        y -= 26
    c.setFont("Helvetica", 10)
    c.drawString(50, y - 10, "Special Inspection Needed: No")
    c.drawString(50, y - 24, "Load Posting: Not required")
    c.showPage()

    # Page 2: borderless table (whitespace-aligned, no grid lines)
    page_furniture(c, 2)
    c.setFont("Helvetica-Bold", 11)
    c.drawString(50, H - 100, "Element Condition Summary")
    rows = [
        ("Element / Component", "Rating", "Comment"),
        ("Deck - Concrete", "6", "Longitudinal cracking with efflorescence near Bent 2."),
        ("Beams - Prestressed", "6", "Hairline crack at midspan of Beam 3."),
        ("Bent Caps - Concrete", "5", "Spall with exposed rebar at Bent 2 cap, about 2 sq ft."),
        ("Bearings - Elastomeric", "7", "Slight bulging at Bent 2."),
        ("Channel - Riprap", "6", "Riprap partially displaced at Abutment 1."),
    ]
    y = H - 125
    for i, (a, b, cmt) in enumerate(rows):
        c.setFont("Helvetica-Bold" if i == 0 else "Helvetica", 9.5)
        c.drawString(50, y, a)
        c.drawString(215, y, b)
        c.drawString(265, y, cmt)
        y -= 16
    c.setFont("Helvetica", 10)
    c.drawString(50, y - 14, "Notes: Ratings reflect the worst condition observed for each component.")
    c.showPage()

    # Page 3: ruled table with grid lines
    page_furniture(c, 3)
    c.setFont("Helvetica-Bold", 11)
    c.drawString(50, H - 100, "Recommended Maintenance")
    table = [
        ("Item", "Location", "Priority", "Action"),
        ("1", "Bent 2 cap", "High", "Patch spall; clean and coat exposed rebar"),
        ("2", "Deck", "Medium", "Seal longitudinal cracks"),
        ("3", "Abutment 1", "Medium", "Replace displaced riprap"),
    ]
    xs = [50, 90, 190, 260, W - 50]
    top, rh = H - 115, 20
    for r, row in enumerate(table):
        yy = top - r * rh
        for k, txt in enumerate(row):
            c.setFont("Helvetica-Bold" if r == 0 else "Helvetica", 9.5)
            c.drawString(xs[k] + 4, yy - 14, txt)
    for r in range(len(table) + 1):
        c.line(xs[0], top - r * rh, xs[-1], top - r * rh)
    for x in xs:
        c.line(x, top, x, top - len(table) * rh)
    c.save()
    print(f"Wrote {OUT}")


if __name__ == "__main__":
    build()
