"""API tests. Run with:  pytest"""

import io
from pathlib import Path

from fastapi.testclient import TestClient
from pypdf import PdfWriter

from app import main
from app.main import app

client = TestClient(app)
SAMPLE = Path(__file__).resolve().parent.parent / "samples" / "sample_inspection_report.pdf"


def blank_pdf(pages: int) -> bytes:
    w = PdfWriter()
    for _ in range(pages):
        w.add_blank_page(width=612, height=792)
    buf = io.BytesIO()
    w.write(buf)
    return buf.getvalue()


def table_rows(table):
    rows = {}
    for c in table["cells"]:
        rows.setdefault(c["row"], {})[c["col"]] = c["text"]
    return [list(r.values()) for _, r in sorted(rows.items())]


def test_health():
    assert client.get("/health").json() == {"status": "ok"}


def test_index_page():
    r = client.get("/")
    assert r.status_code == 200 and "Inspection Document Parser" in r.text


def test_sample_summary():
    s = client.get("/sample").json()["summary"]
    assert s["pages"] == 3
    assert s["tables"] == 2
    assert s["acroform_fields"] == 5
    assert s["boilerplate_lines_removed"] >= 3   # banner + header block repeated on every page


def test_form_fields_read_from_pdf_structure():
    fields = client.get("/sample?detail=tables").json()["acroform_fields"]
    assert fields["substructure_rating"] == "5"
    assert fields["culvert_rating"] == "N"


def test_borderless_table_is_reconstructed():
    tables = client.get("/sample?detail=tables").json()["tables"]
    borderless = next(t for t in tables if t["page"] == 1)   # pages are 0-indexed
    rows = table_rows(borderless)
    assert rows[0] == ["Element / Component", "Rating", "Comment"]
    assert rows[3][:2] == ["Bent Caps - Concrete", "5"]
    assert len(rows) == 6                                    # header + 5 elements, no trailing note
    assert not any("SAMPLE DOCUMENT" in cell for row in rows for cell in row)


def test_ruled_table():
    tables = client.get("/sample?detail=tables").json()["tables"]
    rows = table_rows(next(t for t in tables if t["page"] == 2))
    assert rows[0] == ["Item", "Location", "Priority", "Action"]
    assert len(rows) == 4


def test_banner_stripped_from_text():
    doc = client.get("/sample").json()["document"]
    text = " ".join(p["raw_text"] for p in doc["pages"])
    assert "SAMPLE DOCUMENT - SYNTHETIC DATA" not in text


def test_upload_matches_sample():
    with open(SAMPLE, "rb") as f:
        r = client.post("/parse", files={"file": ("report.pdf", f, "application/pdf")})
    assert r.status_code == 200
    assert r.json()["summary"]["filename"] == "report.pdf"
    assert r.json()["summary"]["tables"] == 2


def test_rejects_non_pdf():
    r = client.post("/parse", files={"file": ("notes.txt", b"hello", "text/plain")})
    assert r.status_code == 400


def test_rejects_corrupt_pdf():
    r = client.post("/parse", files={"file": ("bad.pdf", b"%PDF-1.7 not really", "application/pdf")})
    assert r.status_code == 400


def test_rejects_oversized_file(monkeypatch):
    monkeypatch.setattr(main, "MAX_BYTES", 1000)
    r = client.post("/parse", files={"file": ("big.pdf", b"%PDF-" + b"0" * 2000, "application/pdf")})
    assert r.status_code == 413


def test_rejects_too_many_pages(monkeypatch):
    monkeypatch.setattr(main, "MAX_PAGES", 2)
    r = client.post("/parse", files={"file": ("long.pdf", blank_pdf(3), "application/pdf")})
    assert r.status_code == 413


def test_uploads_are_not_kept(tmp_path, monkeypatch):
    monkeypatch.setattr(main.tempfile, "tempdir", str(tmp_path))
    with open(SAMPLE, "rb") as f:
        client.post("/parse", files={"file": ("report.pdf", f, "application/pdf")})
    assert list(tmp_path.iterdir()) == []
