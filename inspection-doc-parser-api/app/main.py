"""
Inspection Document Parser API

Upload an inspection-style PDF and get back structured JSON: text blocks, tables
(including borderless element / rating / comment tables), AcroForm field values,
and a count of repeated banner/header lines that were stripped.

Uploaded files are written to a temporary file only for the duration of the
request and deleted immediately. Nothing is stored or logged except counts.
"""

import logging
import os
import tempfile
import time
from dataclasses import asdict
from pathlib import Path
from typing import Literal

from fastapi import FastAPI, File, HTTPException, Query, UploadFile
from fastapi.responses import FileResponse
from pypdf import PdfReader
from pypdf.errors import PdfReadError

from . import pdf_parser

MAX_BYTES = int(os.getenv("MAX_UPLOAD_MB", "10")) * 1024 * 1024
MAX_PAGES = int(os.getenv("MAX_PAGES", "30"))

BASE_DIR = Path(__file__).resolve().parent
SAMPLE_PDF = BASE_DIR.parent / "samples" / "sample_inspection_report.pdf"

log = logging.getLogger("parser-api")

app = FastAPI(
    title="Inspection Document Parser API",
    version="1.0.0",
    description=(
        "Turns dense inspection-report PDFs into structured JSON: text blocks, "
        "tables, AcroForm fields, and boilerplate removal. "
        "Demo service: do not upload confidential documents."
    ),
)


def summarize(doc: dict, seconds: float) -> dict:
    pages = doc["pages"]
    meta = doc.get("metadata") or {}
    return {
        "filename": doc["filename"],
        "pages": doc["n_pages"],
        "text_blocks": sum(len(p["blocks"]) for p in pages),
        "tables": sum(len(p["tables"]) for p in pages),
        "acroform_fields": len(meta.get("acroform_fields") or {}),
        "label_value_fields": sum(len(p["form_fields"]) for p in pages),
        "boilerplate_lines_removed": meta.get("boilerplate_lines_removed", 0),
        "parse_seconds": round(seconds, 2),
    }


def parse_path(path: Path, display_name: str, detail: str) -> dict:
    """Validate page count, run the parser, and shape the response."""
    try:
        n_pages = len(PdfReader(str(path)).pages)
    except (PdfReadError, ValueError, OSError):
        raise HTTPException(status_code=400, detail="Could not read this file as a PDF.")
    if n_pages > MAX_PAGES:
        raise HTTPException(
            status_code=413,
            detail=f"PDF has {n_pages} pages; this demo accepts up to {MAX_PAGES}.",
        )

    start = time.perf_counter()
    try:
        doc = asdict(pdf_parser.parse_pdf(str(path)))
    except Exception:
        log.exception("parse failed")
        raise HTTPException(status_code=422, detail="The parser could not process this PDF.")
    doc["filename"] = display_name

    result = {"summary": summarize(doc, time.perf_counter() - start)}
    if detail == "full":
        result["document"] = doc
    else:  # "tables": just what most callers want from an inspection form
        result["tables"] = [t for p in doc["pages"] for t in p["tables"]]
        result["acroform_fields"] = (doc.get("metadata") or {}).get("acroform_fields", {})
    return result


@app.get("/", include_in_schema=False)
def index():
    return FileResponse(BASE_DIR / "static" / "index.html")


@app.get("/health", tags=["service"])
def health():
    return {"status": "ok"}


@app.post("/parse", tags=["parse"])
def parse(
    file: UploadFile = File(..., description="A PDF, up to 10 MB and 30 pages"),
    detail: Literal["full", "tables"] = Query("full", description="full = everything; tables = tables + form fields only"),
):
    """Parse an uploaded PDF into structured JSON."""
    data = file.file.read(MAX_BYTES + 1)
    if len(data) > MAX_BYTES:
        raise HTTPException(status_code=413, detail=f"File is larger than {MAX_BYTES // (1024 * 1024)} MB.")
    if not data.startswith(b"%PDF-"):
        raise HTTPException(status_code=400, detail="Only PDF files are accepted.")

    tmp = tempfile.NamedTemporaryFile(suffix=".pdf", delete=False)
    try:
        tmp.write(data)
        tmp.close()
        name = Path(file.filename or "upload.pdf").name
        return parse_path(Path(tmp.name), name, detail)
    finally:
        os.unlink(tmp.name)  # never keep uploads


@app.get("/sample", tags=["parse"])
def parse_sample(detail: Literal["full", "tables"] = Query("full")):
    """Parse the bundled synthetic sample report, so you can try the API without a PDF."""
    return parse_path(SAMPLE_PDF, SAMPLE_PDF.name, detail)


@app.get("/sample.pdf", tags=["parse"])
def download_sample():
    """Download the synthetic sample report to see what the parser read."""
    return FileResponse(SAMPLE_PDF, media_type="application/pdf", filename=SAMPLE_PDF.name)
