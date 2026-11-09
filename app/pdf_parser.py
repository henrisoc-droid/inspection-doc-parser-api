"""
pdf_parser.py (vendored from the inspection-qa-model repository;
this copy also records how many boilerplate lines were removed)

pdf_parser.py — Multi-format Document Parser (Python 3.12 compatible)
Handles PDF, PowerPoint (.pptx), and Excel (.xlsx) files.
Install: pip install pypdf python-pptx openpyxl --break-system-packages
"""

import os, re, json, logging
from pathlib import Path
from dataclasses import dataclass, asdict
from typing import List, Dict

log = logging.getLogger(__name__)
logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")

# ── Imports ──────────────────────────────────────────────
try:
    from pypdf import PdfReader
    HAS_PDF = True
except ImportError:
    HAS_PDF = False
    log.warning("pypdf not found. Install: pip install pypdf --break-system-packages")

try:
    import pdfplumber
    HAS_PDFPLUMBER = True
except ImportError:
    HAS_PDFPLUMBER = False
    log.warning("pdfplumber not found — table extraction will fall back to the "
                "whitespace-heuristic method, which is unreliable on genuinely "
                "tabular layouts. Install: pip install pdfplumber --break-system-packages")

try:
    from pptx import Presentation
    HAS_PPTX = True
except ImportError:
    HAS_PPTX = False
    log.warning("python-pptx not found. Install: pip install python-pptx --break-system-packages")

try:
    import openpyxl
    HAS_XLSX = True
except ImportError:
    HAS_XLSX = False
    log.warning("openpyxl not found. Install: pip install openpyxl --break-system-packages")


# ── Data structures ───────────────────────────────────────

@dataclass
class TextBlock:
    text: str; x0: float; y0: float; x1: float; y1: float
    font: str; size: float; bold: bool; italic: bool; color: int
    block_type: str; page: int; reading_order: int

@dataclass
class Table:
    cells: list; x0: float; y0: float; x1: float; y1: float
    page: int; n_rows: int; n_cols: int

@dataclass
class FormField:
    field_name: str; field_value: str; field_type: str
    x0: float; y0: float; x1: float; y1: float; page: int

@dataclass
class ParsedPage:
    page_num: int; width: float; height: float
    blocks: list; tables: list; form_fields: list; raw_text: str

@dataclass
class ParsedDocument:
    filename: str; n_pages: int; pages: list
    metadata: dict; doc_type: str


# ── Classifier ────────────────────────────────────────────

def classify_block(text: str, line_num: int) -> str:
    text = text.strip()
    if not text: return "empty"
    if line_num < 5 and len(text) < 100: return "header"
    if text.endswith(":") and len(text) < 60: return "label"
    stripped = text.replace(".", "").replace("-", "").replace(" ", "")
    if stripped.isdigit() and len(text) < 10: return "numeric_field"
    return "text"


def text_to_blocks(lines: List[str], page_num: int,
                   width: float = 612, height: float = 792) -> List[TextBlock]:
    """Convert a list of text lines into TextBlock objects."""
    blocks = []
    for ln, line in enumerate(lines):
        t = line.strip()
        if not t:
            continue
        bt = classify_block(t, ln)
        y0 = (ln / max(len(lines), 1)) * height
        blocks.append(TextBlock(
            text=t, x0=0.0, y0=y0, x1=float(width), y1=y0 + 14,
            font="", size=10.0, bold=(bt == "header"),
            italic=False, color=0,
            block_type=bt, page=page_num, reading_order=ln
        ))
    return blocks


def detect_tables(lines: List[str], page_num: int) -> List[Table]:
    """Detect tables by finding lines with consistent multi-space column alignment."""
    tables = []
    i = 0
    while i < len(lines):
        line = lines[i].strip()
        if re.search(r'\s{3,}', line) and len(line) > 20:
            table_lines = []
            j = i
            while j < len(lines) and (not lines[j].strip() or re.search(r'\s{2,}', lines[j])):
                if lines[j].strip():
                    table_lines.append(lines[j])
                j += 1
            if len(table_lines) >= 2:
                cells = []
                for r_idx, tline in enumerate(table_lines):
                    cols = re.split(r'\s{2,}', tline.strip())
                    for c_idx, col_text in enumerate(cols):
                        if col_text.strip():
                            from dataclasses import fields
                            cells.append({
                                "row": r_idx, "col": c_idx,
                                "text": col_text.strip(),
                                "x0": 0, "y0": float(i + r_idx),
                                "x1": 100, "y1": float(i + r_idx + 1),
                                "page": page_num
                            })
                if cells:
                    n_rows = len(table_lines)
                    n_cols = max(c["col"] for c in cells) + 1
                    tables.append(Table(
                        cells=cells, x0=0, y0=float(i),
                        x1=100, y1=float(j),
                        page=page_num, n_rows=n_rows, n_cols=n_cols
                    ))
                i = j
                continue
        i += 1
    return tables


def extract_form_fields(text: str, page_num: int) -> List[FormField]:
    """Extract label: value pairs as form fields."""
    fields = []
    pattern = re.compile(r'^([A-Za-z][A-Za-z\s/\(\)]{2,40}):\s*(.{0,100})$', re.MULTILINE)
    for m in pattern.finditer(text):
        label = m.group(1).strip()
        value = m.group(2).strip()
        if label and len(label) < 50:
            fields.append(FormField(
                field_name=label, field_value=value, field_type="text",
                x0=0, y0=0, x1=100, y1=10, page=page_num
            ))
    return fields[:50]


# ── PDF Parser ────────────────────────────────────────────

def extract_acroform_fields(reader) -> Dict[str, str]:
    """Extract true AcroForm field name/value pairs directly from the
    PDF's interactive form structure, when present. This reads the
    PDF's own internal field structure rather than guessing labels
    from linearized extracted text — no risk of adjacent fields on a
    dense multi-column form bleeding into each other, which regex-based
    label:value matching is prone to on cramped inspection forms."""
    try:
        fields = reader.get_fields()
    except Exception as e:
        log.warning(f"Could not read AcroForm fields: {e}")
        return {}
    if not fields:
        return {}
    result = {}
    for name, f in fields.items():
        try:
            field_type = f.get("/FT")
        except Exception:
            field_type = None
        if field_type == "/Sig":
            continue  # digital signature field — value is raw binary
                      # certificate data, never usable as text
        try:
            value = f.get("/V")
        except Exception:
            value = None
        if value is None:
            continue
        value_str = str(value).strip()
        name_str = str(name).strip()
        if not name_str or not value_str:
            continue
        # Reject values that aren't clean, mostly-printable text — a
        # cheap guard against any other binary/encoded field slipping
        # through besides signatures.
        printable = sum(1 for c in value_str if c.isprintable())
        if len(value_str) > 300 or printable / max(1, len(value_str)) < 0.9:
            continue
        result[name_str] = value_str
    return result


_DRAWING_NAME_MARKERS = [
    "plans", "origplans", "asbuilt", "as-built", "drawing",
    "sketch", "crosssection", "cross-section", "photos", "photolog",
]

# A page with this many vector objects is a CAD/plan drawing, not a
# data table. pdfplumber's table inference walks every line/curve/rect
# to find cell boundaries, so a dense plan sheet can effectively hang
# the process for hours — this is a hard guard against that.
_MAX_VECTOR_OBJECTS = 3000


def _looks_like_drawing(pdf_path: str) -> bool:
    name = Path(pdf_path).stem.lower()
    return any(marker in name for marker in _DRAWING_NAME_MARKERS)


# ── Borderless tables ─────────────────────────────────────
# Inspection forms often lay out Element / Rating / Comment tables with
# whitespace instead of ruling lines. pdfplumber's "text" table strategy
# treats any page of left-aligned text as one big grid, slicing words in
# half and pulling banners into the "table". Instead, this detector only
# reports a table where several consecutive lines share the same column
# start positions.

_PHRASE_GAP = 15.0      # pt: words closer than this belong to one cell
_COL_TOLERANCE = 6.0    # pt: column starts within this snap together
_MIN_TABLE_ROWS = 3


def _page_lines(page) -> List[List[dict]]:
    """Group a page's words into visual lines, sorted top to bottom."""
    words = page.extract_words(x_tolerance=2, y_tolerance=3, keep_blank_chars=False)
    lines: List[List[dict]] = []
    for w in sorted(words, key=lambda w: (round(w["top"]), w["x0"])):
        if lines and abs(lines[-1][0]["top"] - w["top"]) <= 3:
            lines[-1].append(w)
        else:
            lines.append([w])
    return [sorted(line, key=lambda w: w["x0"]) for line in lines]


def _phrases(line: List[dict]) -> List[dict]:
    """Merge words separated by small gaps into phrases (candidate cells)."""
    out = []
    for w in line:
        if out and w["x0"] - out[-1]["x1"] < _PHRASE_GAP:
            out[-1]["text"] += " " + w["text"]
            out[-1]["x1"] = w["x1"]
        else:
            out.append({"text": w["text"], "x0": w["x0"], "x1": w["x1"], "top": w["top"]})
    return out


def _snap(x: float, anchors: List[float]) -> int:
    for i, a in enumerate(anchors):
        if abs(x - a) <= _COL_TOLERANCE:
            return i
    return -1


def extract_borderless_tables(page) -> List[List[List[str]]]:
    """Return tables as lists of rows (lists of cell strings), same shape
    as pdfplumber's extract_tables(), found by column alignment."""
    rows = [_phrases(line) for line in _page_lines(page)]
    tables, i = [], 0
    while i < len(rows):
        if len(rows[i]) < 2:
            i += 1
            continue
        anchors = [ph["x0"] for ph in rows[i]]
        run = [rows[i]]
        j = i + 1
        while j < len(rows):
            row = rows[j]
            gap = row[0]["top"] - run[-1][0]["top"]
            if gap > 30 or not row:
                break
            cols = [_snap(ph["x0"], anchors) for ph in row]
            if -1 in cols:
                break
            run.append(row)
            j += 1
        while len(run) > 1 and len(run[-1]) < 2:   # a trailing single-cell line
            run.pop(); j -= 1                        # (e.g. a note) isn't a row
        multi = sum(1 for r in run if len(r) >= 2)
        if len(run) >= _MIN_TABLE_ROWS and multi >= _MIN_TABLE_ROWS:
            table = []
            for row in run:
                cells = [""] * len(anchors)
                for ph in row:
                    k = _snap(ph["x0"], anchors)
                    cells[k] = (cells[k] + " " + ph["text"]).strip()
                table.append(cells)
            tables.append(table)
            i = j
        else:
            i += 1
    return tables


def extract_tables_pdfplumber(pdf_path: str) -> Dict[int, List[Table]]:
    """Extract tables per page using pdfplumber's geometry-based table
    detection (reads actual character positions/line geometry from the
    PDF), far more reliable than the whitespace-heuristic detect_tables()
    above for genuinely tabular layouts — e.g. an Element/Component/
    Rating/Comment inspection table with wrapped multi-line cells,
    which the whitespace approach cannot reconstruct at all."""
    result: Dict[int, List[Table]] = {}
    if not HAS_PDFPLUMBER:
        return result
    if _looks_like_drawing(pdf_path):
        log.info(f"Skipping table extraction (drawing/plan sheet): {Path(pdf_path).name}")
        return result
    try:
        with pdfplumber.open(pdf_path) as pdf:
            for i, page in enumerate(pdf.pages):
                try:
                    n_vectors = (len(page.lines) + len(page.curves) + len(page.rects))
                    if n_vectors > _MAX_VECTOR_OBJECTS:
                        log.info(f"  page {i+1}: skipping table extraction "
                                 f"({n_vectors} vector objects — drawing, not a table)")
                        continue
                except Exception:
                    pass
                try:
                    raw_tables = page.extract_tables()
                except Exception:
                    raw_tables = []
                # Fall back to text-alignment-based detection when the
                # default (ruling-line-based) strategy finds nothing.
                # TxDOT inspection tables — Element/Component | Min. |
                # Rating | Comment — are borderless, using whitespace
                # alignment rather than drawn lines, so the default
                # strategy returns them as flat text and the column
                # structure is lost.
                if not raw_tables:
                    try:
                        raw_tables = extract_borderless_tables(page)
                    except Exception:
                        raw_tables = []
                page_tables = []
                for rt in raw_tables:
                    if not rt:
                        continue
                    n_rows = len(rt)
                    n_cols = max((len(row) for row in rt), default=0)
                    # Skip trivially small "tables" — pdfplumber sometimes
                    # returns 1-2 row layout artifacts (a boxed header, a
                    # signature line) that aren't real data tables and
                    # would only add noise downstream.
                    if n_rows < 2 or n_cols < 2:
                        continue
                    cells = []
                    for r_idx, row in enumerate(rt):
                        for c_idx, val in enumerate(row):
                            text = (val or "").strip()
                            if text:
                                cells.append({
                                    "row": r_idx, "col": c_idx, "text": text,
                                    "x0": 0.0, "y0": float(r_idx),
                                    "x1": 100.0, "y1": float(r_idx + 1),
                                    "page": i
                                })
                    if cells:
                        page_tables.append(Table(
                            cells=cells, x0=0, y0=0, x1=100, y1=float(n_rows),
                            page=i, n_rows=n_rows, n_cols=n_cols
                        ))
                if page_tables:
                    result[i] = page_tables
    except Exception as e:
        log.warning(f"pdfplumber table extraction failed for {pdf_path}: {e}")
    return result


_BOILERPLATE_MIN_PAGE_FRACTION = 0.4  # a line repeated on this fraction of
                                       # pages (verbatim, after whitespace
                                       # normalization) is treated as a
                                       # repeated header/footer/banner and
                                       # stripped from every page. Agency
                                       # inspection PDFs reprint a
                                       # confidentiality banner and a Bridge
                                       # Number/Facility Carried/Inspector
                                       # block on nearly every page — left
                                       # in, this text is extracted once
                                       # per page and floods downstream
                                       # training corpora with duplicated
                                       # boilerplate instead of field-note
                                       # content.


def _normalize_line(line: str) -> str:
    """Collapse whitespace so trivially-reflowed repeats of the same
    boilerplate line (e.g. a date or page number shifting spacing) still
    match as the same line."""
    return re.sub(r"\s+", " ", line.strip())


def _find_repeated_boilerplate(pages_lines: List[List[str]]) -> set:
    """Identify lines that recur, near-verbatim, across a large fraction
    of a document's pages. Short lines (page numbers, single words) are
    excluded — they're common by coincidence, not because they're a
    repeated banner, and stripping them would delete real short field
    values too."""
    n_pages = len(pages_lines)
    if n_pages < 3:
        return set()  # not enough pages for "repeated across the doc" to
                       # be a meaningful signal — leave short documents alone
    counts: Dict[str, int] = {}
    for lines in pages_lines:
        seen_this_page = set()
        for line in lines:
            norm = _normalize_line(line)
            if len(norm) < 15:
                continue
            if norm in seen_this_page:
                continue
            seen_this_page.add(norm)
            counts[norm] = counts.get(norm, 0) + 1
    threshold = max(2, int(n_pages * _BOILERPLATE_MIN_PAGE_FRACTION))
    return {norm for norm, c in counts.items() if c >= threshold}


def _strip_boilerplate(lines: List[str], boilerplate: set) -> List[str]:
    return [ln for ln in lines if _normalize_line(ln) not in boilerplate]


def parse_pdf(pdf_path: str) -> ParsedDocument:
    if not HAS_PDF:
        raise ImportError("pypdf required: pip install pypdf --break-system-packages")
    path = Path(pdf_path)
    log.info(f"Parsing PDF: {path.name}")
    reader = PdfReader(str(path))
    meta = reader.metadata or {}
    metadata = {"title": str(meta.get("/Title", "")), "pages": len(reader.pages)}
    metadata["acroform_fields"] = extract_acroform_fields(reader)

    # Pass 1: extract raw text per page. Prefer pdfplumber when available —
    # its word-position-sorted extraction tracks visual reading order far
    # more reliably than pypdf's content-stream order on dense,
    # multi-column forms with checkboxes, which is where labels like
    # "Upstream:" and "Special Inspection Needed:" were bleeding into
    # each other with no separator in the pypdf-only extraction.
    plumber_page_texts: Dict[int, str] = {}
    if HAS_PDFPLUMBER and not _looks_like_drawing(pdf_path):
        try:
            with pdfplumber.open(str(path)) as pdf:
                for i, pg in enumerate(pdf.pages):
                    try:
                        plumber_page_texts[i] = pg.extract_text(
                            x_tolerance=2, y_tolerance=3, layout=False
                        ) or ""
                    except Exception:
                        pass
        except Exception as e:
            log.warning(f"pdfplumber text pass failed for {pdf_path}: {e}")

    raw_texts: List[str] = []
    dims: List[tuple] = []
    for page_num, page in enumerate(reader.pages):
        text = plumber_page_texts.get(page_num)
        if not text:
            try:
                text = page.extract_text() or ""
            except Exception:
                text = ""
        raw_texts.append(text)
        try:
            dims.append((float(page.mediabox.width), float(page.mediabox.height)))
        except Exception:
            dims.append((612.0, 792.0))

    # Pass 2: find lines repeated across most pages (banners, repeated
    # header blocks) and strip them before anything downstream sees them.
    pages_lines = [t.split("\n") for t in raw_texts]
    boilerplate = _find_repeated_boilerplate(pages_lines)
    metadata["boilerplate_lines_removed"] = len(boilerplate)
    if boilerplate:
        log.info(f"  stripping {len(boilerplate)} repeated boilerplate "
                 f"line(s) found across pages")

    parsed_pages = []
    for page_num, lines in enumerate(pages_lines):
        width, height = dims[page_num]
        clean_lines = _strip_boilerplate(lines, boilerplate)
        raw_text = "\n".join(clean_lines)
        blocks = text_to_blocks(clean_lines, page_num, width, height)
        tables = detect_tables(clean_lines, page_num)
        form_fields = extract_form_fields(raw_text, page_num)
        parsed_pages.append(ParsedPage(
            page_num=page_num, width=width, height=height,
            blocks=blocks, tables=tables,
            form_fields=form_fields, raw_text=raw_text
        ))

    # Prefer pdfplumber's geometry-based table extraction when
    # available — the whitespace-heuristic detect_tables() above
    # cannot reliably reconstruct genuinely tabular layouts (e.g. an
    # Element/Component/Rating/Comment inspection table with wrapped
    # multi-line comment cells). Only override a page's tables when
    # pdfplumber actually found something for that page.
    if HAS_PDFPLUMBER:
        plumber_tables = extract_tables_pdfplumber(pdf_path)
        for pg in parsed_pages:
            if plumber_tables.get(pg.page_num):
                pg.tables = plumber_tables[pg.page_num]

    return ParsedDocument(
        filename=path.name, n_pages=len(parsed_pages),
        pages=parsed_pages, metadata=metadata, doc_type="pdf"
    )


# ── PowerPoint Parser ─────────────────────────────────────

def parse_pptx(pptx_path: str) -> ParsedDocument:
    if not HAS_PPTX:
        raise ImportError("python-pptx required: pip install python-pptx --break-system-packages")
    path = Path(pptx_path)
    log.info(f"Parsing PPTX: {path.name}")
    prs = Presentation(str(path))
    parsed_pages = []
    for slide_num, slide in enumerate(prs.slides):
        texts = []
        for shape in slide.shapes:
            if hasattr(shape, "text") and shape.text.strip():
                texts.append(shape.text.strip())
            # Extract table content from slides
            if shape.has_table:
                for row in shape.table.rows:
                    row_text = " | ".join(
                        cell.text.strip() for cell in row.cells if cell.text.strip()
                    )
                    if row_text:
                        texts.append(row_text)
        raw_text = "\n".join(texts)
        lines = raw_text.split("\n")
        blocks = text_to_blocks(lines, slide_num, 960, 540)
        tables = detect_tables(lines, slide_num)
        parsed_pages.append(ParsedPage(
            page_num=slide_num, width=960.0, height=540.0,
            blocks=blocks, tables=tables,
            form_fields=[], raw_text=raw_text
        ))
    metadata = {"title": path.stem, "pages": len(prs.slides)}
    return ParsedDocument(
        filename=path.name, n_pages=len(parsed_pages),
        pages=parsed_pages, metadata=metadata, doc_type="pptx"
    )


# ── Excel Parser ──────────────────────────────────────────

def parse_xlsx(xlsx_path: str) -> ParsedDocument:
    if not HAS_XLSX:
        raise ImportError("openpyxl required: pip install openpyxl --break-system-packages")
    path = Path(xlsx_path)
    log.info(f"Parsing XLSX: {path.name}")
    wb = openpyxl.load_workbook(str(xlsx_path), read_only=True, data_only=True)
    parsed_pages = []
    for sheet_num, sheet in enumerate(wb.worksheets):
        rows_text = []
        cells_data = []
        for r_idx, row in enumerate(sheet.iter_rows(values_only=True)):
            row_vals = [str(cell) if cell is not None else "" for cell in row]
            row_text = " | ".join(v for v in row_vals if v.strip())
            if row_text.strip():
                rows_text.append(row_text)
                for c_idx, val in enumerate(row_vals):
                    if val.strip():
                        cells_data.append({
                            "row": r_idx, "col": c_idx,
                            "text": val.strip(),
                            "x0": float(c_idx * 100),
                            "y0": float(r_idx * 20),
                            "x1": float((c_idx + 1) * 100),
                            "y1": float((r_idx + 1) * 20),
                            "page": sheet_num
                        })
        raw_text = "\n".join(rows_text)
        lines = raw_text.split("\n")
        blocks = text_to_blocks(lines, sheet_num, 1200, 800)
        table = None
        if cells_data:
            n_rows = max(c["row"] for c in cells_data) + 1
            n_cols = max(c["col"] for c in cells_data) + 1
            table = Table(
                cells=cells_data, x0=0, y0=0,
                x1=1200, y1=800,
                page=sheet_num, n_rows=n_rows, n_cols=n_cols
            )
        parsed_pages.append(ParsedPage(
            page_num=sheet_num, width=1200.0, height=800.0,
            blocks=blocks,
            tables=[table] if table else [],
            form_fields=[], raw_text=raw_text
        ))
    wb.close()
    metadata = {"title": path.stem, "pages": len(wb.worksheets)}
    return ParsedDocument(
        filename=path.name, n_pages=len(parsed_pages),
        pages=parsed_pages, metadata=metadata, doc_type="xlsx"
    )


# ── Dispatcher ────────────────────────────────────────────

def parse_document(file_path: str) -> ParsedDocument:
    """Route to the appropriate parser based on file extension."""
    ext = Path(file_path).suffix.lower()
    if ext == ".pdf":
        return parse_pdf(file_path)
    elif ext in (".pptx", ".ppt"):
        return parse_pptx(file_path)
    elif ext in (".xlsx", ".xls"):
        return parse_xlsx(file_path)
    else:
        raise ValueError(f"Unsupported file type: {ext}")


def save_parsed(doc: ParsedDocument, output_path: str):
    with open(output_path, "w", encoding="utf-8") as f:
        json.dump(asdict(doc), f, indent=2)


def batch_parse(input_dir: str, output_dir: str, limit: int = 0,
                 top_level_only: bool = False):
    input_dir  = Path(input_dir)
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    # Supported extensions
    extensions = (".pdf", ".pptx", ".ppt", ".xlsx", ".xls")
    if top_level_only:
        # Only files directly inside each bridge folder — skips
        # Parts/ and Reference/ subfolders entirely. The pairing logic
        # uses non-recursive glob, so those files are parsed and then
        # never read; skipping them saves the bulk of the runtime.
        candidates = list(input_dir.glob("*/*")) + list(input_dir.glob("*"))
    else:
        candidates = list(input_dir.rglob("*"))
    files = sorted(
        {f for f in candidates
         if f.suffix.lower() in extensions and f.is_file()}
    )
    if limit:
        files = files[:limit]

    log.info(f"Found {len(files)} documents in {input_dir}"
             f"{' (top-level only)' if top_level_only else ''}")
    pdfs  = sum(1 for f in files if f.suffix.lower() == ".pdf")
    pptxs = sum(1 for f in files if f.suffix.lower() in (".pptx", ".ppt"))
    xlsxs = sum(1 for f in files if f.suffix.lower() in (".xlsx", ".xls"))
    log.info(f"  PDF: {pdfs}  PPTX: {pptxs}  XLSX: {xlsxs}")

    success, failed, skipped = 0, 0, 0

    for file_path in files:
        rel      = file_path.relative_to(input_dir)
        out_dir  = output_dir / rel.parent
        out_dir.mkdir(parents=True, exist_ok=True)
        out_path = out_dir / (file_path.stem + ".json")

        if out_path.exists():
            log.info(f"Skipping (already parsed): {file_path.name}")
            skipped += 1
            continue
        try:
            parsed = parse_document(str(file_path))
            save_parsed(parsed, str(out_path))
            nb = sum(len(p.blocks) for p in parsed.pages)
            nt = sum(len(p.tables) for p in parsed.pages)
            log.info(f"OK  [{parsed.doc_type.upper()}] {file_path.name} "
                     f"— {parsed.n_pages}p {nb} blocks {nt} tables")
            success += 1
        except Exception as e:
            log.error(f"FAIL {file_path.name}: {e}")
            failed += 1

    log.info(f"Done — {success} parsed, {failed} failed, {skipped} skipped")
    return success, failed


# ── CLI ───────────────────────────────────────────────────

if __name__ == "__main__":
    import argparse
    p = argparse.ArgumentParser(description="Parse PDF/PPTX/XLSX into structured JSON")
    p.add_argument("--pdf-dir",    default="./data/raw_pdfs")
    p.add_argument("--output-dir", default="./data/processed")
    p.add_argument("--limit",      type=int, default=0)
    p.add_argument("--top-level-only", action="store_true",
                   help="Only parse PDFs directly inside each bridge folder, "
                        "skipping Parts/ and Reference/ subfolders (which the "
                        "pairing logic never reads).")
    p.add_argument("--single",     default="")
    args = p.parse_args()

    if args.single:
        doc = parse_document(args.single)
        out = args.single.rsplit(".", 1)[0] + "_parsed.json"
        save_parsed(doc, out)
        print(f"Saved: {out}")
        for pg in doc.pages:
            print(f"  Page {pg.page_num}: {len(pg.blocks)} blocks "
                  f"{len(pg.tables)} tables")
    else:
        batch_parse(args.pdf_dir, args.output_dir, args.limit, args.top_level_only)
