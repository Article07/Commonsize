"""
Reads financial-statement PDFs -- digital OR scanned -- into structured rows.

Pipeline: pages -> text lines (embedded text if the page has any, otherwise
Tesseract OCR with rotation handling) -> page type (Balance Sheet / P&L /
Cash Flow / Notes) -> rows of (label, note ref, values per FY) -> built-in
sanity checks (note items vs the note's own Total row).

Rule-based only (no LLM). Everything uncertain is FLAGGED on the row rather
than silently repaired or guessed, because OCR on financial tables is not
fully reliable: a dropped row, a lost decimal point or a misread digit can
all look plausible. The Streamlit review table and the balance check are the
backstop; this module's job is to extract what it can and say clearly what
it is unsure about.

The heading attached to each row is whatever the PDF itself presents it
under (the Note title for note items; the nearest section heading such as
"Current Liabilities" for face-statement lines). That is what lets a
line item that isn't in our taxonomy be placed "exactly as the PDF
classifies it".
"""

import os
import re
import shutil
from dataclasses import dataclass, field
from pathlib import Path

import pypdfium2 as pdfium

OCR_SCALE = float(os.environ.get("OCR_DPI", "300")) / 72  # render resolution for OCR

_DEFAULT_TESSERACT_PATHS = [
    r"C:\Program Files\Tesseract-OCR\tesseract.exe",
    r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe",
]


class OcrUnavailableError(RuntimeError):
    pass


@dataclass
class PageText:
    page: int                 # 1-based
    lines: list               # list[str]
    source: str               # "text" | "ocr"
    ocr_confidence: float | None = None
    rotation: int = 0
    page_type: str = "OTHER"  # BS | P&L | CF | NOTES | OTHER


@dataclass
class ExtractedRow:
    page: int
    statement: str            # "BS" | "P&L" | "CF" | "NOTE"
    heading: str              # heading/note title exactly as the PDF presents it
    note_no: str              # for NOTE rows: its own note number; for face rows: ""
    note_ref: str             # for face rows: the note number cited on the line
    label: str
    values: dict              # {"FY25": 6000.0, ...} in the document's own units
    raw_text: str
    is_total: bool = False
    flags: list = field(default_factory=list)


@dataclass
class ExtractionResult:
    rows: list
    fy_columns: list          # ordered as they appear on the face statements, e.g. ["FY25", "FY24"]
    unit_multiplier: float    # converts document units to rupees (1000 for Rs.'000)
    unit_label: str
    pages: list               # list[PageText]
    checks: list              # human-readable check results / warnings
    warnings: list
    whole_units: bool = False # figures printed without decimals (e.g. whole Rs. '000): rounding is a whole unit


# --------------------------------------------------------------------------
# Page text acquisition (embedded text, else OCR)
# --------------------------------------------------------------------------

def _configure_tesseract():
    import pytesseract

    cmd = os.environ.get("TESSERACT_CMD") or shutil.which("tesseract")
    if not cmd:
        for candidate in _DEFAULT_TESSERACT_PATHS:
            if Path(candidate).exists():
                cmd = candidate
                break
    if not cmd:
        raise OcrUnavailableError(
            "This PDF has scanned (image-only) pages but Tesseract OCR is not installed "
            "or not on PATH. Set the TESSERACT_CMD environment variable or install Tesseract."
        )
    pytesseract.pytesseract.tesseract_cmd = cmd
    return pytesseract


# text some PDF printers stamp on every page; on a scanned page it would otherwise pass for a text layer
_WATERMARK = re.compile(r"pdf\s*created\s*with|pdffactory|trial\s*version|unregistered|evaluation\s*(?:copy|version)|"
                        r"created\s*(?:by|with)\s*\S*\s*(?:pdf|converter)|www\.\S*pdf\S*\.com|camscanner|scanned\s*(?:by|with)",
                        re.I)


def _embedded_lines(pdf_page):
    """Text lines from the page's embedded text layer, grouped by vertical position (printer watermarks removed)."""
    textpage = pdf_page.get_textpage()
    text = textpage.get_text_range() or ""
    return [ln.strip() for ln in text.replace("\r", "\n").split("\n") if ln.strip() and not _WATERMARK.search(ln)]


def _ocr_image_lines(pytesseract, image):
    """OCR one PIL image -> (lines, mean_confidence)."""
    from pytesseract import Output

    data = pytesseract.image_to_data(
        image, config="--psm 6 -c preserve_interword_spaces=1", output_type=Output.DICT
    )
    groups = {}
    confs = []
    for i, word in enumerate(data["text"]):
        word = (word or "").strip()
        if not word:
            continue
        conf = float(data["conf"][i])
        if conf >= 0:
            confs.append(conf)
        key = (data["block_num"][i], data["par_num"][i], data["line_num"][i])
        groups.setdefault(key, []).append((data["left"][i], word))
    lines = [
        " ".join(w for _, w in sorted(words))
        for _, words in sorted(groups.items())
    ]
    mean_conf = sum(confs) / len(confs) if confs else 0.0
    return lines, mean_conf


# Tesseract's Linux build runs OpenMP threads; inside a container with a fraction of a CPU those threads fight
# over the quota and OCR becomes many times slower. One thread per Tesseract process is much faster there.
os.environ.setdefault("OMP_THREAD_LIMIT", "1")

_TRIAGE_SCALE = 150 / 72      # quick low-resolution look: orientation + page heading
# pages that are not financial statements: auditor's report, CARO annexure, tax / filing covers
_AUDITOR_PAGE = re.compile(r"independent\s+auditor|auditor.?s?\s+report|companies\s*\(\s*auditor|annexure\s*['\"]?\s*[ab]?\s*['\"]?\s*"
                           r"(?:to|referred)|basis\s+for\s+opinion|key\s+audit\s+matters", re.I)
_OTHER_PAGE = re.compile(r"chartered\s+accountants|assessment\s+year|acknowledg|computation\s+of\s+(?:total\s+)?income|"
                         r"directors.?\s*report|board.?s\s+report|\bcaro\b", re.I)
_STATEMENT_PAGE = re.compile(r"balance\s*sheet|profit\s*(?:and|&)\s*loss|notes?\s+to|forming\s+part|cash\s*flow|schedule",
                             re.I)


def _usable_cpus():
    """CPUs this process may really use (a container's CPU quota, not the host's core count)."""
    try:
        quota, period = open("/sys/fs/cgroup/cpu.max").read().split()[:2]
        if quota != "max":
            return max(1, int(int(quota) / int(period)))
    except Exception:
        pass
    try:
        quota = int(open("/sys/fs/cgroup/cpu/cpu.cfs_quota_us").read())
        period = int(open("/sys/fs/cgroup/cpu/cpu.cfs_period_us").read())
        if quota > 0:
            return max(1, quota // period)
    except Exception:
        pass
    try:
        return max(1, len(os.sched_getaffinity(0)))
    except Exception:
        return max(1, os.cpu_count() or 1)


def _orientation(pytesseract, small):
    """Counter-clockwise degrees that turn the page upright (Tesseract's orientation check), or None."""
    try:
        osd = pytesseract.image_to_osd(small, config="--psm 0 -c min_characters_to_try=10")
    except Exception:
        return None
    rotate = re.search(r"Rotate:\s*(\d+)", osd)
    return (360 - int(rotate.group(1))) % 360 if rotate else None


def _heading_text(pytesseract, small):
    """Text of the top part of the page, read at low resolution."""
    top = small.crop((0, 0, small.width, int(small.height * 0.3)))
    try:
        return pytesseract.image_to_string(top, config="--psm 6")
    except Exception:
        return ""


def _skip_reason(heading):
    if _AUDITOR_PAGE.search(heading):
        return "auditor's report"
    if _OTHER_PAGE.search(heading) and not _STATEMENT_PAGE.search(heading):
        return "not a financial statement"
    return None


def _good(lines, conf):
    return conf >= 60 and len(lines) >= 4


def _ocr_page(pytesseract, base, small):
    """
    OCR a page: upright first (most pages). Only if that reads poorly, ask Tesseract which way the page is
    turned and read it once that way; if it cannot tell, try the other rotations, stopping at a good read.
    """
    lines, conf = _ocr_image_lines(pytesseract, base)
    best = (lines, conf, 0)
    if len(lines) >= 4 and conf < 70:
        # a poor, speckled scan: a light median filter often reads it much better; keep it only if clearly so
        from PIL import ImageFilter
        d_lines, d_conf = _ocr_image_lines(pytesseract, base.filter(ImageFilter.MedianFilter(3)))
        if d_conf >= conf + 5:
            best = (d_lines, d_conf, 0)
    if _good(*best[:2]):
        return best
    angle = _orientation(pytesseract, small)
    if angle == 0:
        return best          # upright, just a faint or busy page: other rotations would not read better
    order = ([angle] if angle is not None else []) + [a for a in (90, 270, 180) if a != angle]
    for a in order:
        r_lines, r_conf = _ocr_image_lines(pytesseract, base.rotate(a, expand=True))
        if r_conf > best[1] + 5:
            best = (r_lines, r_conf, a)
        if _good(*best[:2]) and best[1] >= 70:
            break
    return best


def _read_scanned_page(pytesseract, base, small, triage):
    """-> (lines, confidence, rotation, skipped_reason, heading_is_statement)"""
    if triage:
        heading = _heading_text(pytesseract, small)
        reason = _skip_reason(heading)
        if reason:
            return [], None, 0, reason, False
    lines, conf, rot = _ocr_page(pytesseract, base, small)
    return lines, conf, rot, None, None


def read_pages(pdf_source, progress=None, ocr=True, triage=True):
    """
    pdf_source: path or bytes. progress: optional callable(done, total, message).
    Returns list[PageText]. Pages with an embedded text layer are read directly; image-only pages go through
    OCR, in parallel when the machine has more than one usable CPU. Until the first financial statement page,
    a quick look at each scanned page's heading recognises pages that are not financial statements (auditor's
    report, CARO annexure, tax covers); those are not read in full.
    """
    from concurrent.futures import ThreadPoolExecutor

    pdf = pdfium.PdfDocument(pdf_source)
    total = len(pdf)
    results = [None] * total
    pytesseract = None
    workers = _usable_cpus()
    pending = {}
    done = 0
    statements_started = False

    def finish(idx, outcome):
        nonlocal done, statements_started
        lines, conf, rot, skipped, _ = outcome
        results[idx] = PageText(idx + 1, lines, "skipped" if skipped else "ocr", conf, rot)
        if not skipped and classify_page(lines) in ("BS", "P&L", "NOTES", "CF"):
            statements_started = True
        done += 1
        if progress:
            note = f" (skipped: {skipped})" if skipped else ""
            progress(done, total, f"Read page {idx + 1} of {total}{note}")

    with ThreadPoolExecutor(max_workers=workers) as pool:
        for idx in range(total):
            pdf_page = pdf[idx]
            lines = _embedded_lines(pdf_page)
            if sum(len(l) for l in lines) >= 80 or not ocr:
                results[idx] = PageText(idx + 1, lines, "text")
                done += 1
                if progress:
                    progress(done, total, f"Read page {idx + 1} of {total}")
                continue
            if pytesseract is None:
                pytesseract = _configure_tesseract()
            # pdfium is not thread-safe: render here, OCR in the pool
            base = pdf_page.render(scale=OCR_SCALE).to_pil().convert("L")
            small = pdf_page.render(scale=_TRIAGE_SCALE).to_pil().convert("L")
            pending[idx] = pool.submit(_read_scanned_page, pytesseract, base, small, triage and not statements_started)
            while len(pending) >= workers:          # keep memory bounded; also lets triage stop early
                first = min(pending)
                finish(first, pending.pop(first).result())
        for idx in sorted(pending):
            finish(idx, pending[idx].result())
    return results


# --------------------------------------------------------------------------
# Page typing, units, fiscal-year columns
# --------------------------------------------------------------------------

_MONTHS = "jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec"


_STATEMENT_TITLE = re.compile(r"BALANCE\s*SHE[EI]\w*\s*(?:AS|FOR)\b|STATEMENT\s*OF\s*PROFIT|PROFIT\s*(?:AND|&)\s*LOSS\s*"
                              r"(?:STATEMENT|ACCOUNT)\s*(?:FOR|AS)\b", re.I)


def classify_page(lines):
    head = re.sub(r"\s+", " ", " ".join(lines[:14]).upper())
    for line in lines[:10]:
        # a stray scan mark or serial ("4 | NOTES TO AND FORMING ...") may precede the title
        if len(line) < 130 and re.search(r"(?:^|\s)NOTES?\s+(?:TO\s+(?:AND\s+FORMING|THE\s+FINANCIAL|FINANCIAL)|FORMING)\b",
                                         _clean(line), re.I):
            return "NOTES"
        if len(line) < 130 and re.match(r"^\W*NOTES?\s+(?:TO|FORMING|ON)\b", _clean(line), re.I):
            return "NOTES"
    # a notes page that only carries the letterhead and then "NOTE 1 SHARE CAPITAL" / "NOTE 11 PROPERTY ..."
    for line in lines[:16]:
        text = _clean(line)
        if len(text) < 130 and re.match(r"^\W*NOTE\s*(?:NO\.?\s*)?\d{1,2}[A-Z]?\b\s*[:.\-]?\s*"
                                        r"(?!(?:JAN|FEB|MAR|APR|MAY|JUN|JUL|AUG|SEP|OCT|NOV|DEC)[A-Z]*[\s\-/.,]*\d)[A-Z(]",
                                        text, re.I) \
                and not _STATEMENT_TITLE.search(head):
            return "NOTES"
    if "INDEPENDENT AUDITOR" in head:
        return "OTHER"
    if re.search(r"BALANCE\s*SHE[EI]\w*", head):
        return "BS"
    if "PROFIT AND LOSS" in head or "PROFIT & LOSS" in head or "STATEMENT OF PROFIT" in head \
            or "INCOME STATEMENT" in head:
        return "P&L"
    if "CASH FLOW" in head:
        return "CF"
    return "OTHER"


_UNIT_PATTERNS = [
    # "Figures in ₹ Thousands" -- OCR often turns the rupee sign into %, ®, €, 3 ...
    (re.compile(r"in\s+(?:rs\.?|inr|₹)?\s*in\s*['’`]?\s*000|\bi[na]\s+\S{0,3}\s*thousands?|'000|figures\s+in\s+.{0,4}thousand",
                re.I), 1_000, "Rs. in '000"),
    (re.compile(r"in\s+(?:rs\.?\s*)?(?:lakhs?|lacs?)\b|in\s+rs\.?\s*in\s+lakhs?", re.I), 100_000, "Rs. in lakhs"),
    (re.compile(r"in\s+(?:rs\.?\s*)?crores?\b|in\s+rs\.?\s*in\s+crores?", re.I), 10_000_000, "Rs. in crores"),
    (re.compile(r"in\s+(?:rs\.?\s*)?(?:millions?|mn)\b", re.I), 1_000_000, "Rs. in millions"),
    (re.compile(r"(?:amount|amt\.?)\s*in\s+(?:rs\.?|rupees|inr|₹)|\(\s*rupees\s*\)", re.I), 1, "Rs. (whole rupees)"),
]


def detect_unit(pages):
    votes = {}
    for pg in pages:
        text = " ".join(pg.lines[:20])
        for pattern, mult, label in _UNIT_PATTERNS:
            if pattern.search(text):
                votes[(mult, label)] = votes.get((mult, label), 0) + 1
                break
    if not votes:
        return 1, "Not detected (assumed whole rupees)"
    (mult, label), _ = max(votes.items(), key=lambda kv: kv[1])
    return mult, label


def detect_fy_columns(lines):
    """FY labels in the order they appear on a face-statement header, e.g. ['FY25', 'FY24']."""
    text = " ".join(lines[:14])
    found = re.findall(rf"(?:{_MONTHS})[a-z]*[\s,.'’-]*((?:19|20)\d\d)", text, flags=re.I)
    cols = []
    for year in found:
        label = f"FY{year[-2:]}"
        if label not in cols:
            cols.append(label)
    if len(cols) >= 2:
        return cols
    ranged = re.findall(r"(?:19|20)\d\d\s*[-/]\s*(\d\d)\b", text)
    for yy in ranged:
        label = f"FY{yy}"
        if label not in cols:
            cols.append(label)
    return cols


# --------------------------------------------------------------------------
# Row parsing
# --------------------------------------------------------------------------

_NOISE = re.compile(r"[|_\[\]{}~=]+|[^ -~]+")
_DASHES = {"-", "–", "—", "nil", "Nil", "NIL"}
# "Nil" as OCR tends to read it in a figures column
_NIL_WORDS = {"nil", "nill", "nii", "nl", "mil", "hil", "wil", "pil", "nul", "nual", "mill", "rail", "mld", "na", "n.a.",
              "nit", "ni", "nll", "n1l", "ml"}


def _is_dash(tok):
    return tok in _DASHES or tok.lower().strip(".,|:;") in _NIL_WORDS


def _garbled_figure(tok):
    """A token that is clearly a figure OCR could not read: mostly digits with a stray symbol ("$23", "1,7O8")."""
    digits = sum(c.isdigit() for c in tok)
    return digits >= 2 and digits >= len(tok.replace(",", "")) - 2 and not re.fullmatch(r"[\d,().-]+", tok)


def _clean(text):
    text = _NOISE.sub(" ", text)
    return re.sub(r"\s+", " ", text).strip()


def _parse_number(tok):
    """-> (value | None, flags). Dashes mean nil (0)."""
    flags = []
    if _is_dash(tok):
        return 0.0, flags
    t = tok
    negative = False
    if t.startswith("(") and t.endswith(")"):
        negative, t = True, t[1:-1]
    elif t.startswith("-"):
        negative, t = True, t[1:]
    # OCR often reads the decimal point as a comma: 12,920,52 / 117,75
    if re.fullmatch(r"\d{1,3}(?:,\d{3})*,\d{2}", t) and not re.fullmatch(r"\d{1,3},\d{3}", t):
        t = t[: t.rfind(",")] + "." + t[t.rfind(",") + 1:]
        flags.append("decimal_repaired")
    elif re.fullmatch(r"\d{1,3},\d{2}", t):
        t = t.replace(",", ".")
        flags.append("decimal_repaired")
    t = t.replace(",", "")
    if not re.fullmatch(r"\d+(?:\.\d+)?", t):
        return None, flags
    value = float(t)
    return (-value if negative else value), flags


def _is_valuelike(tok, decimals2):
    """'value' | 'suspect' | 'dash' | 'int' | 'word'"""
    if _is_dash(tok):
        return "dash"
    core = tok[1:-1] if tok.startswith("(") and tok.endswith(")") else tok.lstrip("-")
    if not re.fullmatch(r"[\d,]*\d(?:\.\d+)?", core):
        return "word"
    if "." in core:
        return "value"
    if re.fullmatch(r"\d{1,2}", core):
        return "int"  # note reference or small whole number
    if decimals2:
        # in a 2-decimal document a plain 3+ digit token has probably lost its decimal
        return "value" if "," in core and re.search(r",\d{2}$", core) else "suspect"
    return "value"


def _page_uses_2_decimals(lines, document_default):
    """Statements in thousands and notes in rupees-and-paise can sit in one PDF: decide per page when it is clear."""
    two = other = 0
    for line in lines:
        for tok in _clean(line).split():
            t = tok.strip("()")
            if re.fullmatch(r"-?[\d,]*\d\.\d{2}", t):
                two += 1
            elif re.fullmatch(r"-?\d{1,3}(?:,\d{2,3})+", t):
                other += 1
    if two >= 6 and two >= 3 * other:
        return True
    if other >= 6 and other >= 3 * two:
        return False
    return document_default


def _document_uses_2_decimals(all_lines):
    two = other = 0
    for line in all_lines:
        for tok in _clean(line).split():
            if re.fullmatch(r"\(?-?[\d,]*\d\.\d+\)?", tok):
                if re.search(r"\.\d{2}\)?$", tok):
                    two += 1
                else:
                    other += 1
    return two >= 10 and two >= 4 * max(other, 1)


_ENUMERATOR = re.compile(r"^(?:\(?[a-zA-Z]{1,4}[\)\}]|\(?[ivxlIVXL]{1,4}\)|[ivxlIVXL]{1,4}\s|[a-gA-G]\s(?=[A-Z]))\s*")
_TOTAL = re.compile(r"^(?:grand\s+)?total\b(?!\s*outstanding)", re.I)   # not "total outstanding dues of ..."
_NOTE_HEADER = re.compile(
    r"^(?:(?i:note)\s*(?:(?i:no)\.?)?\s*)?((?=[0-9ilIoO]*\d)[0-9ilIoO]{1,2})(?:[.:\-]|\s)+(?:[0-9]{1,2}(?:[.:\-]|\s)+)?[\(\[\{]?\s*"
    r"([A-Za-z][A-Za-z &,/'()\.\-]{2,80}?)\s*:?\s*$"
)
_SKIP_LINE = re.compile(
    r"^(?:particulars|notes?|no\.?|sr\.?\s*no|\(refer|as per our report|for .* (?:&|and) co|"
    r"chartered accountants|place|date|din|m\.?\s*no|f\.?r\.?\s*no|cin|"
    # a signatory ("Partner", "Director (DIN ...)"), not "Partners current capital account" / "Directors remuneration"
    r"(?:partners?|directors?)(?![\s'’]*(?:s\b)?[\s'’]*(?:fixed|current|capital|remuneration|salary|loans?|accounts?|a/c|fund))|"
    r"company information|the accompanying|significant accounting)(?=[\s:.;,)]|$)",
    re.I,
)


def parse_line(line, ncols, decimals2, fy_columns):
    """
    -> (label, note_ref, values{FY: float|None}, flags, n_value_tokens)
    Values are peeled from the right; a trailing small integer left in the label
    is a note reference. Anything ambiguous is flagged, not guessed.
    """
    tokens, flags = [], []
    for tok in _clean(line).split():
        # OCR turns table borders into stray brackets/punctuation: "248.52)" / "(366.36" / "6,000."
        if tok.endswith(")") and "(" not in tok and len(tok) > 1:
            tok, _ = tok[:-1], flags.append("bracket_noise")
        elif tok.startswith("(") and ")" not in tok and len(tok) > 1:
            tok, _ = tok[1:], flags.append("bracket_noise")
        if len(tok) > 1 and tok[-1] in ".,;:" and tok[:-1].replace(",", "").replace(".", "").isdigit():
            tok = tok[:-1]
        if not decimals2 and re.fullmatch(r"\(?-?\d{1,3}(?:[.,]\d{3})+\)?", tok) and "." in tok:
            tok, _ = tok.replace(".", ","), flags.append("decimal_repaired")
        tokens.append(tok)
    peeled = []
    i = len(tokens)
    while i > 0 and len(peeled) < max(ncols, 1):
        kind = _is_valuelike(tokens[i - 1], decimals2)
        if kind in ("value", "dash"):
            peeled.append(tokens[i - 1])
            i -= 1
        elif kind == "suspect":
            peeled.append(tokens[i - 1])
            flags.append("missing_decimal?")
            i -= 1
        elif kind == "int" and not decimals2:
            peeled.append(tokens[i - 1])
            i -= 1
        elif (kind == "int" and decimals2 and i > 1
              and _is_valuelike(tokens[i - 2], decimals2) == "value"):
            # a note reference always sits LEFT of the figures, so a small whole number right of a real
            # figure is a value that lost its decimal point (0.31 read as 31)
            peeled.append(tokens[i - 1])
            flags.append("missing_decimal?")
            i -= 1
        elif peeled and _garbled_figure(tokens[i - 1]):
            # a figure OCR could not read ("$23" for 923) between other figures: keep its column, value unknown
            peeled.append("?")
            flags.append("unreadable_value")
            i -= 1
        elif not peeled and _garbled_figure(tokens[i - 1]) and i > 1 and _is_valuelike(tokens[i - 2], decimals2) in (
                "value", "suspect", "int", "dash"):
            peeled.append("?")
            flags.append("unreadable_value")
            i -= 1
        else:
            break
    peeled.reverse()
    lone_ref = False
    if len(peeled) == 1 and ncols >= 2 and re.fullmatch(r"\d{1,2}", peeled[0]) and i > 0:
        # "(b) TRADE PAYABLES 8": a lone small number on a statement line is its note reference, not a figure
        peeled, i, lone_ref = [], i + 1, True
    remainder = tokens[:i]
    extra_numeric = 0
    # a multi-column schedule (fixed assets, share counts ...): everything numeric to the right of the
    # label belongs to columns that are not financial years -- including OCR-mangled ones like 113:75 or 11/31
    def _numericish(tok):
        return (_is_valuelike(tok, decimals2) in ("value", "suspect", "dash")
                or bool(re.fullmatch(r"[\d,]*\d[:/;]\d+|\d{3,}", tok)))
    while remainder and _numericish(remainder[-1]):
        remainder.pop()
        extra_numeric += 1
    if extra_numeric:
        # the small integer left behind is a rate / note number, not part of the label
        if remainder and re.fullmatch(r"\d{1,2}", remainder[-1]) and len(remainder) > 1:
            remainder.pop()
        if len(remainder) > 1 and re.fullmatch(r"[A-Za-z]", remainder[-1]):
            remainder.pop()   # stray one-letter OCR token from the next column ("Computer & Peripherals S")
        flags.append("multi_column_schedule")

    note_ref = ""
    if remainder and re.fullmatch(r"\d{1,2}", remainder[-1]) and (peeled or decimals2 or lone_ref) and len(remainder) > 1:
        note_ref = remainder[-1]
        remainder = remainder[:-1]

    label = " ".join(remainder).strip(" :-.;,")
    label = _ENUMERATOR.sub("", label).strip()

    values = {}
    if peeled and len(peeled) == ncols and fy_columns and len(fy_columns) == ncols:
        for fy, tok in zip(fy_columns, peeled):
            val, f = _parse_number(tok)
            if val is None:
                flags.append("unreadable_value")
            flags.extend(f)
            values[fy] = val
    elif peeled:
        flags.append("ambiguous_columns")
    return label, note_ref, values, sorted(set(flags)), len(peeled)


def _normalise_note_no(raw):
    return raw.translate(str.maketrans("ilIoO", "11100"))


_FIXED_ASSET_PAGE = re.compile(r"gross\s*block|net\s*block|accumulated\s+depreciation", re.I)


def _repeated_header_lines(pages, zone=6, min_pages=3, similarity=82):
    """(page, raw line) pairs that are letterhead repeated across pages -- company name, address, CIN."""
    from rapidfuzz import fuzz

    def norm(t):
        return re.sub(r"[^a-z]", "", t.lower())

    zones = [(pg.page, [(l, norm(l)) for l in pg.lines[:zone] if len(norm(l)) >= 8]) for pg in pages]
    repeated = set()
    for page, items in zones:
        for raw, key in items:
            seen_on = sum(
                1
                for other_page, other_items in zones
                if other_page != page and any(fuzz.ratio(key, k2) >= similarity for _, k2 in other_items)
            )
            if seen_on >= min_pages - 1:
                repeated.add((page, raw))
    return repeated


_BODY_START = re.compile(r"BALANCE\s*SHE|PROFIT\s*(?:AND|&)\s*LOSS|STATEMENT\s*OF\s*PROFIT|INCOME\s*STATEMENT", re.I)
_BODY_END = re.compile(r"^\W*(?:vide|as\s+per)\s+our\s+report|for\s+and\s+on\s+behalf|^\W*significant\s+accounting\s+polic"
                       r"|^\W*the\s+accompanying\s+notes|^\W*the\s+notes\s+referred\s+to|^\W*(?:net\s*)?(?:profit|surplus)\b.{0,30}(?:apportioned|divisible|distributed|credited)\s+(?:amongst|among|to)\s+(?:the\s+)?partners|^\W*notes?\s+(?:on|to)\s+(?:the\s+)?financial\s+statements\s+\d", re.I)


# letterhead address on every page: "MUMBAI 400 083" -- a place name and an Indian PIN code, not a figure
_ADDRESS_LINE = re.compile(r"^[A-Za-z][A-Za-z!|.\s]{2,25}\s+\d{3}\s?[0-9OoGgQ]{3}\W*$")


def _statement_body(lines):
    """Line numbers of a face statement's own content: from its title to the auditor's / directors' signatures."""
    start = next((i for i, l in enumerate(lines) if _BODY_START.search(_clean(l))), None)
    if start is None:
        return range(len(lines))      # title not readable: keep everything rather than lose the statement
    end = next((i for i in range(start + 1, len(lines)) if _BODY_END.search(_clean(lines[i]))), len(lines))
    return range(start, end)


def extract_rows(pages, fy_columns, decimals2):
    ncols = len(fy_columns) if fy_columns else 2
    rows = []
    boilerplate = _repeated_header_lines(pages)
    section = ""        # nearest valueless heading line on a face statement
    parent = ""         # a valueless line that cites a note, e.g. "Trade Payables 7" -> its indented children
    note_title = ""
    current_note_no = ""
    notes_are_worded = False   # the document heads its notes "NOTE 1 ..."
    for pg in pages:
        if pg.page_type == "OTHER":
            continue
        statement = {"NOTES": "NOTE"}.get(pg.page_type, pg.page_type)
        if statement != "NOTE":
            section, parent = "", ""
        elif _FIXED_ASSET_PAGE.search(" ".join(pg.lines[:25])) and not any(
            _NOTE_HEADER.match(_clean(l).lstrip(": .-")) for l in pg.lines[:8]
        ):
            current_note_no = next(
                (r.note_ref for r in rows if r.statement == "BS" and r.note_ref
                 and re.search(r"property|fixed asset|plant", r.label, re.I)),
                "",
            )
            note_title = "Property, Plant and Equipments (fixed asset schedule)"
        body = _statement_body(pg.lines) if statement in ("BS", "P&L") else range(len(pg.lines))
        page_decimals = _page_uses_2_decimals(pg.lines, decimals2)   # pages of one PDF can differ
        for line_no, raw in enumerate(pg.lines):
            if line_no not in body:
                continue   # letterhead above the statement's title, signatures below it
            text = _clean(raw).lstrip(": .-")
            if not text or len(text) < 3:
                continue
            if statement == "NOTE" and (pg.page, raw) not in boilerplate:
                # checked before the skip list, which would otherwise drop "NOTE 2 : RESERVES & SURPLUS"
                m = _NOTE_HEADER.match(text)
                if m and not re.search(r"\d[\d,]*\.\d", text):
                    worded = bool(re.match(r"^\W*note\b", text, re.I))
                    if worded:
                        notes_are_worded = True
                    elif notes_are_worded:
                        continue   # "1. Administrative & Other Expenses" inside NOTE 24: a sub-heading, not a note
                    current_note_no = _normalise_note_no(m.group(1))
                    note_title = m.group(2).strip(" :")
                    continue
            if (pg.page, raw) in boilerplate or _SKIP_LINE.match(text) or _ADDRESS_LINE.match(text):
                continue
            label, note_ref, values, flags, n_vals = parse_line(raw, ncols, page_decimals, fy_columns)
            if not label:
                continue
            headingish = len(label) <= 70 and re.search(r"[A-Za-z]{3}", label) is not None \
                and not _TOTAL.match(label)
            if n_vals == 0:
                if statement == "NOTE" or not headingish:
                    continue  # narrative text / noise
                if not note_ref:
                    if label[:1].islower():
                        continue   # the wrapped second line of a long label ("small enterprises; and")
                    section, parent = label.strip(": "), ""
                    continue
                parent = label.strip(": ")  # valueless line citing a note: keep it, it resolves via the note
            elif note_ref and statement != "NOTE":
                parent = ""
            heading = note_title if statement == "NOTE" else (parent if parent and parent != label else section)
            row = ExtractedRow(
                page=pg.page,
                statement=statement,
                heading=heading,
                note_no=current_note_no if statement == "NOTE" else "",
                note_ref=note_ref,
                label=re.sub(r"\b(\w+)(?:\s+\1\b)+", r"\1", label, flags=re.I),
                values=values,
                raw_text=raw.strip(),
                is_total=bool(_TOTAL.match(label)),
                flags=flags,
            )
            if pg.source == "ocr":
                row.flags.append("ocr")
            if statement == "NOTE" and (re.match(r"^[a-zA-Z]\s", note_title) or "ageing" in note_title.lower()):
                row.flags.append("sub_schedule")
            if n_vals == 0:
                row.flags.append("no_values_found")
            rows.append(row)
    return rows


# --------------------------------------------------------------------------
# Built-in sanity checks
# --------------------------------------------------------------------------

def run_checks(rows, fy_columns, tolerance=0.05):
    """
    (1) Within each note, every run of items is compared with the Total row that closes it.
    (2) A face-statement figure is compared with the Total of the note it cites.
    Mismatches are reported and the affected rows get a flag; nothing is auto-corrected.
    """
    checks = []
    note_totals = {}          # note_no -> [total rows], matched to a face line by title similarity
    segment, current_key = [], None
    for r in rows:
        if r.statement != "NOTE" or not r.note_no:
            continue
        key = (r.note_no, r.heading)
        if key != current_key:
            segment, current_key = [], key
        if not r.is_total:
            segment.append(r)
            continue
        columns_are_not_years = "multi_column_schedule" in r.flags and "fixed asset" not in r.heading.lower()
        if columns_are_not_years:
            segment = []
            continue
        note_totals.setdefault(r.note_no, []).append(r)
        for fy in fy_columns:
            total = r.values.get(fy)
            item_sum = sum((i.values.get(fy) or 0.0) for i in segment)
            if total is None:
                checks.append(f"Note {r.note_no} ({r.heading}) {fy}: Total unreadable; items sum to {item_sum:,.2f}  CHECK")
            elif abs(item_sum - total) <= tolerance:
                checks.append(f"Note {r.note_no} ({r.heading}) {fy}: items tie to Total {total:,.2f}  OK")
            else:
                checks.append(f"Note {r.note_no} ({r.heading}) {fy}: items sum {item_sum:,.2f} vs Total {total:,.2f}  CHECK")
                for i in segment:
                    if "sum_mismatch" not in i.flags:
                        i.flags.append("sum_mismatch")
        segment = []
    bs_segments, seg = [], []
    for r in rows:
        if r.statement != "BS":
            continue
        if r.is_total:
            bs_segments.append((seg, r))
            seg = []
        elif r.values:
            seg.append(r)
    if len(bs_segments) >= 2 and all(len(items) >= 3 for items, _ in bs_segments[:2]):
        (liab, liab_total), (assets, asset_total) = bs_segments[0], bs_segments[1]
        for fy in fy_columns:
            liab_sum = sum((i.values.get(fy) or 0.0) for i in liab)
            asset_sum = sum((i.values.get(fy) or 0.0) for i in assets)
            verdict = "TALLIES  OK" if abs(liab_sum - asset_sum) <= tolerance else "DOES NOT TALLY  CHECK"
            checks.append(
                f"Balance Sheet {fy}: equity & liabilities {liab_sum:,.2f} vs assets {asset_sum:,.2f}  {verdict}"
            )
    for r in rows:
        if r.statement in ("BS", "P&L") and r.note_ref in note_totals:
            from rapidfuzz import fuzz

            nt = max(note_totals[r.note_ref], key=lambda t: fuzz.token_set_ratio(r.label.lower(), t.heading.lower()))
            for fy in fy_columns:
                face, note = r.values.get(fy), nt.values.get(fy)
                if face is not None and note is not None and abs(face - note) > tolerance:
                    checks.append(f"{r.label} {fy}: statement shows {face:,.2f} but Note {r.note_ref} total is {note:,.2f}  CHECK")
                    if "differs_from_note" not in r.flags:
                        r.flags.append("differs_from_note")
    return checks


# --------------------------------------------------------------------------
# Top-level entry point
# --------------------------------------------------------------------------

def extract_financials(pdf_source, progress=None, ocr=True):
    return extract_from_pages(read_pages(pdf_source, progress=progress, ocr=ocr))


def extract_from_pages(pages):
    """Shared by PDF and Excel input: page text -> rows, checks, units, financial years."""
    for pg in pages:
        if pg.page_type == "OTHER" or pg.source != "sheet":
            pg.page_type = classify_page(pg.lines)   # a workbook section already knows its kind
    # a notes page that simply continues the previous one (no "NOTE n" at its top, e.g. the surplus part of
    # the reserves note) still belongs to the notes
    for prev, pg in zip(pages, pages[1:]):
        figures = sum(1 for l in pg.lines if re.search(r"\d[\d,]{2,}|\bnil\b", l, re.I))
        if pg.page_type == "OTHER" and prev.page_type == "NOTES" and pg.source != "skipped" and figures >= 3:
            pg.page_type = "NOTES"

    warnings = []
    fy_columns = []
    for pg in pages:
        if pg.page_type in ("BS", "P&L"):
            cols = detect_fy_columns(pg.lines)
            if len(cols) >= 2:
                fy_columns = cols
                break
    if not fy_columns:
        warnings.append("Could not read the financial-year column headers; please confirm them manually.")

    multiplier, unit_label = detect_unit(pages)
    if unit_label.startswith("Not detected"):
        warnings.append("Reporting unit (Rs. / '000 / lakhs / crores) not detected; assumed whole rupees. Please confirm.")

    statement_lines = [ln for pg in pages if pg.page_type != "OTHER" for ln in pg.lines]
    decimals2 = _document_uses_2_decimals(statement_lines)

    rows = extract_rows(pages, fy_columns, decimals2)
    checks = run_checks(rows, fy_columns)

    ocr_pages = [p for p in pages if p.source == "ocr"]
    if ocr_pages:
        warnings.append(
            f"{len(ocr_pages)} of {len(pages)} pages were scanned images read by OCR. "
            "Figures may contain misreads or dropped rows -- review them against the PDF."
        )
    if not any(r.statement in ("BS", "P&L") for r in rows):
        warnings.append("No Balance Sheet or Profit & Loss page was recognised.")

    if any(r.statement == "BS" for r in rows) and not any(r.statement == "P&L" for r in rows):
        warnings.append("This file has no Statement of Profit and Loss. Upload the next year's report as well: its "
                        "comparative column supplies this year's profit and loss.")
    # rounding follows how the statements themselves are printed (notes may carry paise while they do not)
    face_lines = [ln for pg in pages if pg.page_type in ("BS", "P&L") for ln in pg.lines]
    whole_units = not _document_uses_2_decimals(face_lines) if face_lines else not decimals2
    return ExtractionResult(rows, fy_columns, multiplier, unit_label, pages, checks, warnings, whole_units=whole_units)
