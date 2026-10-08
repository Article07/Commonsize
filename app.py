"""
Commonsize web app (Streamlit).

    streamlit run app.py        # set APP_PASSWORD first

Upload a company's financials (digital PDF, scanned PDF, or a structured Excel
sheet), review what was read, and download a Common Size workbook built on the
firm's Reference file. Everything is rule-based; no AI service is called.
"""

import hmac
import os
import tempfile
from pathlib import Path

import pandas as pd
import streamlit as st

from config.schedule_map import FY_COLUMN_MAP
import ui
from core import pipeline
from core.routing import all_placeable_schedules, schedule_for
from core.multi_year import Source, combine, fy_key
from core.workbook_reader import merge_new_years, read_workbook

st.set_page_config(page_title="Acumen | Common Size Generator", page_icon=None, layout="wide")
ui.apply_theme()

NOT_PLACED = "(not placed)"


# ----------------------------------------------------------------------------- password gate
def _gate():
    expected = os.environ.get("APP_PASSWORD", "")
    if not expected:
        st.error("APP_PASSWORD is not set on the server, so the app is locked. Ask the administrator to set it.")
        st.stop()
    if st.session_state.get("authed"):
        return
    ui.login_header()
    _, middle, _ = st.columns([1, 1.2, 1])
    with middle:
        # A plain text box, masked with dots by CSS (ui.py), so browser password managers neither fill in a
        # saved password nor offer to create one; they only act on type="password" fields.
        entered = st.text_input("Password", autocomplete="off", key="access_code")
    if entered and hmac.compare_digest(entered.encode(), expected.encode()):
        st.session_state["authed"] = True
        st.rerun()
    elif entered:
        with middle:
            st.error("Incorrect password.")
    ui.footer()
    st.stop()


# ----------------------------------------------------------------------------- helpers
def _fmt(value):
    return f"{value:,.2f}"


def _review_table(items, fy_cols, mult, unit_label, key):
    """Editable table of line items; returns the items after the user's edits (excluded rows removed)."""
    options = [NOT_PLACED] + all_placeable_schedules()
    rows = []
    for item in items:
        row = {
            "Include": True,
            "Line item": item.line_item,
            "Place under": schedule_for(item) or NOT_PLACED,
            "New (yellow)": item.is_new,
            "PDF heading": item.pdf_heading,
            "Page": item.pdf_page or None,
            "Flags": ", ".join(f for f in item.flags if f != "ocr"),
        }
        for fy in fy_cols:
            v = item.fy_values.get(fy)
            row[fy] = None if v is None else v / mult
        rows.append(row)
    df = pd.DataFrame(rows)
    st.caption(f"Amounts are shown in {unit_label}. Correct any misread figure, choose a different schedule in "
               "'Place under', or untick Include. The balance check below updates as you edit.")
    edited = st.data_editor(
        df, key=key, hide_index=True, width="stretch", num_rows="fixed",
        column_config={
            "Place under": st.column_config.SelectboxColumn(options=options, required=True),
            "Include": st.column_config.CheckboxColumn(),
            **{fy: st.column_config.NumberColumn(format="%.2f") for fy in fy_cols},
        },
        disabled=["New (yellow)", "PDF heading", "Page", "Flags"],
        column_order=["Include", "Line item", "Place under", *fy_cols, "New (yellow)", "Flags", "PDF heading", "Page"],
    )
    final = []
    for item, (_, row) in zip(items, edited.iterrows()):
        if not row["Include"]:
            continue
        chosen = row["Place under"]
        if chosen != (schedule_for(item) or NOT_PLACED):
            item.schedule_override = "" if chosen == NOT_PLACED else chosen
        item.line_item = row["Line item"]
        item.fy_values = {fy: (None if pd.isna(row[fy]) else float(row[fy]) * mult) for fy in fy_cols}
        final.append(item)
    return final


def _balance_panel(items, fy_cols, mult, unit_label):
    result = pipeline.check(items, fy_cols, mult)
    worst = max((abs(y.gap) for y in result.years), default=0.0)
    ui.cards([
        ("Years checked", ", ".join(y.fy for y in result.years), ""),
        ("Balance Sheet", "Tallies" if result.balanced else "Does not tally", "ok" if result.balanced else "bad"),
        (f"Largest gap ({unit_label})", _fmt(worst / mult), "ok" if result.balanced else "bad"),
    ])
    if result.balanced:
        st.success("The Balance Sheet tallies in every year.")
    else:
        st.error("The Balance Sheet does not tally. The gap below is real, not rounded away; "
                 "look for a misread figure, a missing line, or an item placed under the wrong schedule.")
    st.dataframe(pd.DataFrame([{
        "Year": y.fy,
        f"Equity + borrowings ({unit_label})": _fmt(y.liabilities_side / mult),
        f"Assets ({unit_label})": _fmt(y.assets_side / mult),
        f"Gap ({unit_label})": _fmt(y.gap / mult),
        f"Profit for the year ({unit_label})": _fmt(y.pat / mult),
        "Status": "OK" if y.balanced else "Does not tally",
    } for y in result.years]), hide_index=True, width="stretch")
    if result.unplaced:
        st.warning("Not placed under any schedule (excluded from the statements): "
                   + ", ".join(i.line_item for i in result.unplaced))
    if len(result.years) > 1 and not result.balanced:
        st.caption("A profit misread in one year also shows up in the following years, because profit rolls into reserves.")
    return result


def _stop():
    ui.footer()
    st.stop()


def _generate(items, fy_cols, company, equity, reported, file_name):
    ui.step(4, "Download", "The workbook uses the firm's Reference format; new line items are highlighted yellow.")
    data, report = pipeline.build_workbook(items, fy_cols, company, equity or None, reported)
    for line in report.log:
        st.warning(line) if line.startswith("WARNING") else st.info(line)
    new = report.new_rows
    if new:
        st.info(f"{len(new)} new line item(s) are highlighted yellow in the workbook (Sch. sheet): "
                + ", ".join(p.label for p in new))
    st.download_button("Download Common Size workbook", data, file_name=file_name,
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", type="primary")


UNIT_CHOICES = {"Rs. (whole rupees)": 1, "Rs. in '000": 1_000, "Rs. in lakhs": 100_000,
                "Rs. in millions": 1_000_000, "Rs. in crores": 10_000_000}


def _is_structured(raw):
    """The older Statement / Line Item / Statement Tag / FY.. layout."""
    import io

    import openpyxl
    try:
        ws = openpyxl.load_workbook(io.BytesIO(raw), read_only=True).worksheets[0]
        first = next(ws.iter_rows(max_row=1, values_only=True), ())
        return any(str(c).strip().lower() == "line item" for c in first if c is not None)
    except Exception:
        return False


def _unit_label(mult):
    return next((label for label, m in UNIT_CHOICES.items() if m == mult), "Rs. (whole rupees)")


def _read_source(upload, kind, bar, index, total):
    """One uploaded file -> Source."""
    raw = upload.getvalue()
    label = f"File {index + 1} of {total}: {upload.name}"
    bar.progress(index / total, text=label)
    if kind == "pdf":
        extraction, mapping = pipeline.read_pdf(
            raw, progress=lambda d, t, m: bar.progress(min((index + d / t) / total, 1.0), text=f"{label}  ({m})"))
    elif upload.name.lower().endswith(".csv") or _is_structured(raw):
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / upload.name
            path.write_bytes(raw)
            items = pipeline.read_structured(path)
        fys = sorted({fy for i in items for fy in i.fy_values}, key=fy_key)
        return Source(upload.name, items, fys)
    else:
        extraction, mapping = pipeline.read_excel(raw)
    return Source(upload.name, mapping.items, sorted(extraction.fy_columns, key=fy_key), extraction.unit_multiplier,
                  extraction.unit_label, extraction, mapping.log + extraction.warnings)


def _combine_into(data):
    sources = data["sources"]
    comb = combine(sources)
    mults = {s.mult for s in sources}
    mult = mults.pop() if len(mults) == 1 else 1
    data.update(items=comb.items, fy_cols=comb.fy_cols, owner=comb.owner, warnings=comb.warnings, mult=mult,
                unit=sources[0].unit if len(sources) == 1 else _unit_label(mult))


def _per_file(data, fy_cols):
    """[(index, source, years of fy_cols taken from it)] for files that have an extraction."""
    out = []
    for i, s in enumerate(data["sources"]):
        years = [fy for fy in fy_cols if data["owner"].get(fy) == i]
        if s.extraction is not None and years:
            out.append((i, s, years))
    return out


def _reported(data, fy_cols):
    reported = {}
    for _, s, years in _per_file(data, fy_cols):
        pat = pipeline.reported_pat(s.extraction)
        reported.update({fy: pat[fy] for fy in years if fy in pat})
    return reported


def _reconciliation_panel(data, items, fy_cols, mult, unit_label):
    """Shows that the workbook's figures equal what each input file itself prints, line by line."""
    many = len(data["sources"]) > 1
    rows, own_gaps = [], []
    for _, s, years in _per_file(data, fy_cols):
        for r in pipeline.reconcile(s.extraction, items, years):
            r["file"] = s.name
            rows.append(r)
        src = pipeline.source_balance(s.extraction, years)
        tol = pipeline.tolerance(s.extraction.unit_multiplier)
        own_gaps += [(s.name, fy, a, b) for fy, (a, b) in src.items() if abs(a - b) > tol]
    if not rows:
        st.info("The input does not print totals that can be compared (for example an unreadable scan).")
        return
    bad = [r for r in rows if not r["ok"]]
    ui.cards([
        ("Figures compared with the input", len(rows), ""),
        ("Matching", len(rows) - len(bad), "ok"),
        ("Differing", len(bad), "bad" if bad else "ok"),
    ])
    if bad:
        st.error(f"{len(bad)} figure(s) differ from the input. See the table; correct the figure in the review table above.")
    else:
        st.success("Every compared figure (revenue, total income, total expenses, PBT, tax, profit, balance-sheet totals) "
                   "equals the input.")
    st.dataframe(pd.DataFrame([{
        **({"File": r["file"]} if many else {}),
        "Year": r["year"], "Line": r["line"], f"Workbook ({unit_label})": _fmt(r["computed"] / mult),
        f"Input ({unit_label})": _fmt(r["reported"] / mult), "Difference": _fmt(r["diff"] / mult),
        "Match": "OK" if r["ok"] else "DIFFERS"} for r in rows]), hide_index=True, width="stretch")
    if own_gaps:
        st.warning("The input's own Balance Sheet does not balance: " + "; ".join(
            (f"{name}, " if many else "") + f"{fy}: total liabilities {a / mult:,.2f} vs total assets {b / mult:,.2f} "
            f"(out by {(a - b) / mult:,.2f})" for name, fy, a, b in own_gaps)
            + ". The workbook reproduces the input faithfully, so this difference shows in the Balance Sheet check. "
              "It has to be corrected in the source financials; the app does not plug it.")


def _suggestions(data, items, fy_cols):
    found = [f for _, s, years in _per_file(data, fy_cols) for f in pipeline.suggest_fixes(s.extraction, items, years)]
    if not found:
        return
    st.markdown("**Suggested corrections**")
    st.write("Each figure below disagrees with its own note. Using the note's figure makes the year's profit equal "
             "the profit the financials print, so it is very likely a misread.")
    st.dataframe(pd.DataFrame([{"Line item": i.line_item, "Year": fy, "As read": _fmt(cur), "Per note": _fmt(alt)}
                               for i, fy, cur, alt in found]), hide_index=True, width="stretch")
    if st.button("Apply the suggested corrections"):
        for item, fy, _, alt in found:
            item.fy_values[fy] = alt
            item.flags.append("corrected_from_note")
        st.session_state["review_version"] = st.session_state.get("review_version", 0) + 1
        st.rerun()


# ----------------------------------------------------------------------------- app
_gate()
ui.header()
ui.hero("Common Size, prepared from the financial statements",
        "Upload a company's financials as PDFs (digital or scanned) or Excel files, one or several years at once. "
        "The app reads them, classifies every line, checks that the balance sheet tallies and the income statement "
        "matches the input, and gives you the firm's Common Size workbook.")
ui.step(1, "Choose the input", "Start a new workbook, or add years to one you already prepared.")
mode = st.radio("What would you like to do?", ["New common size", "Add a new financial year to an existing workbook"],
                horizontal=True)

existing = None
if mode.startswith("Add"):
    old = st.file_uploader("Existing common size workbook (generated by this app)", type=["xlsx"], key="old")
    if old is None:
        _stop()
    try:
        with tempfile.TemporaryDirectory() as tmp:
            p = Path(tmp) / "existing.xlsx"
            p.write_bytes(old.getvalue())
            existing = read_workbook(p)
    except ValueError as exc:
        st.error(str(exc))
        _stop()
    st.success(f"Read '{existing.company}' with years {', '.join(existing.fy_columns) or 'none'}.")

kind_label = st.radio("Input file type", ["PDF (digital or scanned)", "Excel"], horizontal=True)
kind = "pdf" if kind_label.startswith("PDF") else "excel"
uploads = st.file_uploader(
    "Financial statements (Balance Sheet, Profit & Loss and Notes). Several files of the same company can be "
    "uploaded together, e.g. one per financial year.",
    type=["pdf"] if kind == "pdf" else ["xlsx", "xlsm", "csv"], key=f"fin-{kind}", accept_multiple_files=True)
company = st.text_input("Company name", value=existing.company if existing else "")
equity = st.text_input("Number of equity shares (optional)", value=existing.equity_shares if existing else "")

if not uploads:
    _stop()

sig = (kind, tuple((u.name, u.size) for u in uploads))
if st.session_state.get("sig") != sig:
    st.session_state.pop("data", None)
    st.session_state["sig"] = sig
    st.session_state["review_version"] = 0
if "data" not in st.session_state:
    label = "Read the file" if len(uploads) == 1 else f"Read the {len(uploads)} files"
    if st.button(label, type="primary"):
        bar = st.progress(0.0, text="Starting ...")
        sources = []
        for i, upload in enumerate(uploads):
            try:
                sources.append(_read_source(upload, kind, bar, i, len(uploads)))
            except Exception as exc:   # OCR missing, unreadable file ...
                bar.empty()
                st.error(f"Could not read '{upload.name}': {exc}")
                _stop()
        bar.empty()
        data = {"sources": sources}
        _combine_into(data)
        st.session_state["data"] = data
        st.rerun()
    _stop()

data = st.session_state["data"]
sources = data["sources"]
ui.step(2, "Review what was read", "Correct any misread figure, move a line to another schedule, or untick it.")

readable = [(i, s) for i, s in enumerate(sources) if s.extraction is not None]
if readable:
    labels = list(UNIT_CHOICES)
    holder = st.expander("Units of the figures") if len(readable) > 1 else st.container()
    with holder:
        for i, s in readable:
            current = _unit_label(s.mult)
            caption = f"Unit of the figures in {s.name} (detected: {s.unit})" if len(sources) > 1 \
                else f"Unit of the figures (detected: {s.unit})"
            chosen = st.selectbox(caption, labels, index=labels.index(current), key=f"unit-{i}")
            if UNIT_CHOICES[chosen] != s.mult:
                s.items = pipeline.remap(s.extraction, UNIT_CHOICES[chosen]).items
                s.mult, s.unit = UNIT_CHOICES[chosen], chosen
                _combine_into(data)
                st.session_state["review_version"] += 1
                st.rerun()
mult, unit = data["mult"], data["unit"]

if len(sources) > 1:
    by_file = {}
    for fy in data["fy_cols"]:
        by_file.setdefault(sources[data["owner"][fy]].name, []).append(fy)
    st.caption("Each year is taken from the file in which it is the current year: "
               + "; ".join(f"{', '.join(fys)} from {name}" for name, fys in by_file.items()) + ".")
for w in data.get("warnings", []):
    st.warning(w)
with st.expander("What the reader found and checked"):
    for s in sources:
        if len(sources) > 1:
            st.markdown(f"**{s.name}**")
        for line in s.log + (s.extraction.checks if s.extraction is not None else []):
            st.write("- " + line)

unsupported = [fy for fy in data["fy_cols"] if fy not in FY_COLUMN_MAP]
if unsupported:
    st.warning(f"{', '.join(unsupported)} is outside the template's years ({', '.join(FY_COLUMN_MAP)}) and will not be written.")

available = [fy for fy in data["fy_cols"] if fy in FY_COLUMN_MAP]
if existing:
    later = [fy for fy in available if not existing.fy_columns
             or list(FY_COLUMN_MAP).index(fy) > list(FY_COLUMN_MAP).index(existing.fy_columns[-1])]
    if not later:
        st.error("The files have no year later than the latest year already in the workbook.")
        _stop()
    review_cols = st.multiselect("Years to add", later, default=later)
else:
    if len(sources) == 1:
        # a third, older column of a single file is often only a partial comparative
        default = available[-2:] if len(available) > 2 else available
    else:
        # with several files, the years that are some file's own current year; an older comparative-only
        # year can still be ticked
        currents = {s.current for s in sources}
        default = [fy for fy in available if fy in currents] or available
    review_cols = st.multiselect("Years to include", available, default=default)
if not review_cols:
    _stop()
review_cols = sorted(review_cols, key=fy_key)

items = _review_table(data["items"], review_cols, mult, unit, key=f"review-{sig}-{st.session_state['review_version']}")
reported_all = _reported(data, review_cols)

ui.step(3, "Check the numbers", "The balance sheet must tally and the income statement must match the input.")
if existing:
    try:
        merged, all_cols, reported, warns = merge_new_years(existing, items, review_cols, reported_all)
    except ValueError as exc:
        st.error(str(exc))
        _stop()
    for w in warns:
        st.warning(w)
    _balance_panel(merged, all_cols, mult, unit)
    _generate(merged, all_cols, company or existing.company, equity, reported,
              f"{(company or existing.company)} - Common Size.xlsx")
else:
    _suggestions(data, items, review_cols)
    _balance_panel(items, review_cols, mult, unit)
    _reconciliation_panel(data, items, review_cols, mult, unit)
    _generate(items, review_cols, company or "Company", equity, reported_all, f"{company or 'Company'} - Common Size.xlsx")
ui.footer()
