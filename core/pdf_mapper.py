"""
Turns the rows extracted from a PDF (core/pdf_reader.py) into LineItem
records classified exactly like items that arrive through the Excel input
path, so analysis, balance check and the workbook writer are shared.

Which rows become line items
  * A face-statement line that cites a note is replaced by that note's items
    -- the granular detail the template's schedules want -- but ONLY when the
    items add up to the face figure. If they do not (OCR dropped a row, or the
    note is a different kind of table), the face line is used on its own and
    flagged "breakup_unreliable". The statements therefore always tie to the
    face, whatever the quality of the breakup.
  * A face line with no figures (e.g. "Revenue from operations 15") takes its
    figures from the note's Total.
  * A valueless "parent" line whose children carry the figures (Trade
    Payables -> MSME / Others) is skipped so nothing is counted twice.
  * Reserves are handled from their note: the opening balance (first year
    only; the template rolls later years forward from profit) and any other
    movement. The face "Reserves and Surplus" figure is a closing balance and
    is deliberately not used.
  * Notes that no face line cites (EPS, related parties, contingent
    liabilities ...) are not part of the statements and are ignored.

Every item records the heading the PDF presents it under, so an item that is
not in the taxonomy can be placed "exactly as the PDF classifies it".
"""

import re
from dataclasses import dataclass, field, replace

from rapidfuzz import fuzz, process

from core.models import LineItem
from core.resolver import resolve
from core.routing import schedule_for

TOL = 0.05
_SUMMARY_LINE = re.compile(
    r"earning|per\s*equity\s*share|exceptional|extraordinary|profit|loss\b|net\s*worth|^(?:basic|diluted)|^total\b(?!\s*outstanding)"
    r"|ebitda|\bratio\b|%\s*$"
    # the same lines with OCR typos ("PROAIT BEFORE TAX", "(2) DUTED", "Wl, TOTAL REVENUE")
    r"|before\s+tax\b|\bbasic\b|\b(?:diluted|duted|diuted|dlluted)\b|\btotal\s+(?:revenue|income|expenses?)\b"
    r"|dinary\s+items",          # (extra)ordinary items, also as OCR misreads it ("CATRACRDINARY ITEMS")
    re.I,
)
# standard Schedule III line names: a face line whose name OCR garbled ("REVEMUE FROM OPERATIONS") is matched to these
_CANONICAL_LINES = [
    "revenue from operations", "other income", "cost of materials consumed", "purchases of stock in trade",
    "changes in inventories of finished goods work in progress and stock in trade", "employee benefits expense",
    "finance costs", "depreciation and amortization expense", "other expenses", "current tax", "deferred tax",
    "share capital", "long term borrowings", "short term borrowings", "trade payables", "other current liabilities",
    "short term provisions", "long term provisions", "other long term liabilities", "property plant and equipment",
    "intangible assets", "capital work in progress", "non current investments", "current investments",
    "long term loans and advances", "short term loans and advances", "inventories", "trade receivables",
    "cash and cash equivalents", "other current assets", "other non current assets",
]
_SCALES = (1, 1 / 1000, 1 / 100000, 1 / 10000000, 1000)
_RESERVES = re.compile(r"reserves?\s*(?:and|&)\s*surplus", re.I)
# a partnership's / proprietor's capital account that takes the year's profit (not a fixed capital account): it
# rolls forward like Reserves & Surplus -- opening balance, partners' remuneration / interest / drawings, profit
_OWNERS_CAPITAL = re.compile(r"(?:partners?|proprietors?|owners?)[\s'’s]*(?!fixed)(?:current\s*)?capital(?:\s*accounts?|\s*a/?c)?\b"
                             r"|current\s*capital\s*accounts?", re.I)
_OWNERS_NOTE = re.compile(r"owners?[\s'’]*s?\s*funds?|partners?[\s'’]*s?\s*capital|proprietor[\s'’]*s?\s*capital|capital\s*accounts?",
                          re.I)
_OPENING = re.compile(r"opening|as\s*per\s*(?:last|previous)\s*balance\s*sheet|balance\s*b\W*f|brought\s*forward", re.I)
# the year's profit line in a reserves note: the template adds the P&L profit itself
_PROFIT_LINE = re.compile(r"net\s*profit|profit\s*for|profit\s*/\s*\(?loss\)?|closing", re.I)
# schedules that share one line of the income statement, so a note item may move between them
_SAME_IS_LINE = [{"Administrative Expenses", "Selling & Distribution Expenses"}, {"COGS-Purchase", "COGS-Manufacturing"}]


@dataclass
class NoteGroup:
    note_no: str
    title: str
    items: list
    total: object = None   # the closing "Total" row, if any


@dataclass
class MappingResult:
    items: list = field(default_factory=list)
    log: list = field(default_factory=list)
    skipped: list = field(default_factory=list)   # (label, reason)


def _has_values(row):
    return any(v is not None for v in row.values.values())


def _build_groups(note_rows):
    """note_no -> [NoteGroup]; a number can repeat (two different 'Note 15's), so titles tell them apart."""
    groups, current, key = {}, None, None
    for r in note_rows:
        if not r.note_no:
            continue
        k = (r.note_no, r.heading)
        if k != key:
            current = NoteGroup(r.note_no, r.heading, [])
            groups.setdefault(r.note_no, []).append(current)
            key = k
        if r.is_total:
            if current.total is None or _has_values(r):
                current.total = r
        elif _has_values(r):
            current.items.append(r)
    for options in groups.values():
        for g in options:
            _absorb_implicit_total(g)
    return groups


def _absorb_implicit_total(group, fy_columns=None):
    """If the last 'item' equals the sum of the items above it, it is the Total row with a garbled label."""
    if group.total is not None or len(group.items) < 3:
        return
    last, above = group.items[-1], group.items[:-1]
    fys = [fy for fy, v in last.values.items() if v is not None]
    if fys and all(abs(_sum_for(above, fy) - last.values[fy]) <= TOL for fy in fys):
        group.total = last
        group.items = above


def _plain(text):
    """Lower-case words only, for comparing a statement line with a note title ("LONG-TERM ..." = "Long Term ...")."""
    words = re.sub(r"[^a-z]+", " ", (text or "").lower()).split()
    return " ".join(w for w in words if len(w) > 2 or w in ("of", "in", "to"))


def _pick_group(groups, face_row):
    options = groups.get(face_row.note_ref, [])
    if options:
        best = max(options, key=lambda g: fuzz.token_set_ratio(face_row.label.lower(), g.title.lower()))
        if fuzz.token_set_ratio(face_row.label.lower(), best.title.lower()) >= 45:
            return best
    # note number missing or unreadable on a scanned statement ("a2" for 12): find the note by its title
    label = re.sub(r"^[\W\da-z]{0,4}\s", "", face_row.label.lower()).strip()
    label = re.sub(r"\s+\S{1,2}$", "", label)          # trailing OCR remnant of the note number
    if len(label) < 6:
        return None
    scored = sorted(((fuzz.ratio(label, g.title.lower().strip(" :.")), g) for opts in groups.values() for g in opts),
                    key=lambda t: -t[0])
    if scored and scored[0][0] >= 88 and (len(scored) == 1 or scored[1][0] < scored[0][0] - 5):
        return scored[0][1]
    return None


def _sum_for(items, fy):
    return sum((i.values.get(fy) or 0.0) for i in items)


def _reference_values(face_row, group, fy_columns):
    """Figures the note's items must add up to: the face figures, else the note's own Total."""
    if all(face_row.values.get(fy) is not None for fy in fy_columns):
        return {fy: face_row.values[fy] for fy in fy_columns}
    if group is not None and group.total is not None and any(v is not None for v in group.total.values.values()):
        return {fy: group.total.values.get(fy) for fy in fy_columns}
    return None


def _items_tie(items, reference, fy_columns, tol=TOL):
    compared = 0
    for fy in fy_columns:
        ref = reference.get(fy)
        if ref is None:
            continue
        compared += 1
        if abs(_sum_for(items, fy) - ref) > tol:
            return False
    return compared > 0


def _scale_between(note_values, face_values, fy_columns):
    """
    Factor that turns a note's figures into the statement's unit. Statements in Rs. '000 with notes in
    rupees and paise happen within one set of financials; the note's total tells which factor applies.
    """
    """Returns the factor, 1 when there is nothing to compare, None when no factor fits."""
    pairs = [(note_values.get(fy), face_values.get(fy)) for fy in fy_columns]
    pairs = [(n, f) for n, f in pairs if n not in (None, 0) and f not in (None, 0)]
    if not pairs:
        return 1
    for s in _SCALES:
        tight = sum(1 for n, f in pairs if abs(n * s - f) <= max(TOL, 0.6 if s < 1 else TOL, 0.002 * abs(f)))
        # one year agreeing exactly is enough, as long as no year is wildly off (an OCR-misread digit is not)
        if tight >= 1 and all(abs(n * s - f) <= 0.1 * abs(f) for n, f in pairs):
            return s
    return None


def _rescaled(group, scale):
    if scale == 1 or group is None:
        return group
    def sc(r):
        return replace(r, values={k: (v * scale if v is not None else None) for k, v in r.values.items()},
                       flags=list(r.flags) + ["note_in_other_unit"])
    return NoteGroup(group.note_no, group.title, [sc(r) for r in group.items], sc(group.total) if group.total else None)


def _note_reference_values(group, fy_columns):
    if group.total is not None and _has_values(group.total):
        return group.total.values
    return {fy: _sum_for(group.items, fy) for fy in fy_columns}


def _with_deductions(items, reference, fy_columns, tol=TOL):
    """
    Note lines such as "Less: Branch Transfer 6,326,463" are printed as positive figures but are deducted.
    If the note does not add up as printed but does once its "Less" lines are subtracted, use that.
    Returns the items to use (deducted lines negated), or None if neither version ties.
    """
    if _items_tie(items, reference, fy_columns, tol):
        return items
    adjusted = []
    for r in items:
        if re.match(r"^\W*less\b", r.label, re.I):
            r = replace(r, values={fy: (-abs(v) if v is not None else None) for fy, v in r.values.items()},
                        flags=list(r.flags) + ["deduction"])
        adjusted.append(r)
    return adjusted if _items_tie(adjusted, reference, fy_columns, tol) else None


def _scaled(values, fy_columns, multiplier):
    return {fy: (values[fy] * multiplier if values.get(fy) is not None else None) for fy in fy_columns}


def _all_zero(values):
    return all((v is None or abs(v) < 1e-9) for v in values.values())


def _tidy(label):
    """Drop unbalanced brackets left over from table borders: "Services (Inter Branch )" -> "Services Inter Branch"."""
    if label.count("(") != label.count(")"):
        label = label.replace("(", " ").replace(")", " ")
    label = re.sub(r"\s+\)", ")", label)
    label = re.sub(r"^\W*less\s*[:\-]+\s*", "Less: ", label, flags=re.I)
    return re.sub(r"\s{2,}", " ", label).strip()


def _make_item(label, values, statement, heading, source, page, flags, fy_columns, multiplier, resolve_label=None):
    label = _tidy(label)
    res = resolve(resolve_label or label, heading, statement)
    if source == "pdf-face" and (not res.resolved or res.via == "heading"):
        # a statement line whose name OCR garbled ("(7) OLHER CURRENT ASSEIS ly"): match it to the standard
        # Schedule III names; its own name, read this way, beats a guess from the section heading
        plain = re.sub(r"[^a-z ]", " ", label.lower())
        plain = re.sub(r"^\s*(?:[ivxl]{1,4}|[a-h])\s+", "", re.sub(r"\s+", " ", plain)).strip()
        plain = re.sub(r"(?:\s+[a-z]{1,2})+$", "", plain)        # trailing remnant of a note number
        best = process.extractOne(plain, _CANONICAL_LINES, scorer=fuzz.ratio)
        if best and best[1] >= 85:
            res = resolve(best[0], heading, statement)
    if (source == "pdf-face" and statement == "P&L" and label.startswith("Less: ")
            and res.tag not in ("Revenue",) and res.category not in ("Other Income",)):
        # "Less: Remuneration to partners" / "Less: Current tax" under a profit subtotal: an expense line like any
        # other, not a deduction within its schedule
        label = label[len("Less: "):]
    li = LineItem(
        statement=statement,
        line_item=label,
        statement_tag=res.tag,
        fy_values=_scaled(values, fy_columns, multiplier),
        category=res.category,
        source=source,
        pdf_heading=heading,
        pdf_page=page,
        is_new=res.resolved and not res.known,
        matched_item=res.row_label if (res.via == "alias" and res.row_label) else res.matched_item,
        match_score=res.score,
        flags=list(flags),
    )
    if res.known:
        li.confidence = "Matched"
    elif res.resolved:
        li.confidence = f"Uncertain (new item under PDF heading '{heading}')"
    else:
        li.confidence = "Uncertain (could not be placed)"
        li.flags.append("unplaced")
    return li


def _less_sign(group, fy_columns):
    """
    Sign for 'Less: ...' movement rows, per FY. Normally a deduction (-1), but OCR can drop the brackets
    or the note may show the line already signed, so the note's own roll-forward decides: whichever sign
    makes Opening + movements = Closing wins; if neither does, the label's meaning (deduction) stands.
    """
    def row_value(pattern, fy):
        for r in group.items:
            if re.search(pattern, r.label, re.I) and r.values.get(fy) is not None:
                return r.values[fy]
        return None

    signs = {}
    for fy in fy_columns:
        opening, closing = row_value(r"opening", fy), row_value(r"closing", fy)
        if closing is None and group.total is not None:
            closing = group.total.values.get(fy)
        signs[fy] = -1
        if opening is None or closing is None:
            continue
        fixed = 0.0
        less_abs = 0.0
        for r in group.items:
            v = r.values.get(fy)
            low = r.label.lower()
            if v is None or re.search(r"opening", low) or _PROFIT_LINE.search(low):
                continue
            if low.startswith("less"):
                less_abs += abs(v)
            else:
                fixed += v
        profit = row_value(r"net\s*profit|profit\s*for|profit\s*/\s*\(?loss\)?", fy) or 0.0
        base = opening + profit + fixed
        if abs(base - less_abs - closing) <= TOL:
            signs[fy] = -1
        elif abs(base + less_abs - closing) <= TOL:
            signs[fy] = 1
    return lambda fy: signs.get(fy, -1)


def _side_rules(item, statement, label, assets_side):
    """
    Deferred tax: the side of the balance sheet it is printed on decides, whatever the label says
    ("Deferred Tax Liabilities (net)" printed under Non-Current Assets is an asset). assets_side is
    True / False from the line's position, or None when the position is unknown (then the label decides).
    """
    if statement == "BS" and re.search(r"def+er+ed\s*tax", label, re.I):
        is_asset = assets_side if assets_side is not None else bool(re.search(r"asset", label, re.I))
        item.statement_tag = "Deferred Tax Assets" if is_asset else "Deferred Tax Liabilities"
        item.category, item.is_new, item.confidence = "", False, "Matched"
        item.flags = [f for f in item.flags if f != "unplaced"]
    return item


def _is_digital(group):
    rows = group.items + ([group.total] if group.total is not None else [])
    return bool(rows) and not any("ocr" in r.flags for r in rows)


def _map_owners_capital(g, fy_columns, multiplier, result, note_scale=1):
    """
    A partners' / proprietor's capital note: for each partner, 'As per last Balance Sheet', 'Add: remuneration,
    interest, credits', 'Share of profit', 'Less: drawings, TDS'. Openings -> Opening Reserves, movements ->
    Reserves Adjustments, the profit share is left to the template (it adds the P&L profit). Balances printed
    before the first opening line (the fixed capital accounts) and the per-partner closing balances are skipped.
    True if the note had a roll-forward.
    """
    g = _rescaled(g, note_scale)
    start = next((i for i, r in enumerate(g.items) if _OPENING.search(r.label)), None)
    if start is None:
        return False
    sign = 1
    for r in g.items[start:]:
        low = r.label.lower().strip()
        if _OPENING.search(low):
            tag, sign = "Opening Reserves", 1
        elif re.match(r"^\W*less\b", low):
            tag, sign = "Reserves Adjustments", -1
        elif re.match(r"^\W*add\b", low):
            tag, sign = "Reserves Adjustments", 1
        elif _PROFIT_LINE.search(low) or re.search(r"share\s*of\s*(?:net\s*)?(?:profit|loss)", low):
            continue
        elif r.is_total or not _has_values(r):
            continue
        else:
            tag = "Reserves Adjustments"          # a continuation line under 'Add:' / 'Less:'
        values = {fy: (sign * abs(v) if (tag == "Reserves Adjustments" and v is not None) else v)
                  for fy, v in r.values.items()}
        if _all_zero(values):
            continue
        li = _make_item(r.label, values, "BS", g.title, "pdf-note", r.page, r.flags, fy_columns, multiplier)
        li.statement_tag, li.category, li.is_new, li.confidence = tag, "", False, "Matched"
        li.flags = [f for f in li.flags if f not in ("unplaced", "sum_mismatch")]
        result.items.append(li)
    return True


def _map_reserves(groups, fy_columns, multiplier, result, face_reserves=None, note_scale=1):
    found = False
    for options in groups.values():
        for g in options:
            if not _RESERVES.search(g.title) and _OWNERS_NOTE.search(g.title):
                found = _map_owners_capital(g, fy_columns, multiplier, result, note_scale) or found
                continue
            if not _RESERVES.search(g.title):
                continue
            found = True
            # the reserves note may be in rupees while the balance sheet is in thousands
            scale = None
            if face_reserves is not None:
                closing = next((r for r in reversed(g.items) if re.search(r"closing", r.label, re.I) and _has_values(r)), None)
                basis = g.total.values if (g.total is not None and _has_values(g.total)) else (closing.values if closing else {})
                scale = _scale_between(basis, face_reserves.values, fy_columns)
            g = _rescaled(g, scale if scale is not None else note_scale)
            less_sign = _less_sign(g, fy_columns)
            for r in g.items:
                low = r.label.lower()
                if re.search(r"opening", low):
                    tag = "Opening Reserves"
                elif re.search(r"securities\s*premium", low):
                    tag = "Securities Premium"
                elif _PROFIT_LINE.search(low):
                    continue  # derived by the template from the P&L / roll-forward
                else:
                    tag = "Reserves Adjustments"
                values = dict(r.values)
                if tag == "Reserves Adjustments" and low.startswith("less"):
                    values = {fy: (less_sign(fy) * abs(v) if v is not None else None) for fy, v in values.items()}
                if _all_zero(values):
                    continue
                li = _make_item(r.label, values, "BS", g.title, "pdf-note", r.page, r.flags, fy_columns, multiplier)
                li.statement_tag, li.category, li.is_new = tag, "", False
                li.confidence = "Matched"
                li.flags = [f for f in li.flags if f not in ("unplaced", "sum_mismatch")]
                result.items.append(li)
    if not found:
        result.log.append("WARNING: no 'Reserves and Surplus' note found, so opening reserves were not captured; "
                          "the Balance Sheet will not tally until they are entered.")


def map_extraction(extraction, unit_multiplier=None):
    fy = list(extraction.fy_columns)
    mult = extraction.unit_multiplier if unit_multiplier is None else unit_multiplier
    rows = extraction.rows
    result = MappingResult()
    if not fy:
        result.log.append("WARNING: financial-year columns were not detected; nothing could be mapped.")
        return result

    face = [r for r in rows if r.statement in ("BS", "P&L")]
    note_rows = [r for r in rows if r.statement == "NOTE" and "sub_schedule" not in r.flags]
    groups = _build_groups(note_rows)
    parents_with_children = {r.heading for r in face if _has_values(r) and not r.note_ref}

    # the unit of the notes relative to the statements (Rs. against Rs. '000 ...), learnt where both are readable
    found = []
    for F in face:
        g = _pick_group(groups, F) if F.note_ref else None
        if g is not None and _has_values(F):
            nv = _note_reference_values(g, fy)
            if not any(nv.get(y) and F.values.get(y) for y in fy):
                continue                  # nothing to compare (nil lines say nothing about the unit)
            s = _scale_between(nv, F.values, fy)
            if s is not None:
                found.append(s)
    note_scale = max(set(found), key=found.count) if found else 1

    face_reserves = next((r for r in face if r.statement == "BS" and (_RESERVES.search(r.label) or _OWNERS_CAPITAL.search(r.label))
                          and _has_values(r)), None)
    _map_reserves(groups, fy, mult, result, face_reserves, note_scale)

    assets_side, seen_total = None, False   # where on the balance sheet the line sits
    replaced_parents = set()
    # a heading printed without figures ("TRADE PAYABLES 8") whose scanned sub-lines carry the figures: when its
    # note is digital, its sub-lines do not add up to the note but the note's own lines do, use the note's lines
    for heading in dict.fromkeys(r.heading for r in face if r.statement == "BS" and r.heading):
        children = [r for r in face if r.heading == heading and _has_values(r) and not r.note_ref and "ocr" in r.flags]
        if not children or any(r.label == heading and _has_values(r) for r in face):
            continue
        pseudo = replace(children[0], label=heading, note_ref=re.search(r"(\d{1,2})\W*$", heading).group(1)
                         if re.search(r"(\d{1,2})\W*$", heading) else "")
        pg = _pick_group(groups, pseudo)
        if pg is None or not _is_digital(pg) or pg.total is None or not _has_values(pg.total):
            continue
        child_sum = {y: _sum_for(children, y) for y in fy}
        pg = _rescaled(pg, note_scale)
        rounding = 0.5 * max(len(children), 1) + 0.5 if note_scale < 1 else TOL
        if all(abs(child_sum[y] - (pg.total.values.get(y) or 0.0)) <= rounding for y in fy):
            continue                       # the scanned sub-lines agree with the note: keep them
        tol = TOL if note_scale == 1 else max(TOL, 0.5 * len(pg.items) + 0.5)
        note_items = _with_deductions(pg.items, pg.total.values, fy, tol) if len(pg.items) >= 2 else None
        lines = [r for r in (note_items or [pg.total]) if not _all_zero(r.values)]
        for r in lines:
            item = _make_item(r.label if note_items else pg.title, r.values, "BS", pg.title, "pdf-note", r.page,
                              ["from_digital_note"], fy, mult)
            if not item.statement_tag:
                key = schedule_for(_make_item(pg.title, pg.total.values, "BS", heading, "pdf-face", r.page, [], fy, mult))
                if key:
                    item.schedule_override = key
                    item.flags = [f for f in item.flags if f != "unplaced"]
            result.items.append(item)
        replaced_parents.add(heading)
        result.log.append(f"{heading}: taken from Note {pg.note_no} (digital); the scanned sub-lines did not add up to it")
    for F in face:
        label = F.label
        if F.statement == "BS":
            section = F.heading or ""
            if re.search(r"\bassets?\b", section, re.I):
                assets_side = True
            elif re.search(r"liabilit|equity|shareholder|net\s*worth", section, re.I):
                assets_side = False
            elif seen_total:
                assets_side = True        # Schedule III order: equity & liabilities, their Total, then assets
            if F.is_total and re.fullmatch(r"total\W*", label.strip(), re.I):
                seen_total = True
        if F.is_total or _RESERVES.search(label) or (F.statement == "BS" and _OWNERS_CAPITAL.search(label)):
            continue
        if _SUMMARY_LINE.search(label) and not re.search(r"other\s*income", label, re.I):
            result.skipped.append((label, "summary / per-share line"))
            continue
        has_vals = _has_values(F)
        if not has_vals and not F.note_ref:
            continue
        if F.statement in ("BS", "P&L") and F.heading in replaced_parents and not F.note_ref:
            continue          # a scanned sub-line whose heading was taken from its digital note instead
        if not has_vals and label in parents_with_children:
            pg = _pick_group(groups, F)
            if pg is not None and "ocr" in F.flags and _is_digital(pg) and pg.total is not None and _has_values(pg.total):
                # the sub-lines are an OCR reading; the digital note is exact: use the note's lines
                s = note_scale
                children = [r for r in face if r.heading == label and _has_values(r) and r is not F]
                child_sum = {y: _sum_for(children, y) for y in fy}
                s = _scale_between(pg.total.values, child_sum, fy) or note_scale
                pg = _rescaled(pg, s)
                tol = TOL if s == 1 else max(TOL, 0.5 * len(pg.items) + 0.5)
                note_items = _with_deductions(pg.items, pg.total.values, fy, tol) if len(pg.items) >= 2 else None
                lines = [r for r in (note_items or [pg.total]) if not _all_zero(r.values)]
                key_item = _side_rules(_make_item(label, pg.total.values, F.statement, F.heading, "pdf-face", F.page, [],
                                                  fy, mult), F.statement, label, assets_side)
                for r in lines:
                    item = _make_item(r.label if note_items else label, r.values, F.statement, pg.title, "pdf-note",
                                      r.page, ["from_digital_note"], fy, mult)
                    if schedule_for(key_item) and schedule_for(item) != schedule_for(key_item):
                        item.schedule_override = schedule_for(key_item)
                        item.flags = [f for f in item.flags if f != "unplaced"]
                    result.items.append(item)
                replaced_parents.add(label)
                result.log.append(f"{label}: taken from Note {pg.note_no} (digital) instead of the scanned sub-lines")
                continue
            result.skipped.append((label, "parent line; its sub-lines carry the figures"))
            continue

        group = _pick_group(groups, F)        # by note number, or by title when the number is unreadable
        scale = 1
        if group is not None:
            scale = _scale_between(_note_reference_values(group, fy), F.values, fy)
            if scale is None:
                scale = note_scale          # the unit the notes use elsewhere in this document
            group = _rescaled(group, scale)
        heading = F.heading
        flags = list(F.flags)
        same_note = group is not None and fuzz.token_set_ratio(_plain(label), _plain(group.title)) >= 75
        if same_note and group.total is not None and "ocr" in F.flags and _is_digital(group):
            # the statement is a scan but the note is digital text: the note's exact total beats an OCR reading
            face_values, changed = dict(F.values), []
            for y in fy:
                nv, fv = group.total.values.get(y), face_values.get(y)
                if fv == 0 or not nv:
                    continue        # a Nil on either side says nothing reliable: never overwrite with it
                if nv is not None and (fv is None or abs(fv - nv) > max(1.0 if scale < 1 else TOL, 0.002 * abs(nv))):
                    face_values[y] = nv
                    changed.append(y)
            if changed:
                flags.append("from_digital_note")
                result.log.append(f"{label}: {', '.join(changed)} taken from Note {F.note_ref} (digital) instead of the scanned "
                                  f"statement's reading")
                F = replace(F, values=face_values)
                has_vals = True
        reference = _reference_values(F, group, fy)
        # notes in rupees against statements in thousands: each note line was rounded, allow for that
        tol = TOL if scale == 1 else max(TOL, 0.5 * len(group.items) + 0.5)

        note_items = _with_deductions(group.items, reference, fy, tol) if (group is not None and reference) else None
        if group is not None and len(group.items) >= 2 and note_items is not None:
            # the statement line decides the schedule; a note item only chooses the row inside it, so an
            # "Interest from bank deposits" inside the Other Income note stays in Other Income
            face_key = schedule_for(_side_rules(
                _make_item(label, reference, F.statement, heading, "pdf-face", F.page, [], fy, mult),
                F.statement, label, assets_side))
            for r in note_items:
                values = dict(r.values)
                if _all_zero(values):
                    continue
                kept = [f for f in r.flags if f not in ("sum_mismatch", "differs_from_note")]
                item = _make_item(r.label, values, F.statement, group.title, "pdf-note", r.page, kept, fy, mult)
                own_key = schedule_for(item)
                if face_key and own_key != face_key and not any({own_key, face_key} <= g for g in _SAME_IS_LINE):
                    item.schedule_override = face_key
                    item.flags = [f for f in item.flags if f != "unplaced"] + ["placed_under_statement_heading"]
                result.items.append(item)
            result.log.append(f"{label}: itemised from Note {F.note_ref} ({len(group.items)} lines, ties to the statement)")
            continue

        if has_vals:
            if _all_zero(F.values):
                result.skipped.append((label, "nil in all years"))
                continue
            if group is not None and len(group.items) >= 2:
                flags.append("breakup_unreliable")
                result.log.append(f"{label}: Note {F.note_ref} breakup does not tie to the statement, so the statement figure is used")
            parent_heading = heading
            under_inventory_change = re.search(r"changes?\s*in\s*inventor", heading, re.I) and re.search(
                r"finished|work.in.progress|\bwip\b|stock.in.trade", label, re.I)
            label_for_resolve = f"Changes in inventories {label}" if under_inventory_change else None
            if not under_inventory_change and re.search(r"changes?\s*in\s*inventor", heading, re.I):
                parent_heading = ""
            face_values = dict(F.values)
            if group is not None and group.total is not None:
                # a statement figure OCR could not read is taken from the note's own total
                missing = [y for y in fy if face_values.get(y) is None and group.total.values.get(y) is not None]
                for y in missing:
                    face_values[y] = group.total.values[y]
                if missing:
                    flags.append("value_from_note_total")
            item = _make_item(label, face_values, F.statement, parent_heading, "pdf-face", F.page,
                              flags, fy, mult, resolve_label=label_for_resolve)
            _side_rules(item, F.statement, label, assets_side)
            if group is not None and group.total is not None:
                # the note's own total, kept as a possible correction when the face figure looks misread
                for fy_ in fy:
                    face_v, note_v = F.values.get(fy_), group.total.values.get(fy_)
                    if face_v is not None and note_v is not None and abs(face_v - note_v) > TOL:
                        item.alt_values[fy_] = note_v * mult
            result.items.append(item)
        elif group is not None and group.total is not None and _has_values(group.total):
            vals = {fy_: group.total.values.get(fy_) for fy_ in fy}
            if not _all_zero(vals):
                flags.append("value_from_note_total")
                result.items.append(_make_item(label, vals, F.statement, heading, "pdf-note", group.total.page,
                                               flags, fy, mult))
                result.log.append(f"{label}: figures taken from the Total of Note {F.note_ref}")
        else:
            result.skipped.append((label, "no figures found"))

    return result
