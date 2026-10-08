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
from dataclasses import dataclass, field

from rapidfuzz import fuzz

from core.models import LineItem
from core.resolver import resolve

TOL = 0.05
_SUMMARY_LINE = re.compile(
    r"earning|per\s*equity\s*share|exceptional|extraordinary|profit|loss\b|net\s*worth|^(?:basic|diluted)|^total\b",
    re.I,
)
_RESERVES = re.compile(r"reserves?\s*(?:and|&)\s*surplus", re.I)


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


def _pick_group(groups, face_row):
    options = groups.get(face_row.note_ref, [])
    if not options:
        return None
    best = max(options, key=lambda g: fuzz.token_set_ratio(face_row.label.lower(), g.title.lower()))
    if fuzz.token_set_ratio(face_row.label.lower(), best.title.lower()) < 45:
        return None
    return best


def _sum_for(items, fy):
    return sum((i.values.get(fy) or 0.0) for i in items)


def _reference_values(face_row, group, fy_columns):
    """Figures the note's items must add up to: the face figures, else the note's own Total."""
    if all(face_row.values.get(fy) is not None for fy in fy_columns):
        return {fy: face_row.values[fy] for fy in fy_columns}
    if group is not None and group.total is not None and any(v is not None for v in group.total.values.values()):
        return {fy: group.total.values.get(fy) for fy in fy_columns}
    return None


def _items_tie(items, reference, fy_columns):
    compared = 0
    for fy in fy_columns:
        ref = reference.get(fy)
        if ref is None:
            continue
        compared += 1
        if abs(_sum_for(items, fy) - ref) > TOL:
            return False
    return compared > 0


def _scaled(values, fy_columns, multiplier):
    return {fy: (values[fy] * multiplier if values.get(fy) is not None else None) for fy in fy_columns}


def _all_zero(values):
    return all((v is None or abs(v) < 1e-9) for v in values.values())


def _make_item(label, values, statement, heading, source, page, flags, fy_columns, multiplier, resolve_label=None):
    res = resolve(resolve_label or label, heading, statement)
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
            if v is None or re.search(r"opening|closing|net\s*profit|profit\s*for", low):
                continue
            if low.startswith("less"):
                less_abs += abs(v)
            else:
                fixed += v
        profit = row_value(r"net\s*profit|profit\s*for", fy) or 0.0
        base = opening + profit + fixed
        if abs(base - less_abs - closing) <= TOL:
            signs[fy] = -1
        elif abs(base + less_abs - closing) <= TOL:
            signs[fy] = 1
    return lambda fy: signs.get(fy, -1)


def _map_reserves(groups, fy_columns, multiplier, result):
    found = False
    for options in groups.values():
        for g in options:
            if not _RESERVES.search(g.title):
                continue
            found = True
            less_sign = _less_sign(g, fy_columns)
            for r in g.items:
                low = r.label.lower()
                if re.search(r"opening", low):
                    tag = "Opening Reserves"
                elif re.search(r"securities\s*premium", low):
                    tag = "Securities Premium"
                elif re.search(r"net\s*profit|closing|profit\s*for", low):
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

    _map_reserves(groups, fy, mult, result)

    for F in face:
        label = F.label
        if F.is_total or _RESERVES.search(label):
            continue
        if _SUMMARY_LINE.search(label) and not re.search(r"other\s*income", label, re.I):
            result.skipped.append((label, "summary / per-share line"))
            continue
        has_vals = _has_values(F)
        if not has_vals and not F.note_ref:
            continue
        if not has_vals and label in parents_with_children:
            result.skipped.append((label, "parent line; its sub-lines carry the figures"))
            continue

        group = _pick_group(groups, F) if F.note_ref else None
        heading = F.heading
        reference = _reference_values(F, group, fy)
        flags = list(F.flags)

        if group is not None and len(group.items) >= 2 and reference and _items_tie(group.items, reference, fy):
            for r in group.items:
                values = dict(r.values)
                if _all_zero(values):
                    continue
                kept = [f for f in r.flags if f not in ("sum_mismatch", "differs_from_note")]
                result.items.append(_make_item(r.label, values, F.statement, group.title, "pdf-note", r.page,
                                               kept, fy, mult))
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
            item = _make_item(label, F.values, F.statement, parent_heading, "pdf-face", F.page,
                              flags, fy, mult, resolve_label=label_for_resolve)
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
