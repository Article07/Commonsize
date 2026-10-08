"""
Decides where a line item belongs.

Input is the item's label and the heading the PDF presents it under (a note
title such as "Other Current Assets", or a face-statement section). Output
says whether the item is KNOWN (matched to the taxonomy in
`CS format dropdown.xlsx`, or to a sub-category alias) or NEW.

A NEW item is still placed: under the schedule that the PDF's own heading
points to ("exactly as it is classified in the PDF financials"), flagged so
the writer can highlight it yellow and the IRL can ask management about it.
Only an item whose heading also resolves to nothing stays unplaced.

Matching order (first hit wins):
  1. taxonomy item, exact name
  2. taxonomy item, fuzzy (token-set similarity) -- skipped when two items that
     point to different places score about the same (e.g. "Insurance")
  3. sub-category alias on the label ("Trade receivables", "Other income" ...)
  4. the project's keyword rules for expense categories (Section 3 of the
     instructions), for P&L items the taxonomy does not name
  5. alias on the HEADING -> NEW item under that heading
"""

import re
from dataclasses import dataclass

from rapidfuzz import fuzz

from config.classification_rules import PNL_DIRECT_TAGS
from config.taxonomy import BS_ALIASES, GENERIC_FILLERS, GENERIC_ITEM_NAMES, MATCHABLE_TAXONOMY, PNL_ALIASES

FUZZY_CUTOFF = 88
AMBIGUITY_GAP = 4


@dataclass
class Resolution:
    statement: str = ""      # "P&L" | "BS" | "" (unresolved)
    tag: str = ""
    category: str = ""
    matched_item: str = ""
    score: float = 0.0
    known: bool = False      # True: label matched taxonomy/alias.  False: new item (or unresolved)
    via: str = "none"        # taxonomy-exact | taxonomy-fuzzy | alias | keyword | heading | none
    heading_used: str = ""
    row_label: str = ""      # template row to use when only a total is available

    @property
    def resolved(self):
        return bool(self.tag or self.category)


def normalise(text):
    text = re.sub(r"\(.*?\)", " ", text.lower())
    text = text.replace("&", " and ")
    text = re.sub(r"[^a-z0-9 ]", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def _build_index():
    index = {"P&L": [], "BS": []}
    for entry in MATCHABLE_TAXONOMY:
        match_name = normalise(entry.item)
        if match_name in GENERIC_ITEM_NAMES or len(match_name) < 3:
            continue
        stmt = "P&L" if entry.target_kind == "category" else "BS"
        index[stmt].append((match_name, entry))
    return index


_INDEX = _build_index()


def _from_target(statement, kind, target, **kw):
    if statement == "P&L":
        is_tag = kind == "tag" or target in PNL_DIRECT_TAGS
        return Resolution(statement, tag=target if is_tag else "", category="" if is_tag else target, **kw)
    return Resolution(statement, tag=target, **kw)


def _taxonomy_match(statement, n):
    candidates = []
    for match_name, entry in _INDEX[statement]:
        if n == match_name:
            return _from_target(statement, entry.target_kind, entry.target,
                                matched_item=entry.item, score=100.0, known=True, via="taxonomy-exact")
        score = fuzz.token_set_ratio(n, match_name)
        if score >= FUZZY_CUTOFF:
            candidates.append((score, fuzz.ratio(n, match_name), entry))
    if not candidates:
        return None
    candidates.sort(key=lambda c: (c[0], c[1]), reverse=True)
    top_score, top_ratio, top = candidates[0]
    rivals = [c for c in candidates[1:] if c[2].target != top.target and top_score - c[0] <= AMBIGUITY_GAP
              and top_ratio - c[1] <= 8]
    if rivals:
        return None  # same-strength matches pointing at different places: do not guess
    return _from_target(statement, top.target_kind, top.target,
                        matched_item=top.item, score=float(top_score), known=True, via="taxonomy-fuzzy")


def _alias_match(statement, text, known, via, heading_used=""):
    table = PNL_ALIASES if statement == "P&L" else BS_ALIASES
    low = text.lower()
    for pattern, kind, target, row_label in table:
        if re.search(pattern, low):
            return _from_target(statement, kind, target, matched_item=target, score=0.0,
                                known=known, via=via, heading_used=heading_used, row_label=row_label)
    return None


def _keyword_match(label):
    from core.classifier import keyword_category

    category, hits = keyword_category(label)
    if category:
        return Resolution("P&L", category=category, matched_item=", ".join(hits), known=True, via="keyword")
    return None


def resolve(label, heading="", statement=None):
    n = normalise(label)
    statements = [statement] if statement in ("P&L", "BS") else ["P&L", "BS"]
    for st in statements:
        found = _taxonomy_match(st, n) or _alias_match(st, label, True, "alias")
        if not found and st == "P&L":
            found = _keyword_match(label)
        if found:
            return found
    if heading:
        is_filler = n in GENERIC_FILLERS
        for st in statements:
            found = _alias_match(st, heading, is_filler, "heading", heading_used=heading)
            if found:
                return found
    return Resolution(statement=statement or "", via="none")
