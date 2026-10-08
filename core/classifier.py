"""
Rule-based functional expense classification.

Only applies to P&L rows that don't already carry a Statement Tag of
"Revenue" (revenue isn't an expense to classify) -- everything else on the
P&L gets scored against config.classification_rules.CATEGORY_KEYWORDS.
Items with no keyword hit, or a tie between categories, are left as
"Uncertain -- Needs Review" per Section 5/6 of the instructions: never force
a classification when confidence is low.
"""

from config.classification_rules import (
    CATEGORY_KEYWORDS,
    PNL_DIRECT_TAGS,
    UNCERTAIN_LABEL,
)

# Note: unmatched items are flagged UNCERTAIN rather than auto-assigned to
# "Other Operating Expenses" -- Section 3 of the instructions treats that
# category as a considered judgment call ("document rationale"), which a
# keyword miss cannot substitute for.


def _score_item(name_lower):
    scores = {}
    matches = {}
    for category, keywords in CATEGORY_KEYWORDS.items():
        hits = [kw for kw in keywords if kw in name_lower]
        if hits:
            scores[category] = len(hits)
            matches[category] = hits
    return scores, matches


def keyword_category(label):
    """Section-3 keyword rules only. Returns (category | None, matched keywords); None if nothing or a tie."""
    scores, matches = _score_item(label.lower())
    if not scores:
        return None, []
    best = max(scores.values())
    winners = [c for c, sc in scores.items() if sc == best]
    if len(winners) != 1:
        return None, sorted(set(sum((matches[w] for w in winners), [])))
    return winners[0], matches[winners[0]]


def classify_item(line_item):
    """Mutates and returns the LineItem with .category/.confidence/.matched_keywords set."""
    if line_item.statement != "P&L" or line_item.statement_tag in PNL_DIRECT_TAGS:
        line_item.category = ""
        line_item.confidence = ""
        return line_item

    name_lower = line_item.line_item.lower()
    scores, matches = _score_item(name_lower)

    if not scores:
        line_item.category = UNCERTAIN_LABEL
        line_item.confidence = "Uncertain"
        line_item.matched_keywords = []
        return line_item

    best_score = max(scores.values())
    winners = [c for c, s in scores.items() if s == best_score]

    if len(winners) > 1:
        # Genuine ambiguity between categories -- flag rather than guess.
        line_item.category = UNCERTAIN_LABEL
        line_item.confidence = f"Uncertain (tied between: {', '.join(winners)})"
        line_item.matched_keywords = sorted(set(sum((matches[w] for w in winners), [])))
        return line_item

    winner = winners[0]
    line_item.category = winner
    line_item.confidence = "Matched"
    line_item.matched_keywords = matches[winner]
    return line_item


def classify_all(line_items):
    return [classify_item(li) for li in line_items]
