"""Safe Excel worksheet naming.

Excel limits a worksheet name to 31 characters, rejects ``[ ] : * ? / \\``,
forbids a leading or trailing apostrophe, reserves "History", and compares
names case-insensitively when deciding whether two sheets clash.

The workbook is generated per warehouse, so names are composed at runtime
("Mississauga - Calgary CO-OP") and can easily break those limits. Everything
that creates a sheet goes through :class:`SheetNamer`, which returns a name
that is always legal, always unique, and still readable.
"""
from __future__ import annotations

import re
from collections.abc import Iterable

MAX_SHEET_NAME = 31

# Characters Excel rejects outright inside a worksheet name.
_INVALID_CHARS = re.compile(r"[\[\]:*?/\\]")
_WS = re.compile(r"\s+")

# Excel keeps this name for the shared-workbook change history.
_RESERVED = {"history"}

# Room reserved after a site label for the longest thing that follows it:
# " - " plus a customer report name. Sized so one warehouse reads the same on
# every one of its sheets rather than being shortened differently each time.
_SITE_SUFFIX_BUDGET = 16


def sanitize_sheet_name(name: object) -> str:
    """Strip everything Excel forbids from *name*. Never returns an empty string."""
    text = _INVALID_CHARS.sub(" ", "" if name is None else str(name))
    text = _WS.sub(" ", text).strip()
    text = text.strip("'").strip()
    return text or "Sheet"


def shorten(text: str, limit: int) -> str:
    """Trim *text* to *limit* characters, preferring a whole-word boundary."""
    text = (text or "").strip()
    if len(text) <= limit:
        return text
    if limit <= 0:
        return ""
    kept = ""
    for word in text.split():
        candidate = f"{kept} {word}".strip()
        if len(candidate) > limit:
            break
        kept = candidate
    # Only keep the word-boundary trim while it still says something useful;
    # otherwise a hard cut carries more meaning than one short word.
    if len(kept) >= max(3, limit // 2):
        return kept
    return text[:limit].strip()


def site_label_map(sites: Iterable[str]) -> dict[str, str]:
    """Map each detected site to the short label used across all its sheets.

    Shortening a warehouse name once, here, is what keeps "Calgary Issue
    Tracker", "Calgary Orders" and "Calgary - AMZ" recognizable as one block.
    """
    labels: dict[str, str] = {}
    used: set[str] = set()
    for site in sites:
        clean = sanitize_sheet_name(site)
        label = shorten(clean, MAX_SHEET_NAME - _SITE_SUFFIX_BUDGET)
        if label.casefold() in used:  # two long site names trimmed to the same label
            stem = shorten(clean, MAX_SHEET_NAME - _SITE_SUFFIX_BUDGET - 2)
            n = 2
            while f"{stem} {n}".casefold() in used:
                n += 1
            label = f"{stem} {n}"
        used.add(label.casefold())
        labels[site] = label
    return labels


class SheetNamer:
    """Hands out unique, Excel-legal worksheet names in creation order."""

    def __init__(self) -> None:
        self._used: set[str] = set(_RESERVED)

    def allocate(self, name: str) -> str:
        """Return a legal, unused worksheet name closest to *name*."""
        base = shorten(sanitize_sheet_name(name), MAX_SHEET_NAME)
        candidate = base
        n = 2
        while candidate.casefold() in self._used:
            suffix = f" ({n})"
            candidate = f"{shorten(base, MAX_SHEET_NAME - len(suffix))}{suffix}"
            n += 1
        self._used.add(candidate.casefold())
        return candidate
