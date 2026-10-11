"""CSL style name -> Zotero style id resolution + validation.

Alias targets are Zotero style ids (http://www.zotero.org/styles/<slug>).
Slugs were verified this session against freshly downloaded CSL files
(<info><id>) where available; see PREVIEW_FILE_FALLBACKS for slugs whose
offline preview file has a different name.
"""

from __future__ import annotations

import re

STYLE_PREFIX = "http://www.zotero.org/styles/"

# CLI short names -> Zotero style ids (spec: apa/mla/harvard/ieee/vancouver/
# chicago/nature or any CSL id)
STYLE_ALIASES: dict[str, str] = {
    "apa": "apa",
    "mla": "modern-language-association",
    "harvard": "harvard-cite-them-right",
    "ieee": "ieee",
    "vancouver": "vancouver",
    "chicago": "chicago-notes-bibliography-17th-edition",
    "nature": "nature",
    # convenience aliases (file-backed ids, verified against <info><id>)
    "elsevier-harvard": "elsevier-harvard",
    "chicago-author-date": "chicago-author-date-16th-edition",
    "nlm": "nlm-citation-sequence",
}

# when the style id has no same-named offline .csl file, use this file for
# PROVISIONAL previews only (prefs always carry the canonical style id)
PREVIEW_FILE_FALLBACKS: dict[str, str] = {
    "vancouver": "nlm-citation-sequence",
    "chicago-note-bibliography": "chicago-notes-bibliography-17th-edition",
    "chicago-note-bibliography-16th-edition":
        "chicago-notes-bibliography-16th-edition",
}

SLUG_RE = re.compile(r"^[a-z0-9]+(?:-[a-z0-9.]+)*$")


class StyleError(ValueError):
    """Raised for unusable --style values; message lists valid options."""


def resolve_style_id(name_or_id: str) -> str:
    """
    Accepts a short alias ('harvard'), a bare slug
    ('harvard-cite-them-right') or a full style URL, and returns the full
    Zotero style URL.  Raises StyleError with the valid short names for
    clearly unusable input.
    """
    s = (name_or_id or "").strip()
    if not s:
        raise StyleError(_style_error_message("empty style name"))
    if s.startswith("http://www.zotero.org/styles/"):
        slug = s[len(STYLE_PREFIX):]
        if not SLUG_RE.match(slug):
            raise StyleError(_style_error_message(f"bad style URL {name_or_id!r}"))
        return s
    if s.startswith(("http://", "https://")):
        raise StyleError(_style_error_message(
            f"{name_or_id!r} is not a zotero.org style URL"))
    low = s.lower()
    if low in STYLE_ALIASES:
        return STYLE_PREFIX + STYLE_ALIASES[low]
    if SLUG_RE.match(low):
        return STYLE_PREFIX + low      # any CSL id / slug passes through
    raise StyleError(_style_error_message(f"unknown style {name_or_id!r}"))


def style_slug(style_id: str) -> str:
    return style_id[len(STYLE_PREFIX):] if style_id.startswith(STYLE_PREFIX) \
        else style_id


def _style_error_message(problem: str) -> str:
    names = ", ".join(sorted(
        {"apa", "mla", "harvard", "ieee", "vancouver", "chicago", "nature"}
    ))
    return (
        f"{problem}. Valid short names: {names}; "
        "or any CSL style slug (e.g. 'elsevier-harvard'); "
        f"or a full '{STYLE_PREFIX}<slug>' URL."
    )
