"""
Preview text rendering for the Word fields.

Primary path: citeproc-py with an offline .csl style file (found via
ZOTEROCONV_CSL_DIR, the repo ./csl directory -- populated from the official
citation-style-language repositories -- or a local Zotero styles folder).

Fallback path (user directive): when the target style file is NOT available
offline, previews are clearly marked PROVISIONAL approximations; Zotero
rewrites every preview on Refresh, so they never need to be exact.
"""

from __future__ import annotations

import os
import re
from pathlib import Path

from zoteroconv.styles import PREVIEW_FILE_FALLBACKS, style_slug

REPO_CSL_DIR = Path(__file__).resolve().parents[1] / "csl"
DEPENDENT_PARENT_RE = re.compile(
    r'rel="independent-parent"[^>]*href="http://www\.zotero\.org/styles/([^"]+)"'
)


def candidate_style_dirs(extra_dirs=None) -> list[Path]:
    dirs: list[Path] = []
    for d in extra_dirs or []:
        dirs.append(Path(d))
    env = os.environ.get("ZOTEROCONV_CSL_DIR")
    if env:
        dirs.extend(Path(p) for p in env.split(os.pathsep) if p)
    dirs.append(REPO_CSL_DIR)
    home = Path.home() / ".zotero" / "zotero"
    if home.is_dir():
        dirs.extend(sorted(home.glob("*/styles")))
    return dirs


def find_style_file(slug: str, extra_dirs=None) -> Path | None:
    """Locate <slug>.csl offline; follows dependent-style redirects once and
    consults PREVIEW_FILE_FALLBACKS for renamed styles."""
    tried: set[str] = set()
    queue = [slug]
    fallback = PREVIEW_FILE_FALLBACKS.get(slug)
    if fallback:
        queue.append(fallback)
    for cand in queue:
        if cand in tried:
            continue
        tried.add(cand)
        for d in candidate_style_dirs(extra_dirs):
            p = d / f"{cand}.csl"
            if p.is_file():
                text = p.read_text(encoding="utf-8", errors="replace")[:4000]
                m = DEPENDENT_PARENT_RE.search(text)
                if m and m.group(1) not in tried:
                    continue    # dependent stub: try the parent instead
                return p
        parent = None
        for d in candidate_style_dirs(extra_dirs):
            p = d / f"{cand}.csl"
            if p.is_file():
                m = DEPENDENT_PARENT_RE.search(
                    p.read_text(encoding="utf-8", errors="replace")[:4000])
                if m:
                    parent = m.group(1)
        if parent and parent not in tried:
            queue.append(parent)
    return None


def _surname(csl: dict) -> str:
    for a in csl.get("author") or []:
        if a.get("family"):
            return str(a["family"])
        if a.get("name"):
            return str(a["name"])
    return ""


def _year(csl: dict) -> str:
    try:
        y = (csl.get("issued") or {}).get("date-parts", [[None]])[0][0]
        return str(y) if y is not None else "n.d."
    except (IndexError, TypeError, AttributeError):
        return "n.d."


def fallback_inline(items: list[dict]) -> str:
    """PROVISIONAL author-date approximation, used only when the style file
    is unavailable.  Zotero overwrites it on Refresh."""
    rendered = []
    for it in items:
        sur = _surname(it)
        year = _year(it)
        if not sur:
            rendered.append(str(it.get("title") or "")[:40] or year)
        else:
            rendered.append(f"{sur}, {year}")
    return "(" + "; ".join(rendered) + ")"


def fallback_entry(csl: dict) -> str:
    """PROVISIONAL bibliography entry approximation."""
    authors = []
    for a in csl.get("author") or []:
        if a.get("family"):
            given = a.get("given") or ""
            initials = "".join(
                p[0] + "." for p in re.split(r"[\s-]+", given) if p)
            authors.append(f"{a['family']}, {initials}".rstrip(", "))
        elif a.get("name"):
            authors.append(str(a["name"]))
    parts = []
    if authors:
        parts.append(", ".join(authors))
    year = _year(csl)
    if year != "n.d.":
        parts.append(f"({year})")
    if csl.get("title"):
        parts.append(f"{csl['title']}.")
    if csl.get("container-title"):
        seg = str(csl["container-title"])
        if csl.get("volume"):
            seg += f", {csl['volume']}"
            if csl.get("issue"):
                seg += f"({csl['issue']})"
        parts.append(seg + ".")
    elif csl.get("publisher"):
        place = csl.get("publisher-place")
        parts.append((f"{place}: " if place else "") + f"{csl['publisher']}.")
    if csl.get("page"):
        parts.append(f"pp. {csl['page']}.")
    if csl.get("DOI"):
        parts.append(f"https://doi.org/{csl['DOI']}")
    elif csl.get("URL"):
        parts.append(str(csl["URL"]))
    return " ".join(parts)


class PreviewRenderer:
    """Renders citation-cluster previews and bibliography entry previews.

    ``render(clusters, bib_items)`` returns ``(inline_previews, bib_entries)``
    where clusters is a list of per-field CSL item lists (in citation order)
    and bib_items are the CSL dicts destined for the bibliography (every
    cluster item must be among them)."""

    def __init__(self, style_id: str, locale: str = "en-US", style_dirs=None):
        self.style_id = style_id
        self.locale = locale or "en-US"
        self.slug = style_slug(style_id)
        self.notes: list[str] = []
        self.style_path = find_style_file(self.slug, style_dirs)
        self.engine = "provisional-fallback"
        if self.style_path is None:
            self.notes.append(
                f"CSL style file for {self.style_id!r} not available offline; "
                "preview texts are PROVISIONAL approximations (Zotero "
                "overwrites all previews on Refresh)"
            )
        else:
            try:
                import citeproc  # noqa: F401
                self.engine = f"citeproc-py ({self.style_path.name})"
            except ImportError:
                self.style_path = None
                self.engine = "provisional-fallback"
                self.notes.append(
                    "citeproc-py not installed; preview texts are PROVISIONAL "
                    "approximations (Zotero overwrites them on Refresh)"
                )

    @property
    def provisional(self) -> bool:
        return self.style_path is None

    # -- sanitising ------------------------------------------------------------

    @staticmethod
    def _sanitize(csl: dict, key: str) -> dict:
        out = dict(csl)
        out["id"] = key
        for var in ("author", "editor"):
            people = out.get(var)
            if not people:
                continue
            fixed = []
            for p in people:
                if p.get("name") and not p.get("family"):
                    # citeproc-py does not support single-field 'name'
                    fixed.append({"family": p["name"], "given": ""})
                else:
                    fixed.append(p)
            out[var] = fixed
        return out

    # -- rendering ---------------------------------------------------------------

    def render(self, clusters: list[list[dict]],
               bib_items: list[dict]) -> tuple[list[str], list[str]]:
        if self.provisional:
            return (
                [fallback_inline(c) for c in clusters],
                [fallback_entry(b) for b in bib_items],
            )
        try:
            return self._render_citeproc(clusters, bib_items)
        except Exception as exc:                # never fail the conversion
            self.notes.append(
                f"citeproc-py rendering failed ({type(exc).__name__}: {exc}); "
                "falling back to PROVISIONAL previews"
            )
            self.engine = "provisional-fallback"
            return (
                [fallback_inline(c) for c in clusters],
                [fallback_entry(b) for b in bib_items],
            )

    def _render_citeproc(self, clusters, bib_items):
        import warnings

        from citeproc import (Citation, CitationItem, CitationStylesBibliography,
                              CitationStylesStyle)
        from citeproc.formatter import plain
        from citeproc.source.json import CiteProcJSON

        keys = {}
        sanitized = []
        for i, it in enumerate(bib_items):
            k = f"z{i}"
            keys[id(it)] = k
            sanitized.append(self._sanitize(it, k))
        with warnings.catch_warnings():
            # citeproc-py warns about CSL keys it does not model (license,
            # journalAbbreviation, single-field names...); previews are
            # provisional, so these warnings are noise
            warnings.simplefilter("ignore")
            source = CiteProcJSON(sanitized)
            style = CitationStylesStyle(str(self.style_path), locale=self.locale)
            bib = CitationStylesBibliography(style, source, plain)

            inlines = []
            for cluster in clusters:
                cites = []
                for it in cluster:
                    k = keys.get(id(it))
                    if k is None:
                        raise KeyError(
                            "citation item missing from bibliography set")
                    cites.append(CitationItem(k))
                citation = Citation(cites)
                bib.register(citation)
                inlines.append(str(bib.cite(citation, None)))
            # citeproc-py only lists items that were registered in a citation
            # in its bibliography output; kept-uncited items (--keep-uncited)
            # need pseudo-citations, whose inline text we discard.  Resulting
            # numbering/order for those items is UNVERIFIED against Zotero.
            cited_keys = {keys[id(it)] for cluster in clusters for it in cluster}
            extras = [it for it in bib_items if keys[id(it)] not in cited_keys]
            for it in extras:
                extra = Citation([CitationItem(keys[id(it)])])
                bib.register(extra)
                bib.cite(extra, None)
            if extras:
                self.notes.append(
                    f"{len(extras)} uncited item(s) kept in the bibliography "
                    "were previewed via pseudo-citations (numbering/order "
                    "UNVERIFIED against Zotero)"
                )
            bib.sort()
            entries = [str(e) for e in bib.bibliography()]
        return inlines, entries
