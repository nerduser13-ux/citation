"""
zoteroconv.parse_reference_list -- locate and parse the numbered reference list.

Capabilities (Phase 1 spec):
  * find the list by the LAST heading matching
    References | Bibliography | Works Cited | Reference List,
    stopping at the next heading of equal-or-higher level (appendices are
    not swallowed), at 'Appendix/Acknowledgements/...' text, at an existing
    ZOTERO_BIBL field paragraph, or at the first non-entry paragraph;
  * typed numbers "1." / "[1]" / "(1)" / "1)", Word auto-numbering via
    w:numPr resolved from numbering.xml (real numbering, NOT paragraph
    counting), references inside a table (row = entry), entries wrapped
    across paragraphs (continuation merging);
  * parse each entry into CSL-JSON: type inference (article-journal, book,
    chapter, webpage, report, thesis, paper-conference), authors (persons
    with initials, 'et al', corporate names), title, container-title,
    issued, volume, issue, page, DOI, PMID, ISBN, URL, accessed;
    DOI/PMID/ISBN are extracted by regex FIRST; confidence high/medium/low;
    the raw text is always kept; the title field is NEVER the whole raw
    string.
  * map existing ZOTERO_ITEM fields' itemData back to typed reference
    numbers (DOI first, then normalised title, then fuzzy title >= 0.90);
    ambiguous or missing matches are reported, never auto-picked
    (user decision 1 for Phase 1).
"""

from __future__ import annotations

import re
import unicodedata
from dataclasses import dataclass, field as dc_field
from difflib import SequenceMatcher

from .docx_io import DOCUMENT_PART, DocxPackage
from .zotero_fields import NS, ZoteroField, field_runs, find_zotero_fields, w

# ---------------------------------------------------------------------------
# data model
# ---------------------------------------------------------------------------

@dataclass
class Reference:
    number: int
    raw: str                                  # full raw text (joined paragraphs)
    csl: dict = dc_field(default_factory=dict)  # CSL-JSON itemData (no "id")
    confidence: str = "low"                   # 'high' | 'medium' | 'low'
    issues: list[str] = dc_field(default_factory=list)
    elements: list = dc_field(default_factory=list)  # w:p / w:tr elements
    part: str = DOCUMENT_PART
    number_source: str = "typed"              # 'typed' | 'auto' | 'table'


@dataclass
class ReferenceList:
    part: str
    heading: object | None                    # w:p element of the heading
    heading_level: int | None                 # 1..9 via style, else None
    references: list[Reference]
    region_elements: list                     # body-level elements belonging to the list
    missing_numbers: list[int] = dc_field(default_factory=list)
    stopped_at: object | None = None
    notes: list[str] = dc_field(default_factory=list)

    def numbers(self) -> set[int]:
        return {r.number for r in self.references}


# ---------------------------------------------------------------------------
# helpers
# ---------------------------------------------------------------------------

def par_text(par) -> str:
    """Visible text of a paragraph (w:t only; ignores instrText)."""
    return "".join(t.text or "" for t in par.xpath(".//w:t", namespaces=NS))


def cell_text(tc) -> str:
    return " ".join(
        s for s in (par_text(p).strip() for p in tc.xpath(".//w:p", namespaces=NS)) if s
    )


HEADING_TEXT_RE = re.compile(
    r"^(references|bibliography|works\s+cited|reference\s+list)\s*[:.]?$", re.I
)
STOP_TEXT_RE = re.compile(
    r"^(appendix|acknowledgement|acknowledgments?|acknowledgements|"
    r"glossary|index|supplementary\s+material|supplements?)\b",
    re.I,
)
TYPED_NUM_RE = re.compile(
    r"^\s*(?:\[(\d{1,3})\]|\((\d{1,3})\)|(\d{1,3})\s*[.)\]]|(\d{1,3})(?=\s+[A-ZÀ-ÖØ-Þ]))"
)


def typed_number(text: str) -> int | None:
    m = TYPED_NUM_RE.match(text)
    if not m:
        return None
    for g in m.groups():
        if g is not None:
            return int(g)
    return None


def strip_number_prefix(text: str) -> str:
    m = TYPED_NUM_RE.match(text)
    if not m:
        return text.strip()
    return text[m.end() :].lstrip(" \t")


class StyleIndex:
    """styleId -> heading level (1..9; Title = 0), from word/styles.xml."""

    def __init__(self, styles_root):
        self.level_of: dict[str, int] = {}
        if styles_root is None:
            return
        for st in styles_root.xpath("//w:style", namespaces=NS):
            sid = st.get(w("styleId"))
            if not sid:
                continue
            name = (st.xpath("w:name/@w:val", namespaces=NS) or [""])[0]
            m = re.match(r"heading\s+(\d)$", name, re.I) or re.fullmatch(
                r"Heading(\d)", sid
            )
            if m:
                self.level_of[sid] = int(m.group(1))
            elif name.lower() == "title" or sid == "Title":
                self.level_of[sid] = 0


def paragraph_heading_level(par, style_index: StyleIndex) -> int | None:
    vals = par.xpath("./w:pPr/w:pStyle/@w:val", namespaces=NS)
    if not vals:
        return None
    return style_index.level_of.get(vals[0])


# ---------------------------------------------------------------------------
# numbering.xml resolution (real numbers, not paragraph counting)
# ---------------------------------------------------------------------------

class NumberingResolver:
    """Resolve w:numPr paragraphs to their actual list numbers."""

    def __init__(self, numbering_root):
        self.abstract: dict[int, dict[int, dict]] = {}
        self.num_to_abs: dict[int, int] = {}
        self.start_override: dict[int, dict[int, int]] = {}
        if numbering_root is None:
            return
        for an in numbering_root.xpath("//w:abstractNum", namespaces=NS):
            aid_s = an.get(w("abstractNumId"))
            if aid_s is None or not aid_s.lstrip("-").isdigit():
                continue
            aid = int(aid_s)
            lvls: dict[int, dict] = {}
            for lvl in an.xpath("w:lvl", namespaces=NS):
                ilvl = int(lvl.get(w("ilvl"), "0"))
                fmt = (lvl.xpath("w:numFmt/@w:val", namespaces=NS) or ["decimal"])[0]
                start = int((lvl.xpath("w:start/@w:val", namespaces=NS) or ["1"])[0])
                lvls[ilvl] = {"fmt": fmt, "start": start}
            self.abstract[aid] = lvls
        for num in numbering_root.xpath("//w:num", namespaces=NS):
            nid = int(num.get(w("numId"), "0"))
            absref = num.xpath("w:abstractNumId/@w:val", namespaces=NS)
            if absref:
                self.num_to_abs[nid] = int(absref[0])
            ovr: dict[int, int] = {}
            for lo in num.xpath("w:lvlOverride", namespaces=NS):
                ilvl = int(lo.get(w("ilvl"), "0"))
                so = lo.xpath("w:startOverride/@w:val", namespaces=NS)
                if so:
                    ovr[ilvl] = int(so[0])
            if ovr:
                self.start_override[nid] = ovr

    def resolve_document(self, doc_root) -> dict:
        """
        w:p element -> number for every w:numPr paragraph in document order.

        Keyed by live element references (NOT id()): lxml proxies are only
        identity-stable while referenced; the returned dict keeps them alive.
        """
        counters: dict[tuple[int, int], int] = {}
        seen_numids: set[int] = set()
        out: dict = {}
        for par in doc_root.iter(w("p")):
            numpr = par.xpath("./w:pPr/w:numPr", namespaces=NS)
            if not numpr:
                continue
            nid_vals = numpr[0].xpath("w:numId/@w:val", namespaces=NS)
            if not nid_vals:
                continue
            nid = int(nid_vals[0])
            ilvl_vals = numpr[0].xpath("w:ilvl/@w:val", namespaces=NS)
            ilvl = int(ilvl_vals[0]) if ilvl_vals else 0
            if nid == 0:  # numId 0 means "numbering removed"
                out[par] = None
                continue
            absid = self.num_to_abs.get(nid)
            lvl = self.abstract.get(absid, {}).get(ilvl) if absid is not None else None
            if lvl is None:
                out[par] = None
                continue
            if nid not in seen_numids:
                seen_numids.add(nid)
                ovr = self.start_override.get(nid, {}).get(ilvl)
                if ovr is not None:
                    counters[(absid, ilvl)] = ovr - 1
            if lvl["fmt"] == "bullet" or lvl["fmt"] == "none":
                out[par] = None
                continue
            key = (absid, ilvl)
            if key not in counters:
                counters[key] = lvl["start"] - 1
            counters[key] += 1
            for (a, l) in list(counters):
                if a == absid and l > ilvl:
                    counters[(a, l)] = self.abstract[a].get(l, {"start": 1})["start"] - 1
            out[par] = counters[key]
        return out


# ---------------------------------------------------------------------------
# region scanning
# ---------------------------------------------------------------------------

CONTINUATION_START_RE = re.compile(
    r"^(available\s+(from|at)|https?://|doi\s*:|\[cited|\[internet|accessed|www\.)",
    re.I,
)


def looks_like_continuation(prev_raw: str, text: str) -> bool:
    """True when `text` (unnumbered paragraph) continues the previous entry."""
    t = text.strip()
    if not t:
        return False
    if CONTINUATION_START_RE.match(t):
        return True
    if t[0].islower():
        return True
    prev = prev_raw.rstrip()
    if prev and prev[-1] in "-–—,;:(":
        return True
    return False


def find_reference_list(pkg: DocxPackage) -> ReferenceList | None:
    """
    Locate + parse the reference list in word/document.xml.
    Returns None when no References/Bibliography/... heading exists.
    """
    root = pkg.document
    style_index = StyleIndex(pkg.styles)
    resolver = NumberingResolver(pkg.numbering)
    numbers_by_par = resolver.resolve_document(root)

    # paragraphs occupied by an existing ZOTERO_BIBL field (decision 2):
    # the typed-list scan must stop there; Phase 3 replaces both.
    # Uses live element references (lxml proxy identity, see zotero_fields).
    bibl_zone: set = set()
    for f in find_zotero_fields(root, DOCUMENT_PART):
        if f.kind == "BIBL":
            runs = field_runs(root, f)
            for par in root.iter(w("p")):
                if any(r in runs for r in par.iter(w("r"))):
                    bibl_zone.add(par)

    all_pars = list(root.iter(w("p")))
    candidates = [
        (i, p)
        for i, p in enumerate(all_pars)
        if HEADING_TEXT_RE.match(par_text(p).strip())
    ]
    if not candidates:
        return None
    head_idx, heading = candidates[-1]  # LAST heading wins (spec)
    heading_level = paragraph_heading_level(heading, style_index)

    notes: list[str] = []
    if len(candidates) > 1:
        notes.append(
            f"{len(candidates)} heading candidates found; used the last one "
            f"(paragraph #{head_idx})"
        )

    body = heading.getparent()
    if body is not None and body.tag == w("body"):
        children = list(body)
        siblings = children[children.index(heading) + 1 :]
    else:
        notes.append("heading is not a direct child of w:body; flat paragraph walk used")
        siblings = all_pars[head_idx + 1 :]

    references: list[Reference] = []
    region: list = []
    pending_empty: list = []
    missing: list[int] = []
    expected = 1
    current: Reference | None = None
    stopped_at = None

    def flush_pending() -> None:
        nonlocal pending_empty
        region.extend(pending_empty)
        pending_empty = []

    def add_entry(number: int, raw: str, elements: list, source: str) -> Reference:
        nonlocal current, expected
        flush_pending()
        ref = Reference(
            number=number,
            raw=raw,
            elements=elements,
            number_source=source,
            part=DOCUMENT_PART,
        )
        references.append(ref)
        current = ref
        expected = number + 1
        return ref

    def try_entry_text(text: str, elements: list, source: str, in_table: bool) -> str:
        """Process one paragraph/row text. Returns action taken."""
        nonlocal current, stopped_at
        if in_table:
            tn = typed_number(text)
            if tn is not None:
                if references and tn < expected:
                    return "stop"
                if tn > expected:
                    missing.extend(range(expected, tn))
                    notes.append(
                        f"gap in table list: numbers {expected}..{tn - 1} missing"
                    )
                add_entry(tn, strip_number_prefix(text), elements, "table")
                return "entry"
            if not text.strip():
                return "empty"
            if current is not None and looks_like_continuation(current.raw, text):
                flush_pending()
                current.raw += " " + text.strip()
                current.elements.extend(elements)
                return "continuation"
            return "stop"
        # normal paragraph
        tn = typed_number(text)
        if tn is not None:
            if references and tn < expected:
                return "stop"
            if tn > expected:
                missing.extend(range(expected, tn))
                notes.append(f"gap: numbers {expected}..{tn - 1} missing from list")
            add_entry(tn, strip_number_prefix(text), elements, "typed")
            return "entry"
        auto = numbers_by_par.get(elements[0]) if elements else None
        if auto is not None and elements and elements[0].tag == w("p"):
            if references and auto < expected:
                return "stop"
            if auto > expected:
                missing.extend(range(expected, auto))
                notes.append(f"gap: numbers {expected}..{auto - 1} missing from list")
            add_entry(auto, text.strip(), elements, "auto")
            return "entry"
        if not text.strip():
            pending_empty.append(elements[0])
            return "empty"
        if current is not None and looks_like_continuation(current.raw, text):
            flush_pending()
            current.raw += " " + text.strip()
            current.elements.extend(elements)
            return "continuation"
        return "stop"

    for el in siblings:
        if el.tag == w("sectPr"):
            break
        if el.tag == w("tbl"):
            rows = el.xpath("./w:tr", namespaces=NS)
            consumed_any = False
            stop_table = False
            for tr in rows:
                row_text = " ".join(
                    s
                    for s in (
                        cell_text(tc) for tc in tr.xpath("./w:tc", namespaces=NS)
                    )
                    if s
                )
                action = try_entry_text(row_text, [tr], "table", in_table=True)
                if action == "entry" or action == "continuation":
                    consumed_any = True
                elif action == "stop":
                    stop_table = True
                    break
            if stop_table or not consumed_any:
                stopped_at = el
                notes.append("stopped at table without further recognizable entries")
                break
            flush_pending()
            region.append(el)
            continue
        if el.tag != w("p"):
            stopped_at = el
            notes.append(f"stopped at non-paragraph element <{el.tag}>")
            break
        if el in bibl_zone:
            stopped_at = el
            notes.append("stopped at paragraph of an existing ZOTERO_BIBL field")
            break
        text = par_text(el)
        lvl = paragraph_heading_level(el, style_index)
        if lvl is not None and (heading_level is None or lvl <= heading_level):
            stopped_at = el
            notes.append(f"stopped at heading (level {lvl}): {text.strip()[:60]!r}")
            break
        if STOP_TEXT_RE.match(text.strip()):
            stopped_at = el
            notes.append(f"stopped at section text: {text.strip()[:60]!r}")
            break
        action = try_entry_text(text, [el], "typed", in_table=False)
        if action == "entry" or action == "continuation":
            region.append(el)
        elif action == "stop":
            stopped_at = el
            notes.append(
                f"stopped at non-entry paragraph: {text.strip()[:60]!r}"
            )
            break
        # 'empty' -> stays pending; appended to region only if more entries follow

    if not references:
        notes.append("heading found but no numbered entries recognized")

    # ---- parse every entry into CSL-JSON ----------------------------------
    for ref in references:
        csl, confidence, issues = parse_reference_text(ref.raw)
        ref.csl = csl
        ref.confidence = confidence
        ref.issues.extend(issues)

    return ReferenceList(
        part=DOCUMENT_PART,
        heading=heading,
        heading_level=heading_level,
        references=references,
        region_elements=region,
        missing_numbers=missing,
        stopped_at=stopped_at,
        notes=notes,
    )


# ---------------------------------------------------------------------------
# CSL-JSON parsing of one entry
# ---------------------------------------------------------------------------

MONTHS: dict[str, int] = {}
for _i, _m in enumerate(
    [
        "january", "february", "march", "april", "may", "june",
        "july", "august", "september", "october", "november", "december",
    ]
):
    MONTHS[_m] = _i + 1
for _abbr, _full in [
    ("jan", "january"), ("feb", "february"), ("mar", "march"),
    ("apr", "april"), ("jun", "june"), ("jul", "july"),
    ("aug", "august"), ("sep", "september"), ("sept", "september"),
    ("oct", "october"), ("nov", "november"), ("dec", "december"),
]:
    MONTHS[_abbr] = MONTHS[_full]

DOI_URL_RE = re.compile(r"https?://(?:dx\.)?doi\.org/(10\.\d{4,9}/[^\s\]\)>]+)", re.I)
DOI_BARE_RE = re.compile(r"\b(10\.\d{4,9}/[^\s\]\)>]+)")
PMID_RE = re.compile(r"\b(?:PMID|PubMed(?:\s+ID)?|MEDLINE)\s*:?\s*(\d{5,9})\b", re.I)
ISBN_RE = re.compile(
    r"\bISBN(?:-1[03])?\s*:?\s*((?:97[89][\s\-]?)?(?:\d[\s\-]?){9}[\dXx])\b", re.I
)
URL_RE = re.compile(r"(?:https?://|www\.)[^\s\]\)>,;\"']+")
ACCESSED_RE = re.compile(
    r"\[\s*(?:cited|accessed)\s+([^\]]+)\]|\(\s*accessed\s+([^)]+)\)"
    r"|\baccessed\s+((?:[A-Z][a-z]+\.?\s+\d{1,2},?\s+)?(?:19|20)\d{2})"
)

# sentence splitter: '.' after lowercase word / ')' / ']' / quote, or after
# an initials group (1-4 capitals preceded by space/comma, e.g. "C." or
# "JM." -- but not inside dotted abbreviations like "U.S."); plus a
# digit-period directly before a year ("100-1. 2023.").
SPLIT_RE = re.compile(
    r"(?:[a-z]\.|\)\.|\]\.|[\"”’]\.|(?<=[\s,])[A-Z]{1,4}\.(?!\s*[A-Z]{1,2}\.)"
    r"|(?<=\d)\.(?=\s*(?:19|20)\d{2}\b))"
    r"\s+(?=[A-ZÀ-ÖØ-Þ0-9(\[\{\"“])"
)

JOURNAL_RE = re.compile(
    r"^(?P<container>[^.]{2,160}?)\.\s*(?P<year>(?:19|20)\d{2})\s*;\s*"
    r"(?P<vol>\d{1,4})\s*(?:\((?P<issue>\d{1,4})\))?\s*"
    r"(?::\s*(?P<page>[A-Za-z]?\d{1,6}(?:\s*[-–—]\s*[A-Za-z]?\d{1,6})?))?\s*\.?"
)
JOURNAL_LOOSE_RE = re.compile(
    r"^(?P<container>[A-ZÀ-ÖØ-Þ][^.]{2,160}?)\.\s*(?P<year>(?:19|20)\d{2})\s*\.?$"
)
CITY_PUB_RE = re.compile(
    r"^(?P<city>[A-ZÀ-ÖØ-Þ][\w .\-]{1,60}?)\s*:\s*(?P<publisher>[^:;.]{2,90}?)"
    r"\s*;\s*(?P<year>(?:19|20)\d{2})"
)
REPORT_SEG_RE = re.compile(
    r"(?:\b(?:Working\s+Papers?|Technical\s+Reports?|Occasional\s+Papers?|"
    r"Discussion\s+Papers?|Research\s+Notes?|Staff\s+Notes?|Report\s+No|"
    r"No\.\s*\d+)\b)|(?:^[A-ZÀ-ÖØ-Þ]{2,}(?:\s+[A-Za-z0-9][\w\-]*)*\s+\d[\w\-]*\s*$)",
    re.I,
)
PAGE_ANYWHERE_RE = re.compile(r"\bpp?\.?\s*(\d{1,5}(?:\s*[-–—]\s*\d{1,5})?)\b")
IN_SEG_RE = re.compile(r"^In\s*[:.]?\s+\S", re.I)
THESIS_RE = re.compile(
    r"\b(Ph\.?\s?D\.?|Doctoral|Master(?:'s|\u2019s)?|M\.?Sc\.?|B\.?Sc\.?|"
    r"dissertation|thesis)\b",
    re.I,
)
UNIVERSITY_RE = re.compile(
    r"\b(University|Universit[éeä]|Univ\.|Institute\s+of\s+Technology|"
    r"Polytechnic|School\s+of|Faculty\s+of|College)\b"
)

PERSON_NAME_RE = re.compile(
    r"^[^\s,;:][^\s,;:]*(?:\s+[A-ZÀ-ÖØ-Þ][\w'’\-]*)*\s+[A-ZÀ-ÖØ-Þ]{1,4}$"
)
ORG_CONNECTIVES = {
    "and", "of", "for", "the", "at", "on", "de", "la", "le", "del",
    "von", "van", "du", "des", "&",
}


def _trim_punct(s: str) -> str:
    return s.rstrip(".,;:)]>}\"'”’").strip()


def parse_dateish(s: str) -> list | None:
    """Parse a date-ish string into CSL date-parts [y] / [y,m] / [y,m,d]."""
    s = s.strip().rstrip(".")
    m = re.fullmatch(r"(\d{1,2})\s+([A-Za-z]+)\s+((?:19|20)\d{2})", s)
    if m:
        mon = MONTHS.get(m.group(2).lower()) or MONTHS.get(m.group(2).lower()[:4])
        return [m.group(3), mon, int(m.group(1))] if mon else [m.group(3)]
    m = re.fullmatch(r"((?:19|20)\d{2})\s+([A-Za-z]+)\s+(\d{1,2})", s)
    if m:  # Vancouver "[cited 2024 Mar 3]"
        mon = MONTHS.get(m.group(2).lower()) or MONTHS.get(m.group(2).lower()[:4])
        return [m.group(1), mon, int(m.group(3))] if mon else [m.group(1)]
    m = re.fullmatch(r"([A-Za-z]+)\s+(\d{1,2}),?\s+((?:19|20)\d{2})", s)
    if m:
        mon = MONTHS.get(m.group(1).lower()) or MONTHS.get(m.group(1).lower()[:4])
        return [m.group(3), mon, int(m.group(2))] if mon else [m.group(3)]
    m = re.fullmatch(r"([A-Za-z]+)\s+((?:19|20)\d{2})", s)
    if m:
        mon = MONTHS.get(m.group(1).lower()) or MONTHS.get(m.group(1).lower()[:4])
        return [m.group(2), mon] if mon else [m.group(2)]
    m = re.fullmatch(r"((?:19|20)\d{2})", s)
    if m:
        return [m.group(1)]
    m = re.fullmatch(r"((?:19|20)\d{2})[/-](\d{1,2})(?:[/-](\d{1,2}))?", s)
    if m:
        parts = [m.group(1), int(m.group(2))]
        if m.group(3):
            parts.append(int(m.group(3)))
        return parts
    return None


def looks_institutional(seg: str) -> bool:
    s = seg.strip().rstrip(".")
    if not s or len(s) > 90 or any(c in s for c in ":;"):
        return False
    words = s.split()
    if not words:
        return False
    if any(ch.isdigit() for ch in s):
        return False
    if not (words[0][:1].isupper()):
        return False
    alpha_words = [x for x in words if any(c.isalpha() for c in x)]
    for x in alpha_words:
        core = re.sub(r"[^A-Za-zÀ-ÿ]", "", x)
        if not core:
            continue
        if core[0].isupper():
            continue
        if x.lower() in ORG_CONNECTIVES:
            continue
        return False
    allcaps = [x for x in alpha_words if len(x) > 1 and x.isupper()]
    if len(allcaps) * 2 > len(alpha_words):
        return False
    return True


def parse_author_segment(seg: str) -> tuple[list | None, str, bool]:
    """
    Decide whether the first segment is an author list.
    Returns (names, kind, etal) with kind in {'person','org','none'}.
    """
    s = seg.strip()
    if not s:
        return None, "none", False
    s = s.rstrip(".")
    etal = bool(re.search(r"\bet\s+al\.?$", s, re.I))
    s2 = re.sub(r",?\s*et\s+al\.?$", "", s, flags=re.I).strip().rstrip(",;").strip()
    if not s2:
        return None, "none", etal
    parts = [p.strip() for p in s2.split(",") if p.strip()]
    if parts and all(PERSON_NAME_RE.match(p) for p in parts):
        names = []
        for p in parts:
            toks = p.rsplit(None, 1)
            if len(toks) == 2 and re.fullmatch(r"[A-ZÀ-ÖØ-Þ]{1,4}", toks[1]):
                names.append({"family": toks[0], "given": toks[1]})
            else:
                names.append({"family": p})
        return names, "person", etal
    if looks_institutional(s2):
        if re.search(r"\s+and\s+", s2):
            left, right = re.split(r"\s+and\s+", s2, maxsplit=1)
            if (
                len(left.split()) >= 2
                and len(right.split()) >= 2
                and looks_institutional(left)
                and looks_institutional(right)
            ):
                return [{"name": left.strip()}, {"name": right.strip()}], "org", etal
        return [{"name": s2}], "org", etal
    return None, "none", etal


def split_segments(work: str) -> list[str]:
    pts = [m.end() for m in SPLIT_RE.finditer(work)]
    segs: list[str] = []
    prev = 0
    for p in pts:
        segs.append(work[prev:p].strip())
        prev = p
    segs.append(work[prev:].strip())
    return [s for s in segs if s]


def parse_reference_text(raw: str) -> tuple[dict, str, list[str]]:
    """
    Parse one reference entry's raw text into (csl_item, confidence, issues).
    The title field is never the whole raw string.
    """
    issues: list[str] = []
    text = " ".join(raw.split())
    csl: dict = {}

    # ---- identifiers FIRST (spec) ----
    doi = None
    m = DOI_URL_RE.search(text)
    doi_from_url = False
    if m:
        doi = _trim_punct(m.group(1))
        doi_from_url = True
    else:
        m = DOI_BARE_RE.search(text)
        if m:
            doi = _trim_punct(m.group(1))
    if doi:
        csl["DOI"] = doi
    m = PMID_RE.search(text)
    if m:
        csl["PMID"] = m.group(1)
    m = ISBN_RE.search(text)
    if m:
        csl["ISBN"] = re.sub(r"[\s\-]", "", m.group(1))
    m = ACCESSED_RE.search(text)
    if m:
        raw_date = next((g for g in m.groups() if g), None)
        parts = parse_dateish(raw_date) if raw_date else None
        if parts:
            csl["accessed"] = {"date-parts": [parts]}
        else:
            issues.append(f"accessed date kept unparsed: {raw_date!r}")
    urls = [_trim_punct(u) for u in URL_RE.findall(text)]
    non_doi_urls = [u for u in urls if "doi.org" not in u.lower()]
    if non_doi_urls:
        csl["URL"] = non_doi_urls[0]
        if len(non_doi_urls) > 1:
            issues.append(f"multiple URLs found; kept first: {non_doi_urls[0]!r}")
    elif not doi_from_url and urls:
        csl["URL"] = urls[0]

    # ---- working text: strip identifiers/urls/accessed markers ----
    work = text
    for u in urls:
        work = work.replace(u, " ")
    work = re.sub(r"\[\s*(?:cited|accessed)[^\]]*\]", " ", work, flags=re.I)
    work = re.sub(r"\(\s*accessed[^)]*\)", " ", work, flags=re.I)
    if "PMID" in csl:
        work = PMID_RE.sub(" ", work)
    if "ISBN" in csl:
        work = ISBN_RE.sub(" ", work)
    work, n_pub = re.subn(r"Published\s+online\b[^.]*\.?", " ", work, flags=re.I)
    if n_pub:
        issues.append("source had a 'Published online' date (not used as issued)")
    work = re.sub(r"Available\s+(?:from|at)\s*:?", " ", work, flags=re.I)
    work = " ".join(work.split()).strip()

    segments = split_segments(work)

    # ---- authors ----
    consumed = 0
    names, kind, etal = parse_author_segment(segments[0] if segments else "")
    if names:
        csl["author"] = names
        consumed = 1
        if etal:
            issues.append("author list truncated by 'et al' in the source")
        if kind == "org":
            if len(names) == 2:
                issues.append("corporate author split at 'and' into two creators")
            elif len(names[0].get("name", "").split()) == 1:
                issues.append("single-word segment assumed to be a corporate author")
            else:
                issues.append("corporate/institutional author")
        elif kind == "person":
            issues.append("personal names carry initials only (as printed)")

    # ---- title ----
    title = ""
    if consumed < len(segments):
        title = segments[consumed].rstrip(".").strip()
        title = re.sub(r"\s*\[Internet\]\s*", " ", title, flags=re.I).strip()
        consumed += 1
    if title and title != text.strip():
        csl["title"] = title
    elif title:
        # whole raw text would become the title -> forbidden (spec)
        issues.append("entry is a single segment; title left empty on purpose")
    else:
        issues.append("could not isolate a title segment")

    rest_segments = [s.strip() for s in segments[consumed:] if s.strip()]
    rest = " ".join(rest_segments)

    # ---- type + bibliographic fields ----
    in_seg = next((s for s in rest_segments if IN_SEG_RE.match(s)), None)
    journal_m = JOURNAL_RE.match(rest) if rest else None
    report_seg = next(
        (
            s
            for s in rest_segments
            if REPORT_SEG_RE.search(s.rstrip(".")) and not IN_SEG_RE.match(s)
        ),
        None,
    )
    citypub_m = CITY_PUB_RE.search(rest) if rest else None
    thesis_m = THESIS_RE.search(rest)
    internet_m = re.search(r"\[Internet\]|Available\s+from", text, re.I)

    if in_seg:
        container = re.sub(r"^In\s*[:.]?\s+", "", in_seg, flags=re.I).rstrip(".")
        m_ed = re.search(
            r"[,;]?\s*([A-ZÀ-ÖØ-Þ][^;]*?),?\s*\b(?:editors?|eds?)\b\.?\s*$",
            container,
            re.I,
        )
        if m_ed:
            ed_names, _ed_kind, _ed_etal = parse_author_segment(m_ed.group(1))
            container = container[: m_ed.start()].strip().rstrip(",;")
            if ed_names:
                csl["editor"] = ed_names
        csl["container-title"] = container.strip()
        csl["type"] = (
            "paper-conference"
            if re.search(
                r"proceedings|conference|congress|symposium|workshop|meeting|abstracts",
                container,
                re.I,
            )
            else "chapter"
        )
        pm = PAGE_ANYWHERE_RE.search(rest)
        if pm:
            csl["page"] = pm.group(1).replace("–", "-").replace("—", "-")
        cp = CITY_PUB_RE.search(rest)
        if cp:
            csl["publisher-place"] = cp.group("city").strip()
            csl["publisher"] = cp.group("publisher").strip()
            csl["issued"] = {"date-parts": [[cp.group("year")]]}
    elif thesis_m:
        csl["type"] = "thesis"
        genre_seg = next(
            (s for s in rest_segments if THESIS_RE.search(s)), thesis_m.group(0)
        )
        csl["genre"] = genre_seg.rstrip(".")
        uni_seg = next((s for s in rest_segments if UNIVERSITY_RE.search(s)), None)
        if citypub_m:
            csl["publisher-place"] = citypub_m.group("city").strip()
            csl["publisher"] = citypub_m.group("publisher").strip()
            csl["issued"] = {"date-parts": [[citypub_m.group("year")]]}
        elif uni_seg:
            csl["publisher"] = uni_seg.rstrip(".")
    elif journal_m:
        csl["type"] = "article-journal"
        csl["container-title"] = journal_m.group("container").strip().rstrip(".")
        csl["issued"] = {"date-parts": [[journal_m.group("year")]]}
        csl["volume"] = journal_m.group("vol")
        if journal_m.group("issue"):
            csl["issue"] = journal_m.group("issue")
        if journal_m.group("page"):
            csl["page"] = re.sub(r"\s+", "", journal_m.group("page")).replace(
                "–", "-"
            ).replace("—", "-")
        rest = rest[journal_m.end() :].strip()
        rest_segments = split_segments(rest) if rest else []
        if rest:
            issues.append(f"unparsed tail after journal pattern: {rest[:80]!r}")
    elif report_seg:
        csl["type"] = "report"
        csl["number"] = report_seg.rstrip(".")
    elif internet_m:
        # "[Internet]" / "Available from:" are decisive Vancouver webpage
        # markers even when a "City: Publisher; Year" segment is present.
        csl["type"] = "webpage"
        if citypub_m:
            csl["publisher-place"] = citypub_m.group("city").strip()
            csl["publisher"] = citypub_m.group("publisher").strip()
            csl["issued"] = {"date-parts": [[citypub_m.group("year")]]}
    elif citypub_m:
        csl["type"] = "book"
        csl["publisher-place"] = citypub_m.group("city").strip()
        csl["publisher"] = citypub_m.group("publisher").strip()
        csl["issued"] = {"date-parts": [[citypub_m.group("year")]]}
    else:
        loose_m = JOURNAL_LOOSE_RE.match(rest) if rest else None
        if loose_m and not REPORT_SEG_RE.search(loose_m.group("container")):
            csl["type"] = "article-journal"
            csl["container-title"] = loose_m.group("container").strip().rstrip(".")
            csl["issued"] = {"date-parts": [[loose_m.group("year")]]}
            issues.append("journal detected without volume/pages")
        elif csl.get("URL"):
            csl["type"] = "webpage"
        else:
            csl["type"] = "document"
            issues.append("no type markers recognized; generic 'document' used")

    # ---- issued fallback ----
    if "issued" not in csl:
        for s in rest_segments:
            parts = parse_dateish(s)
            if parts and len(parts) >= 2:
                csl["issued"] = {"date-parts": [parts]}
                break
        if "issued" not in csl:
            for s in rest_segments:
                parts = parse_dateish(s)
                if parts:
                    csl["issued"] = {"date-parts": [parts]}
                    break
        if "issued" not in csl:
            m = re.search(r"(?<![,\w:])\b((?:19|20)\d{2})\b(?![\d(:\w])", rest)
            if m:
                csl["issued"] = {"date-parts": [[m.group(1)]]}
            else:
                issues.append("no publication year found")

    # ---- confidence ----
    core = 0
    if csl.get("author"):
        core += 1
    if csl.get("title"):
        core += 1
    if csl.get("issued"):
        core += 1
    if (
        csl.get("DOI")
        or csl.get("PMID")
        or csl.get("ISBN")
        or (csl.get("container-title") and (csl.get("volume") or csl.get("page")))
        or csl.get("publisher")
    ):
        core += 1
    confidence = "high" if core == 4 else ("medium" if core == 3 else "low")
    if not csl.get("title"):
        confidence = "low"
    if csl.get("type") == "document" and confidence == "high":
        confidence = "medium"
    return csl, confidence, issues


# ---------------------------------------------------------------------------
# mapping existing Zotero fields back to typed reference numbers (decision 1)
# ---------------------------------------------------------------------------

@dataclass
class FieldItemMapping:
    field_index: int
    citation_id: str | None
    item_index: int
    uri: str | None
    item_key: str | None
    doi: str | None
    title: str | None
    ref_number: int | None
    method: str | None            # 'doi' | 'title' | 'title-fuzzy' | None
    candidates: list[int] = dc_field(default_factory=list)
    item_data: dict | None = None  # full CSL-JSON itemData from the field payload


def norm_doi(d: str) -> str:
    d = d.strip().lower()
    d = re.sub(r"^https?://(?:dx\.)?doi\.org/", "", d)
    return d.rstrip(".,;)")


def norm_title(t: str) -> str:
    t = unicodedata.normalize("NFKC", t).lower()
    t = re.sub(r"[^a-z0-9]+", " ", t)
    return " ".join(t.split())


FUZZY_TITLE_THRESHOLD = 0.90


def map_existing_fields_to_references(
    fields: list[ZoteroField], refs: list[Reference]
) -> list[FieldItemMapping]:
    """
    Map every citationItem of every existing ZOTERO_ITEM field to a typed
    reference number: DOI exact match first, then normalised title equality,
    then fuzzy title (SequenceMatcher ratio >= 0.90). Ambiguous results keep
    ref_number=None and list all candidates (never auto-picked).
    """
    out: list[FieldItemMapping] = []
    for fi, f in enumerate(fields):
        if f.kind != "ITEM":
            continue
        for ii, item in enumerate(f.citation_items()):
            data = item.get("itemData") or {}
            uris = item.get("uris") or []
            uri = uris[0] if uris else None
            key = uri.rstrip("/").rsplit("/", 1)[-1] if uri else None
            doi = data.get("DOI")
            title = data.get("title")
            mapping = FieldItemMapping(
                field_index=fi,
                citation_id=f.citation_id,
                item_index=ii,
                uri=uri,
                item_key=key,
                doi=doi,
                title=title,
                ref_number=None,
                method=None,
                item_data=data,
            )
            if doi:
                hits = [
                    r.number
                    for r in refs
                    if r.csl.get("DOI") and norm_doi(r.csl["DOI"]) == norm_doi(doi)
                ]
                if len(hits) == 1:
                    mapping.ref_number = hits[0]
                    mapping.method = "doi"
                    out.append(mapping)
                    continue
                if len(hits) > 1:
                    mapping.candidates = hits
                    out.append(mapping)
                    continue
            if title:
                nt = norm_title(title)
                exact = [
                    r.number
                    for r in refs
                    if r.csl.get("title") and norm_title(r.csl["title"]) == nt
                ]
                if len(exact) == 1:
                    mapping.ref_number = exact[0]
                    mapping.method = "title"
                    out.append(mapping)
                    continue
                if len(exact) > 1:
                    mapping.candidates = exact
                    out.append(mapping)
                    continue
                scored = [
                    (
                        SequenceMatcher(
                            None, nt, norm_title(r.csl["title"])
                        ).ratio(),
                        r.number,
                    )
                    for r in refs
                    if r.csl.get("title")
                ]
                good = sorted(
                    [s for s in scored if s[0] >= FUZZY_TITLE_THRESHOLD], reverse=True
                )
                if len(good) == 1:
                    mapping.ref_number = good[0][1]
                    mapping.method = "title-fuzzy"
                elif len(good) > 1:
                    mapping.candidates = [g[1] for g in good]
            out.append(mapping)
    return out
