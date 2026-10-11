"""
zoteroconv.detect_citations -- find numeric in-text citations with exact
(paragraph, start, end) spans.

Scans word/document.xml, word/footnotes.xml and word/endnotes.xml, but in the
document part ONLY the body before the reference-list heading (never the
reference list itself, never anything after it).

Patterns (Phase 1 spec):
  [1]   [1,2]   [1, 3-5]   [1-4]  (hyphen / en dash / em dash)
  (1)   (1,2)   (1;3)
  superscript runs (w:vertAlign superscript) -- several w:t in one run or
  several adjacent runs are handled via a per-character map across runs.

Exclusions / ambiguity rules (spec + user decisions):
  * characters inside an existing Zotero field  -> excluded (decision 1)
  * numbers after Table/Figure/Fig./Eq./Equation/Section/Appendix labels
                                                  -> ambiguous (not converted)
  * matches followed by a measurement unit       -> ambiguous (dosages)
  * numbers in tables unless in bracket form     -> excluded
  * years (4+ digits) never match (patterns are limited to 1-3 digits;
    superscript digit groups of 4+ are skipped)
  * "(2) the second option" enumerations: parenthetical counts only if the
    number is a valid reference number AND the preceding non-space character
    is a word/punctuation char (not paragraph start, not after a colon, not
    followed by " lowercase-word")
  * numbers not present in the reference list    -> ambiguous
Everything doubtful lands in `ambiguous` (reported, NOT converted).
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field as dc_field

from .docx_io import (
    DOCUMENT_PART,
    ENDNOTES_PART,
    FOOTNOTES_PART,
    DocxPackage,
)
from .parse_reference_list import ReferenceList, find_reference_list
from .zotero_fields import NS, ZoteroField, w, zotero_field_zones

SCANNED_PARTS = (DOCUMENT_PART, FOOTNOTES_PART, ENDNOTES_PART)

BRACKET_RE = re.compile(
    r"\[\s*(\d{1,3}(?:\s*[,;]\s*\d{1,3}|\s*[-\u2013\u2014]\s*\d{1,3})*)\s*\]"
)
PAREN_RE = re.compile(
    r"\(\s*(\d{1,3}(?:\s*[,;]\s*\d{1,3}|\s*[-\u2013\u2014]\s*\d{1,3})*)\s*\)"
)
SUP_MATCH_RE = re.compile(
    r"[\(\[]?\d{1,3}(?:\s*[,;]\s*\d{1,3}|\s*[-\u2013\u2014]\s*\d{1,3})*[\)\]]?"
)
UNIT_TOKENS = {
    "m", "cm", "mm", "km", "nm", "µm", "μm", "kg", "g", "mg", "mcg", "µg",
    "μg", "ng", "l", "ml", "dl", "cl", "mol", "mmol", "x", "n",
}
LABEL_BEFORE_RE = re.compile(
    r"(?:Tables?|Figures?|Figs?\.|Eqs?\.|Equations?|Sections?|Sec\.|"
    r"Appendix|Appendices|Charts?|Schemes?|Boxes?|Steps?|Rules?|Items?|"
    r"Levels?|Phases?|Lines?|Columns?|No\.)\s*(?:\d+|[IVX]+)?\s*$",
    re.I,
)
UNIT_AFTER_RE = re.compile(
    r"^\s*(?:mg|mcg|µg|μg|ug|ng|g|kg|ml|dl|cl|l|IU|mmol|µmol|μmol|nmol|mol|"
    r"mmHg|kPa|°C|%|units?|capsules?|tablets?|drops?|times|fold|"
    r"days?|weeks?|months?|years?|hrs?|min|sec)\b",
    re.I,
)


@dataclass
class CharInfo:
    ch: str
    t_elem: object          # the w:t element this char lives in
    t_offset: int           # offset of the char inside t_elem.text
    run: object             # ancestor w:r element
    superscript: bool
    field_zone: bool        # char belongs to an existing Zotero field


@dataclass
class ParagraphText:
    part: str
    par: object             # w:p element
    index: int              # document-order index among all w:p of the part
    text: str               # concatenated w:t text
    chars: list[CharInfo]
    in_table: bool


@dataclass
class CitationMatch:
    part: str
    par_index: int
    paragraph: object
    start: int              # char offset in ParagraphText.text (inclusive)
    end: int                # exclusive
    text: str               # matched raw text, e.g. '[2,5]'
    numbers: list[int]
    kind: str               # 'bracket' | 'paren' | 'superscript'
    status: str             # 'ok' | 'ambiguous' | 'excluded'
    reason: str = ""


@dataclass
class DetectionResult:
    citations: list[CitationMatch] = dc_field(default_factory=list)   # ok only
    ambiguous: list[CitationMatch] = dc_field(default_factory=list)
    excluded: list[CitationMatch] = dc_field(default_factory=list)
    existing_fields: list = dc_field(default_factory=list)  # list[ZoteroField]
    ref_numbers: set[int] | None = None
    notes: list[str] = dc_field(default_factory=list)
    paragraph_texts: dict[str, list[ParagraphText]] = dc_field(default_factory=dict)

    def all_matches(self) -> list[CitationMatch]:
        return self.citations + self.ambiguous + self.excluded


def expand_numbers(inner: str) -> tuple[list[int] | None, str | None]:
    """'2, 5-7;9' -> [2,5,6,7,9]; returns (None, reason) when unparsable."""
    nums: list[int] = []
    for chunk in re.split(r"[,;]", inner):
        chunk = chunk.strip()
        if not chunk:
            continue
        m = re.fullmatch(r"(\d{1,3})\s*[-\u2013\u2014]\s*(\d{1,3})", chunk)
        if m:
            a, b = int(m.group(1)), int(m.group(2))
            if b < a:
                return None, f"descending range {a}-{b}"
            if b - a > 100:
                return None, f"range too wide {a}-{b}"
            nums.extend(range(a, b + 1))
        elif re.fullmatch(r"\d{1,3}", chunk):
            nums.append(int(chunk))
        else:
            return None, f"unparsable component {chunk!r}"
    if not nums:
        return None, "no numbers inside match"
    seen: set[int] = set()
    out: list[int] = []
    for n in nums:
        if n not in seen:
            seen.add(n)
            out.append(n)
    return out, None


def build_paragraph_texts(pkg: DocxPackage) -> tuple[dict[str, list[ParagraphText]], list[ZoteroField]]:
    """
    Character map across runs for every scanned part. Every char knows its
    w:t element + offset, its run, whether the run is superscript, and
    whether it lies inside an existing Zotero field.
    """
    out: dict[str, list[ParagraphText]] = {}
    all_fields: list[ZoteroField] = []
    for part in SCANNED_PARTS:
        root = pkg.tree_or_none(part)
        if root is None:
            continue
        zones, fields = zotero_field_zones(root, part)
        all_fields.extend(fields)
        pts: list[ParagraphText] = []
        for idx, par in enumerate(root.xpath("//w:p", namespaces=NS)):
            chars: list[CharInfo] = []
            for t in par.xpath(".//w:t", namespaces=NS):
                text = t.text or ""
                run_anc = t.xpath("ancestor::w:r[1]", namespaces=NS)
                run = run_anc[0] if run_anc else None
                sup = False
                in_zone = False
                if run is not None:
                    va = run.xpath("./w:rPr/w:vertAlign/@w:val", namespaces=NS)
                    sup = bool(va) and va[0] == "superscript"
                    in_zone = run in zones
                for off, ch in enumerate(text):
                    chars.append(
                        CharInfo(
                            ch=ch,
                            t_elem=t,
                            t_offset=off,
                            run=run,
                            superscript=sup,
                            field_zone=in_zone,
                        )
                    )
            pts.append(
                ParagraphText(
                    part=part,
                    par=par,
                    index=idx,
                    text="".join(c.ch for c in chars),
                    chars=chars,
                    in_table=bool(par.xpath("ancestor::w:tbl[1]", namespaces=NS)),
                )
            )
        out[part] = pts
    return out, all_fields


def _superscript_candidates(pt: ParagraphText) -> list[CitationMatch]:
    cands: list[CitationMatch] = []
    n = len(pt.chars)
    i = 0
    while i < n:
        if not pt.chars[i].superscript:
            i += 1
            continue
        j = i
        while j < n and pt.chars[j].superscript:
            j += 1
        span = pt.text[i:j]
        if re.search(r"\d", span) and not re.fullmatch(r"\s*\d{4,}\s*", span):
            for m in SUP_MATCH_RE.finditer(span):
                start, end = i + m.start(), i + m.end()
                # not part of a longer digit group (years etc.)
                if end < n and pt.text[end].isdigit():
                    continue
                if start > 0 and pt.text[start - 1].isdigit():
                    continue
                nxt = pt.text[end] if end < len(pt.text) else ""
                prev = pt.text[start - 1] if start > 0 else ""
                inner = m.group().strip("()[]")
                single = re.fullmatch(r"\d", inner) is not None
                if nxt.isalpha():
                    reason = "ordinal-like superscript (letter follows digits)"
                elif single and prev.isalpha():
                    pw = re.search(r"([A-Za-zµμ]+)$", pt.text[:start])
                    if pw and pw.group(1).lower() in UNIT_TOKENS:
                        reason = (
                            "single superscript digit after a unit token "
                            f"({pw.group(1)!r}) -- measurement, not a citation"
                        )
                    else:
                        reason = ""  # e.g. "markets3" -- a normal Vancouver citation
                else:
                    reason = ""
                cands.append(
                    CitationMatch(
                        part=pt.part,
                        par_index=pt.index,
                        paragraph=pt.par,
                        start=start,
                        end=end,
                        text=m.group(),
                        numbers=[],
                        kind="superscript",
                        status="ok",
                        reason=reason,
                    )
                )
        i = j if j > i else i + 1
    return cands


def _text_pattern_candidates(pt: ParagraphText) -> list[CitationMatch]:
    cands: list[CitationMatch] = []
    for rx, kind in ((BRACKET_RE, "bracket"), (PAREN_RE, "paren")):
        for m in rx.finditer(pt.text):
            cands.append(
                CitationMatch(
                    part=pt.part,
                    par_index=pt.index,
                    paragraph=pt.par,
                    start=m.start(),
                    end=m.end(),
                    text=m.group(),
                    numbers=[],
                    kind=kind,
                    status="ok",
                    reason="",
                )
            )
    return cands


def _classify(
    cand: CitationMatch,
    pt: ParagraphText,
    ref_numbers: set[int] | None,
) -> CitationMatch:
    """Decide ok / ambiguous / excluded for one candidate (mutates + returns)."""
    text = pt.text
    # 1. inside an existing Zotero field -> excluded (decision 1: never touch)
    if any(c.field_zone for c in pt.chars[cand.start : cand.end]):
        cand.status = "excluded"
        cand.reason = "inside an existing Zotero field (never modified)"
        return cand
    # 2. numbers must be parsable
    nums, err = expand_numbers(
        cand.text.strip("()[]") if cand.kind == "superscript" else
        text[cand.start : cand.end].strip("()[]")
    )
    if nums is None:
        cand.status = "ambiguous"
        cand.reason = err or "unparsable numbers"
        return cand
    cand.numbers = nums
    # 3. label before ("Table 2 [2]" etc.)
    before = text[max(0, cand.start - 40) : cand.start]
    if LABEL_BEFORE_RE.search(before):
        cand.status = "ambiguous"
        cand.reason = "follows a Table/Figure/Eq./Section/... label"
        return cand
    # 4. measurement unit after (dosages etc.)
    if UNIT_AFTER_RE.match(text[cand.end : cand.end + 20]):
        cand.status = "ambiguous"
        cand.reason = "followed by a measurement unit (dosage/measurement?)"
        return cand
    # 5. kind-specific rules
    if cand.kind == "superscript" and cand.reason:
        cand.status = "ambiguous"
        return cand
    if cand.kind in ("paren", "superscript") and pt.in_table:
        cand.status = "excluded"
        cand.reason = "number in a table (only bracket form converts there)"
        return cand
    if cand.kind == "paren":
        j = cand.start - 1
        while j >= 0 and text[j] == " ":
            j -= 1
        if j < 0:
            cand.status = "ambiguous"
            cand.reason = "enumeration-like: (n) at paragraph start"
            return cand
        if text[j] == ":":
            cand.status = "ambiguous"
            cand.reason = "enumeration-like: (n) directly after a colon"
            return cand
        if not (text[j].isalnum() or text[j] in ".,;:!?)]}\"'’–—%/"):
            cand.status = "ambiguous"
            cand.reason = f"preceded by {text[j]!r} (not a word/punctuation char)"
            return cand
        after = text[cand.end : cand.end + 2]
        preceded_by_space = cand.start > 0 and text[cand.start - 1] == " "
        if preceded_by_space and len(after) == 2 and after[0] == " " and after[1].islower():
            cand.status = "ambiguous"
            cand.reason = "enumeration-like: (n) followed by ' lowercase-word'"
            return cand
    # 6. validity against the reference list
    if ref_numbers is None:
        cand.status = "ambiguous"
        cand.reason = "no reference list found; cannot validate numbers"
        return cand
    invalid = [n for n in cand.numbers if n not in ref_numbers]
    if invalid:
        cand.status = "ambiguous"
        cand.reason = f"number(s) {invalid} not present in the reference list"
        return cand
    cand.status = "ok"
    cand.reason = ""
    return cand


def detect_citations(
    pkg: DocxPackage, ref_list: ReferenceList | None = None
) -> DetectionResult:
    """
    Detect all numeric in-text citations in the scanned parts.
    If ref_list is None it is discovered via find_reference_list(pkg).
    """
    if ref_list is None:
        ref_list = find_reference_list(pkg)
    notes: list[str] = []
    ref_numbers: set[int] | None = None
    if ref_list is not None and ref_list.references:
        ref_numbers = ref_list.numbers()
    elif ref_list is not None:
        notes.append("reference heading found but no entries parsed")
    else:
        notes.append("no reference list found in the document")

    paragraph_texts, existing_fields = build_paragraph_texts(pkg)
    notes.append(f"existing Zotero fields found: {len(existing_fields)} (skipped)")

    region = set(ref_list.region_elements) if ref_list is not None else set()
    heading = ref_list.heading if ref_list is not None else None
    if heading is not None:
        region.add(heading)

    result = DetectionResult(
        existing_fields=existing_fields,
        ref_numbers=ref_numbers,
        notes=notes,
        paragraph_texts=paragraph_texts,
    )

    part_order = {p: i for i, p in enumerate(SCANNED_PARTS)}
    for part, pts in paragraph_texts.items():
        heading_index = None
        if part == DOCUMENT_PART and heading is not None:
            for pt in pts:
                if pt.par is heading:
                    heading_index = pt.index
                    break
        for pt in pts:
            if pt.par in region:
                continue
            if (
                part == DOCUMENT_PART
                and heading_index is not None
                and pt.index > heading_index
            ):
                continue  # never scan the reference list or anything after it
            cands = _superscript_candidates(pt) + _text_pattern_candidates(pt)
            if not cands:
                continue
            # priority: superscript > bracket > paren; drop overlapping losers
            prio = {"superscript": 0, "bracket": 1, "paren": 2}
            cands.sort(key=lambda c: (c.start, -c.end, prio[c.kind]))
            kept: list[CitationMatch] = []
            for c in cands:
                clash = next(
                    (
                        k
                        for k in kept
                        if c.start < k.end and k.start < c.end
                    ),
                    None,
                )
                if clash is not None:
                    c.status = "excluded"
                    c.reason = f"overlaps a {clash.kind} match at the same position"
                    result.excluded.append(c)
                    continue
                kept.append(c)
            for c in kept:
                _classify(c, pt, ref_numbers)
                if c.status == "ok":
                    result.citations.append(c)
                elif c.status == "ambiguous":
                    result.ambiguous.append(c)
                else:
                    result.excluded.append(c)

    result.citations.sort(key=lambda c: (part_order.get(c.part, 99), c.par_index, c.start))
    result.ambiguous.sort(key=lambda c: (part_order.get(c.part, 99), c.par_index, c.start))
    result.excluded.sort(key=lambda c: (part_order.get(c.part, 99), c.par_index, c.start))
    return result
