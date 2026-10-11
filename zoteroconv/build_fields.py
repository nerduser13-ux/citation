"""
Phase 3 -- builders for live Zotero Word fields, the bibliography field and
the ZOTERO_PREF_* document properties.

Every structural constant below was extracted from
tests/fixtures/zotero_sample.docx (real Zotero 10.0.6 output) and is quoted
with evidence in docs/zotero_format.md (SS = section):

* citation instr:  ' ADDIN ZOTERO_ITEM CSL_CITATION {json} '            (SS3.3)
* payload key order: citationID, properties, citationItems, schema      (SS3.4)
* properties keys:   unsorted, formattedCitation, plainCitation, noteIndex
* citationItem keys: id, uris, itemData (no empty prefix/suffix keys)
* itemData order:    id, type, remaining variables case-insensitive
                     alphabetical, then name vars, then dates last       (SS3.5)
* issued date-parts: year STRING, month/day INT                          (SS3.5)
* field runs: begin/instrText/separate are bare (no rPr, no fldLock,
  no dirty); preview runs carry the surrounding formatting              (SS3.1/3.2/3.6)
* bibl instr: ' ADDIN ZOTERO_BIBL {"uncited":[],"omitted":[],"custom":[]}
  CSL_BIBLIOGRAPHY '                                                    (SS4.1)
* bibl layout: begin/instr/separate + first entry inside the first
  entry paragraph (pStyle Bibliography); end fldChar in a following
  EMPTY paragraph without pPr                                           (SS4.2)
* Bibliography style definition: sample styles.xml verbatim minus the
  w:rsid session artifact                                               (SS4.5)
* prefs: <data data-version="3" zotero-version="..."> raw string, split
  at exactly 255 chars into ZOTERO_PREF_1..N with pids ascending from 2 (SS2.2/2.4)
* custom.xml namespaces, fmtid, content type and relationship type      (SS1.3/2.1/8)
"""

from __future__ import annotations

import json
import random
import re
import string
from copy import deepcopy
from dataclasses import dataclass, field as dc_field
from typing import Any

from lxml import etree

from zoteroconv.docx_io import (
    CONTENT_TYPES_PART,
    CUSTOM_PROPS_CONTENT_TYPE,
    CUSTOM_PROPS_PART,
    CUSTOM_PROPS_REL_TYPE,
    DOCUMENT_PART,
    ROOT_RELS_PART,
    STYLES_PART,
    DocxPackage,
)
from zoteroconv.zotero_fields import NS, W_NS, w

# ---------------------------------------------------------------------------
# exact constants (docs/zotero_format.md SS8)
# ---------------------------------------------------------------------------

ZOTERO_ITEM_INSTR_PREFIX = " ADDIN ZOTERO_ITEM CSL_CITATION "
ZOTERO_BIBL_INSTR_PREFIX = " ADDIN ZOTERO_BIBL "
ZOTERO_BIBL_INSTR_SUFFIX = " CSL_BIBLIOGRAPHY "
SCHEMA_URL = (
    "https://github.com/citation-style-language/schema/raw/master/"
    "csl-citation.json"
)
PREF_FMTID = "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}"
VT_NS = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"
CUSTOM_NS = (
    "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties"
)
PREF_CHUNK_SIZE = 255
PREF_FIRST_PID = 2
PREF_DATA_TMPL = (
    '<data data-version="3" zotero-version="{zv}">'
    '<session id="{sid}"/>'
    '<style id="{style}"{locale_attr} hasBibliography="1" '
    'bibliographyStyleHasBeenSet="1"/>'
    '<prefs><pref name="fieldType" value="Field"/>'
    '<pref name="automaticJournalAbbreviations" value="true"/></prefs></data>'
)
DEFAULT_ZOTERO_VERSION = "10.0.6"

# sample styles.xml verbatim (w:rsid session artifact intentionally omitted)
BIBLIOGRAPHY_STYLE_XML = (
    f'<w:style xmlns:w="{W_NS}" w:type="paragraph" w:styleId="Bibliography">'
    '<w:name w:val="Bibliography"/>'
    '<w:basedOn w:val="Normal"/>'
    '<w:next w:val="Normal"/>'
    '<w:uiPriority w:val="37"/>'
    '<w:unhideWhenUsed/>'
    '<w:pPr>'
    '<w:spacing w:after="0" w:line="240" w:lineRule="auto"/>'
    '<w:ind w:left="720" w:hanging="720"/>'
    '</w:pPr>'
    '</w:style>'
)

# keys observed in the sample's itemData (SS3.5); anything beyond this set is
# reported as an extrapolation by convert.py
SAMPLE_ITEMDATA_KEYS = {
    "id", "type", "abstract", "container-title", "DOI", "ISSN", "issue",
    "journalAbbreviation", "license", "page", "source", "title",
    "title-short", "volume", "author", "issued",
}

ID_ALPHABET = string.ascii_letters + string.digits
NAME_VARS = {"author", "editor", "translator", "collection-editor",
             "composer", "illustrator", "interviewer", "recipient",
             "container-author", "editorial-director", "original-author"}
DATE_VARS = {"issued", "accessed", "event-date", "original-date",
             "submitted", "container"}


# ---------------------------------------------------------------------------
# payload builders (exact key order, compact JSON)
# ---------------------------------------------------------------------------

def new_citation_id(rng: random.Random, taken: set[str]) -> str:
    while True:
        cid = "".join(rng.choice(ID_ALPHABET) for _ in range(8))
        if cid not in taken:
            taken.add(cid)
            return cid


def _norm_date_parts(d: dict) -> dict:
    """year STRING, month/day INT (SS3.5)."""
    out = deepcopy(d)
    parts = (out.get("date-parts") or [])
    fixed = []
    for row in parts:
        new_row = []
        for i, v in enumerate(row):
            if v is None:
                continue
            if i == 0:
                new_row.append(str(v))
            else:
                try:
                    new_row.append(int(v))
                except (TypeError, ValueError):
                    new_row.append(v)
        fixed.append(new_row)
    if fixed:
        out["date-parts"] = fixed
    return out


def ordered_item_data(csl: dict, numeric_id: int) -> dict:
    """itemData with the sample's exact key order: id, type, remaining
    variables case-insensitive alphabetical, then name variables, then date
    variables last (SS3.5; relative placement of multiple name/date vars is
    an extrapolation flagged by convert.py)."""
    out: dict[str, Any] = {"id": numeric_id, "type": csl.get("type", "document")}
    plain = sorted(
        (k for k in csl
         if k not in ("id", "type") and k not in NAME_VARS and k not in DATE_VARS),
        key=lambda k: k.lower(),
    )
    for k in plain:
        out[k] = deepcopy(csl[k])
    for k in sorted((k for k in csl if k in NAME_VARS), key=lambda k: k.lower()):
        out[k] = deepcopy(csl[k])
    for k in sorted((k for k in csl if k in DATE_VARS), key=lambda k: k.lower()):
        v = deepcopy(csl[k])
        if isinstance(v, dict) and "date-parts" in v:
            v = _norm_date_parts(v)
        out[k] = v
    return out


def make_citation_payload(
    citation_id: str,
    entries: list[tuple[int, str, dict]],   # (numeric_id, uri, csl itemData)
    formatted: str,
    note_index: int = 0,
) -> dict:
    """Exact sample key order (SS3.4).  No prefix/suffix keys (SS6.2)."""
    return {
        "citationID": citation_id,
        "properties": {
            "unsorted": False,
            "formattedCitation": formatted,
            "plainCitation": formatted,
            "noteIndex": note_index,
        },
        "citationItems": [
            {"id": nid, "uris": [uri], "itemData": ordered_item_data(csl, nid)}
            for nid, uri, csl in entries
        ],
        "schema": SCHEMA_URL,
    }


def make_bibl_payload(uncited_uris: list[str]) -> dict:
    return {"uncited": list(uncited_uris), "omitted": [], "custom": []}


def dumps_instr_payload(payload: dict) -> str:
    return json.dumps(payload, separators=(",", ":"), ensure_ascii=False)


def citation_instr(payload: dict) -> str:
    return ZOTERO_ITEM_INSTR_PREFIX + dumps_instr_payload(payload) + " "


def bibl_instr(payload: dict) -> str:
    return (ZOTERO_BIBL_INSTR_PREFIX + dumps_instr_payload(payload)
            + ZOTERO_BIBL_INSTR_SUFFIX)


# ---------------------------------------------------------------------------
# run / paragraph builders
# ---------------------------------------------------------------------------

def fld_run(kind: str):
    """Bare fldChar run: no rPr, no fldLock, no dirty (SS3.1/3.2)."""
    r = etree.Element(w("r"))
    fc = etree.SubElement(r, w("fldChar"))
    fc.set(w("fldCharType"), kind)
    return r


def instr_run(text: str):
    r = etree.Element(w("r"))
    it = etree.SubElement(r, w("instrText"))
    it.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    it.text = text
    return r


def text_run(text: str, rpr=None):
    r = etree.Element(w("r"))
    if rpr is not None:
        r.append(deepcopy(rpr))
    t = etree.SubElement(r, w("t"))
    t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    t.text = text
    return r


def rpr_without_superscript(run) -> object | None:
    """rPr template for preview runs: the matched span's own formatting but
    never superscript (user decision a) -- Zotero previews in both fixtures
    carry no rStyle (SS3.6)."""
    if run is None:
        return None
    rpr = run.find("./w:rPr", namespaces=NS)
    if rpr is None:
        return None
    cp = deepcopy(rpr)
    for va in cp.findall("./w:vertAlign", namespaces=NS):
        cp.remove(va)
    if len(cp) == 0:
        return None
    return cp


def citation_field_runs(instr: str, preview: str, rpr=None) -> list:
    """begin / instrText / separate / preview / end -- one paragraph (SS3.1)."""
    runs = [fld_run("begin"), instr_run(instr), fld_run("separate")]
    if preview:
        runs.append(text_run(preview, rpr))
    runs.append(fld_run("end"))
    return runs


def new_par(style_val: str | None = None):
    p = etree.Element(w("p"))
    if style_val is not None:
        ppr = etree.SubElement(p, w("pPr"))
        ps = etree.SubElement(ppr, w("pStyle"))
        ps.set(w("val"), style_val)
    return p


def bibliography_paragraphs(instr: str, entries: list[str]) -> list:
    """SS4.2 single-entry shape generalised to N entries (extrapolation
    SS6.4, flagged): first entry paragraph holds begin/instr/separate +
    entry runs (pStyle Bibliography); each further entry its own paragraph;
    a final EMPTY paragraph without pPr holds the end fldChar."""
    pars = []
    first = new_par("Bibliography")
    first.append(fld_run("begin"))
    first.append(instr_run(instr))
    first.append(fld_run("separate"))
    if entries:
        first.append(text_run(entries[0]))       # no rPr: sample SS4.4
        for e in entries[1:]:
            p = new_par("Bibliography")
            p.append(text_run(e))
            pars.append(p)
    pars.insert(0, first)
    tail = new_par(None)                          # empty paragraph, no pPr
    tail.append(fld_run("end"))
    pars.append(tail)
    return pars


# ---------------------------------------------------------------------------
# run splitting + span replacement (decision a)
# ---------------------------------------------------------------------------

def char_map(par) -> tuple[str, list[dict]]:
    """Standalone char map of one w:p (same traversal order as the detector:
    document-order w:t elements)."""
    chars: list[dict] = []
    text_parts: list[str] = []
    for t in par.xpath(".//w:t", namespaces=NS):
        run = t.xpath("ancestor::w:r[1]", namespaces=NS)
        run = run[0] if run else None
        sup = False
        if run is not None:
            va = run.xpath("./w:rPr/w:vertAlign/@w:val", namespaces=NS)
            sup = bool(va) and va[0] == "superscript"
        s = t.text or ""
        text_parts.append(s)
        for i, ch in enumerate(s):
            chars.append({"ch": ch, "t": t, "off": i, "run": run, "sup": sup})
    return "".join(text_parts), chars


def _split_run_after(run, t, offset: int):
    """Split t.text at offset: keep [0:offset] in place, move [offset:] into
    a new run (rPr deep-copied) inserted directly after `run`."""
    full = t.text or ""
    new_run = etree.Element(w("r"))
    rpr = run.find("./w:rPr", namespaces=NS)
    if rpr is not None:
        new_run.append(deepcopy(rpr))
    new_t = etree.SubElement(new_run, w("t"))
    new_t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    new_t.text = full[offset:]
    t.text = full[:offset]
    parent = run.getparent()
    parent.insert(parent.index(run) + 1, new_run)
    return new_run, new_t


def replace_span_with_runs(par, start: int, end: int, new_runs: list) -> None:
    """Remove exactly the runs covering text[start:end] and insert new_runs
    at that position.  Boundary runs are split so text before/after keeps
    its original run and formatting (decision a)."""
    text, chars = char_map(par)
    if text[start:end] == "":
        raise ValueError("empty span")
    span = chars[start:end]

    # group span chars into (t, run, from, to) segments
    segments: list[list] = []
    for c in span:
        if segments and segments[-1][0] is c["t"]:
            segments[-1][3] = c["off"] + 1
        else:
            segments.append([c["t"], c["run"], c["off"], c["off"] + 1])

    first_run = segments[0][1]
    if first_run is None:
        raise ValueError("span chars have no ancestor run")

    # split leading partial t
    t0, r0, i0, i1 = segments[0]
    if i0 > 0:
        new_run, new_t = _split_run_after(r0, t0, i0)
        segments[0] = [new_t, new_run, 0, i1 - i0]
    # split trailing partial t
    tL, rL, j0, j1 = segments[-1]
    if j1 < len(tL.text or ""):
        _split_run_after(rL, tL, j1)

    parent = first_run.getparent()
    insert_idx = parent.index(
        segments[0][1] if segments[0][1].getparent() is parent else first_run
    )
    removed = []
    for _t, run, _a, _b in segments:
        if run is not None and run not in removed:
            removed.append(run)
    for run in removed:
        rp = run.getparent()
        if rp is not None:
            rp.remove(run)
    for i, nr in enumerate(new_runs):
        parent.insert(insert_idx + i, nr)


# ---------------------------------------------------------------------------
# Bibliography paragraph style (SS4.5)
# ---------------------------------------------------------------------------

def ensure_bibliography_style(pkg: DocxPackage) -> bool:
    """Inject the sample's Bibliography style when missing.  Returns True if
    styles.xml was modified."""
    styles = pkg.styles
    if styles is None:
        return False
    existing = styles.xpath("//w:style[@w:styleId='Bibliography']",
                            namespaces=NS)
    if existing:
        return False
    el = etree.fromstring(BIBLIOGRAPHY_STYLE_XML)
    styles.append(el)
    pkg.mark_modified(STYLES_PART)
    return True


# ---------------------------------------------------------------------------
# ZOTERO_PREF_* properties (SS1.3, SS2)
# ---------------------------------------------------------------------------

@dataclass
class PrefsChange:
    old_style_id: str | None = None
    new_style_id: str | None = None
    old_locale: str | None = None
    session_id: str | None = None
    created_part: bool = False
    chunk_count: int = 0
    pids: list[int] = dc_field(default_factory=list)
    notes: list[str] = dc_field(default_factory=list)

    @property
    def style_changed(self) -> bool:
        return self.old_style_id != self.new_style_id


def split_pref_chunks(data: str) -> list[str]:
    """Raw character cuts at exactly 255 chars (SS2.4)."""
    return [data[i:i + PREF_CHUNK_SIZE]
            for i in range(0, len(data), PREF_CHUNK_SIZE)] or [""]


def read_prefs(pkg: DocxPackage) -> tuple[str | None, str | None, str | None]:
    """(session_id, zotero_version, style_id) from existing prefs, if any."""
    root = pkg.custom_props
    if root is None:
        return None, None, None
    chunks = []
    for p in root.xpath("//*[local-name()='property']"):
        name = p.get("name") or ""
        m = re.fullmatch(r"ZOTERO_PREF_(\d+)", name)
        if m:
            v = p.xpath("*[local-name()='lpwstr']")
            chunks.append((int(m.group(1)), v[0].text if v and v[0].text else ""))
    if not chunks:
        return None, None, None
    chunks.sort()
    joined = "".join(c for _, c in chunks)
    try:
        data = etree.fromstring(joined.encode("utf-8"))
    except etree.XMLSyntaxError:
        return None, None, None
    sid = data.xpath("//session/@id")
    zv = data.get("zotero-version")
    style = data.xpath("//style/@id")
    return (sid[0] if sid else None, zv, style[0] if style else None)


def build_pref_data(style_id: str, locale: str | None, session_id: str,
                    zotero_version: str) -> str:
    locale_attr = f' locale="{locale}"' if locale else ""
    return PREF_DATA_TMPL.format(
        zv=zotero_version, sid=session_id, style=style_id,
        locale_attr=locale_attr,
    )


def write_prefs(pkg: DocxPackage, style_id: str, locale: str | None,
                session_id: str | None = None,
                rng: random.Random | None = None) -> PrefsChange:
    """Create-or-update ZOTERO_PREF_* in docProps/custom.xml: chunked at 255
    chars, names ZOTERO_PREF_1..N, ascending pids from 2, never duplicated;
    part + [Content_Types] Override + root relationship created when missing
    (decision e)."""
    rng = rng or random.Random()
    change = PrefsChange(new_style_id=style_id)
    old_session, old_zv, old_style = read_prefs(pkg)
    change.old_style_id = old_style
    change.session_id = session_id or old_session or "".join(
        rng.choice(ID_ALPHABET) for _ in range(8))
    zv = old_zv or DEFAULT_ZOTERO_VERSION
    if old_zv is None:
        change.notes.append(
            f"zotero-version={zv!r} taken from the fixtures (no existing prefs)"
        )

    root = pkg.custom_props
    if root is None:
        change.created_part = True
        root = etree.Element(
            "{%s}Properties" % CUSTOM_NS, nsmap={None: CUSTOM_NS, "vt": VT_NS}
        )
        pkg.add_part_tree(CUSTOM_PROPS_PART, root)
        # [Content_Types].xml Override
        ct = pkg.content_types
        ct_ns = ct.nsmap.get(None)
        existing = ct.xpath(
            f"/*[local-name()='Types']/*[local-name()='Override']"
            f"[@PartName='/{CUSTOM_PROPS_PART}']")
        if not existing:
            ov = etree.SubElement(ct, f"{{{ct_ns}}}Override")
            ov.set("PartName", "/" + CUSTOM_PROPS_PART)
            ov.set("ContentType", CUSTOM_PROPS_CONTENT_TYPE)
            pkg.mark_modified(CONTENT_TYPES_PART)
        # package relationship
        rels = pkg.root_rels
        rels_ns = rels.nsmap.get(None)
        have = rels.xpath(
            f"/*[local-name()='Relationships']/*[local-name()='Relationship']"
            f"[@Type='{CUSTOM_PROPS_REL_TYPE}']")
        if not have:
            ids = [int(m.group(1))
                   for el in rels
                   for m in [re.fullmatch(r"rId(\d+)", el.get("Id") or "")]
                   if m]
            rel = etree.SubElement(rels, f"{{{rels_ns}}}Relationship")
            rel.set("Id", f"rId{max(ids, default=0) + 1}")
            rel.set("Type", CUSTOM_PROPS_REL_TYPE)
            rel.set("Target", CUSTOM_PROPS_PART)
            pkg.mark_modified(ROOT_RELS_PART)
        change.notes.append(
            "docProps/custom.xml created (part + content-type Override + "
            "package relationship)"
        )
    else:
        # remove existing ZOTERO_PREF_* properties, keep everything else
        for p in list(root.xpath("//*[local-name()='property']")):
            if re.fullmatch(r"ZOTERO_PREF(_\d+)?", p.get("name") or ""):
                p.getparent().remove(p)

    taken_pids = {
        int(p.get("pid"))
        for p in root.xpath("//*[local-name()='property']")
        if (p.get("pid") or "").isdigit()
    }
    first_pid = PREF_FIRST_PID
    while any(pid in taken_pids
              for pid in range(first_pid, first_pid + 64)):
        first_pid = max(taken_pids) + 1

    data = build_pref_data(style_id, locale, change.session_id, zv)
    chunks = split_pref_chunks(data)
    change.chunk_count = len(chunks)
    for i, chunk in enumerate(chunks):
        prop = etree.SubElement(root, f"{{{CUSTOM_NS}}}property")
        prop.set("fmtid", PREF_FMTID)
        pid = first_pid + i
        prop.set("pid", str(pid))
        change.pids.append(pid)
        prop.set("name", f"ZOTERO_PREF_{i + 1}")
        v = etree.SubElement(prop, f"{{{VT_NS}}}lpwstr")
        v.text = chunk
    pkg.mark_modified(CUSTOM_PROPS_PART)
    if old_style and old_style != style_id:
        change.notes.append(f"style changed: {old_style} -> {style_id}")
    elif old_style:
        change.notes.append(f"style unchanged: {style_id}")
    return change
