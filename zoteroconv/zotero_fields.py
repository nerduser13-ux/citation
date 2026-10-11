"""
zoteroconv.zotero_fields -- discovery and parsing of existing Zotero Word
fields inside OOXML parts.

Structural facts implemented here are documented with evidence in
docs/zotero_format.md (Phase 0). Kinds:

  ITEM : ' ADDIN ZOTERO_ITEM CSL_CITATION {json} '
  BIBL : ' ADDIN ZOTERO_BIBL {"uncited":[],"omitted":[],"custom":[]} CSL_BIBLIOGRAPHY '

A field is the run sequence: begin fldChar run -> instrText run(s) ->
separate fldChar run -> preview run(s) -> end fldChar run (facts 3.1, 4.2).
"""

from __future__ import annotations

import json
import re
from dataclasses import dataclass, field as dc_field

from lxml import etree

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = {"w": W_NS}


def w(tag: str) -> str:
    return f"{{{W_NS}}}{tag}"


# Exact strings reproduced from the fixtures (docs/zotero_format.md §8).
CITATION_INSTR_HEAD = " ADDIN ZOTERO_ITEM CSL_CITATION "
BIBL_INSTR_HEAD = " ADDIN ZOTERO_BIBL "
BIBL_INSTR_TAIL = " CSL_BIBLIOGRAPHY "
BIBL_PAYLOAD_DEFAULT = '{"uncited":[],"omitted":[],"custom":[]}'
CITATION_SCHEMA_URL = (
    "https://github.com/citation-style-language/schema/raw/master/csl-citation.json"
)
PREF_FMTID = "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}"
PREF_CHUNK_SIZE = 255  # observed fixed split size (docs/zotero_format.md 2.4)


@dataclass
class ZoteroField:
    """One existing Zotero field found in a part."""

    part: str
    kind: str                      # 'ITEM' or 'BIBL'
    begin_fldchar: object          # w:fldChar element (type=begin)
    separate_fldchar: object | None
    end_fldchar: object
    instr: str                     # concatenated instrText (raw)
    payload: dict | None = None    # parsed JSON payload (None if unparsable)
    payload_error: str | None = None
    closed: bool = True

    @property
    def begin_run(self):
        return self.begin_fldchar.getparent()

    @property
    def end_run(self):
        return self.end_fldchar.getparent()

    @property
    def begin_par(self):
        found = self.begin_run.xpath("ancestor::w:p[1]", namespaces=NS)
        return found[0] if found else None

    @property
    def end_par(self):
        found = self.end_run.xpath("ancestor::w:p[1]", namespaces=NS)
        return found[0] if found else None

    @property
    def citation_id(self) -> str | None:
        if self.kind == "ITEM" and isinstance(self.payload, dict):
            return self.payload.get("citationID")
        return None

    def citation_items(self) -> list[dict]:
        if self.kind == "ITEM" and isinstance(self.payload, dict):
            items = self.payload.get("citationItems")
            if isinstance(items, list):
                return items
        return []


def parse_instruction(instr: str) -> tuple[str | None, dict | None, str | None]:
    """
    Classify a field instruction and parse its JSON payload.

    Returns (kind, payload, error). kind is 'ITEM', 'BIBL' or None.
    """
    if "ADDIN ZOTERO_ITEM" in instr:
        kind = "ITEM"
    elif "ADDIN ZOTERO_BIBL" in instr:
        kind = "BIBL"
    else:
        return None, None, None
    idx = instr.find("{")
    if idx == -1:
        return kind, None, "no JSON payload found in instruction"
    try:
        obj, _end = json.JSONDecoder().raw_decode(instr, idx)
        if not isinstance(obj, dict):
            return kind, None, "JSON payload is not an object"
        return kind, obj, None
    except json.JSONDecodeError as exc:
        return kind, None, f"JSON parse error: {exc}"


def find_zotero_fields(root, part: str = "word/document.xml") -> list[ZoteroField]:
    """
    All Zotero fields in one part's XML tree, in document order
    (nesting-safe stack scan; unclosed fields are reported with closed=False).
    """
    fields: list[ZoteroField] = []
    open_stack: list[dict] = []
    for el in root.iter():
        if el.tag == w("fldChar"):
            ftype = el.get(w("fldCharType"))
            if ftype == "begin":
                open_stack.append({"begin": el, "instr": [], "separate": None})
            elif ftype == "separate":
                if open_stack:
                    open_stack[-1]["separate"] = el
            elif ftype == "end":
                if open_stack:
                    rec = open_stack.pop()
                    instr = "".join((t.text or "") for t in rec["instr"])
                    kind, payload, err = parse_instruction(instr)
                    if kind is not None:
                        fields.append(
                            ZoteroField(
                                part=part,
                                kind=kind,
                                begin_fldchar=rec["begin"],
                                separate_fldchar=rec["separate"],
                                end_fldchar=el,
                                instr=instr,
                                payload=payload,
                                payload_error=err,
                                closed=True,
                            )
                        )
        elif el.tag == w("instrText") and open_stack:
            open_stack[-1]["instr"].append(el)
    for rec in reversed(open_stack):
        instr = "".join((t.text or "") for t in rec["instr"])
        kind, payload, err = parse_instruction(instr)
        if kind is not None:
            fields.append(
                ZoteroField(
                    part=part,
                    kind=kind,
                    begin_fldchar=rec["begin"],
                    separate_fldchar=rec["separate"],
                    end_fldchar=rec["begin"],  # no end known
                    instr=instr,
                    payload=payload,
                    payload_error=err or "field is not closed",
                    closed=False,
                )
            )
    return fields


def field_runs(root, field_obj: ZoteroField) -> set:
    """
    Set of every w:r ELEMENT from the field's begin run through its end run
    (document order, inclusive). Used to mark "inside existing Zotero field"
    zones in the character map (Decision 1: never touch these).

    NOTE: lxml element proxies are temporary objects; identity (`is`, `in`,
    dict keys) is only stable while a Python reference to the proxy is kept
    alive. Callers must hold on to the returned set (verified empirically);
    never use id() of elements across iterations.
    """
    if not field_obj.closed:
        runs = []
        collecting = False
        for r in root.iter(w("r")):
            if r is field_obj.begin_run:
                collecting = True
            if collecting:
                runs.append(r)
        return set(runs)
    runs = list(root.iter(w("r")))
    try:
        i0 = runs.index(field_obj.begin_run)
        i1 = runs.index(field_obj.end_run)
    except ValueError:
        return set()
    if i1 < i0:
        i0, i1 = i1, i0
    return set(runs[i0 : i1 + 1])


def zotero_field_zones(root, part: str = "word/document.xml") -> tuple[set, list[ZoteroField]]:
    """(union of field run ELEMENTS, field list) for one part."""
    fields = find_zotero_fields(root, part)
    zone: set = set()
    for f in fields:
        zone |= field_runs(root, f)
    return zone, fields
