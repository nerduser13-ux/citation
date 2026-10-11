"""
Structural comparison of docx packages (decision h): keys, key ORDER and run
structure -- never values.  Used to prove a converted document reproduces the
structure of tests/fixtures/zotero_sample.docx.
"""

from __future__ import annotations

import re
from typing import Any

from zoteroconv.build_fields import DATE_VARS, NAME_VARS
from zoteroconv.docx_io import DOCUMENT_PART, DocxPackage
from zoteroconv.zotero_fields import NS, find_zotero_fields, w

JSON_INSTR_RE = re.compile(r"\{.*\}", re.S)
ARTIFACT_TAGS = {"proofErr", "lastRenderedPageBreak", "bookmarkStart",
                 "bookmarkEnd", "commentRangeStart", "commentRangeEnd"}


def _local(tag) -> str:
    return tag.rsplit("}", 1)[-1] if isinstance(tag, str) else str(tag)


def _attrs(el) -> dict | None:
    if el is None:
        return None
    return {_local(k): v for k, v in el.attrib.items()}


def par_ppr_signature(p) -> dict | None:
    ppr = p.find("./w:pPr", namespaces=NS)
    if ppr is None:
        return None
    return {
        "style": (ppr.find("./w:pStyle", namespaces=NS).get(w("val"))
                  if ppr.find("./w:pStyle", namespaces=NS) is not None else None),
        "spacing": _attrs(ppr.find("./w:spacing", namespaces=NS)),
        "ind": _attrs(ppr.find("./w:ind", namespaces=NS)),
        "children": [_local(c.tag) for c in ppr],
    }


def run_pattern(par) -> list[str]:
    """Semantic run sequence of one paragraph: Word artifacts (proofErr etc.)
    are skipped; consecutive text runs are collapsed to a single 't'."""
    out: list[str] = []
    for child in par:
        tag = _local(child.tag)
        if tag in ("pPr",) or tag in ARTIFACT_TAGS:
            continue
        if tag != "r":
            out.append(f"other:{tag}")
            continue
        kinds = []
        for sub in child:
            stag = _local(sub.tag)
            if stag == "fldChar":
                kinds.append("fld:" + (sub.get(w("fldCharType")) or "?"))
            elif stag == "instrText":
                kinds.append("instr")
            elif stag == "t":
                kinds.append("t")
            elif stag == "rPr":
                continue
            else:
                kinds.append("run-child:" + stag)
        for k in kinds or ["empty-run"]:
            if k == "t" and out and out[-1] == "t":
                continue                      # collapse consecutive text runs
            out.append(k)
    return out


def itemdata_order_class(d: dict) -> str:
    keys = list(d.keys())
    problems = []
    if keys[:2] != ["id", "type"]:
        problems.append("does not start with id,type")
    idx = {k: i for i, k in enumerate(keys)}
    plain = [k for k in keys[2:] if k not in NAME_VARS and k not in DATE_VARS]
    names = [k for k in keys if k in NAME_VARS]
    dates = [k for k in keys if k in DATE_VARS]
    if plain != sorted(plain, key=str.lower):
        problems.append("plain variables not case-insensitive alphabetical")
    if names and plain and idx[names[0]] < idx[plain[-1]]:
        problems.append("name variables interleaved with plain variables")
    if names != sorted(names, key=str.lower):
        problems.append("name variables not alphabetical")
    if dates and names and idx[dates[0]] < idx[names[-1]]:
        problems.append("date variables before name variables")
    if dates != sorted(dates, key=str.lower):
        problems.append("date variables not alphabetical")
    if problems:
        return "NON-CONFORMANT: " + "; ".join(problems)
    return "id,type | alpha vars | name vars | date vars (sample-conformant)"


def instr_form(instr: str) -> str:
    return JSON_INSTR_RE.sub("{}", instr)


def field_structure(f) -> dict:
    sig: dict[str, Any] = {
        "kind": f.kind,
        "instr_form": instr_form(f.instr),
    }
    if f.payload is not None:
        sig["payload_key_order"] = list(f.payload.keys())
        if f.kind == "ITEM" and isinstance(f.payload.get("properties"), dict):
            sig["properties_key_order"] = list(f.payload["properties"].keys())
            items = f.payload.get("citationItems") or []
            if items:
                sig["citationItem_key_order"] = list(items[0].keys())
                data = items[0].get("itemData")
                if isinstance(data, dict):
                    sig["itemData_order_class"] = itemdata_order_class(data)
                    sig["itemData_id_equals_item_id"] = (
                        data.get("id") == items[0].get("id"))
                    sig["uris_is_list_of_zotero_uris"] = all(
                        isinstance(u, str) and u.startswith("http://zotero.org/")
                        for u in (items[0].get("uris") or []))
    # paragraphs spanned by the field, and their run patterns
    begin_par, end_par = f.begin_par, f.end_par
    body = begin_par.getparent()
    pars = []
    if end_par.getparent() is body:
        i0, i1 = body.index(begin_par), body.index(end_par)
        pars = list(body)[i0:i1 + 1]
    else:
        pars = [begin_par, end_par]
    sig["paragraph_count"] = len(pars)
    sig["paragraph_run_patterns"] = [run_pattern(p) for p in pars]
    sig["begin_par_pPr"] = par_ppr_signature(begin_par)
    sig["end_par_pPr"] = par_ppr_signature(end_par)
    sig["end_par_text_empty"] = not "".join(
        t.text or "" for t in end_par.xpath(".//w:t", namespaces=NS)).strip() \
        if end_par is not begin_par else None
    return sig


def style_definition_signature(styles_root, style_id: str) -> dict | None:
    if styles_root is None:
        return None
    hits = styles_root.xpath(f"//w:style[@w:styleId='{style_id}']",
                             namespaces=NS)
    if not hits:
        return None
    st = hits[0]
    sig: dict[str, Any] = {
        "type": st.get(w("type")),
        "children": [_local(c.tag) for c in st if _local(c.tag) != "rsid"],
    }
    for child in st:
        tag = _local(child.tag)
        if tag == "pPr":
            sig["pPr"] = {
                "spacing": _attrs(child.find("./w:spacing", namespaces=NS)),
                "ind": _attrs(child.find("./w:ind", namespaces=NS)),
            }
        elif tag in ("name", "basedOn", "next", "uiPriority"):
            sig[tag] = child.get(w("val"))
        elif tag == "rsid":
            continue                      # session artifact, deliberately ignored
    return sig


def package_signature(pkg: DocxPackage) -> dict:
    sig: dict[str, Any] = {}
    names = pkg.part_names()
    sig["has_parts"] = {
        "custom_props": "docProps/custom.xml" in names,
        "footnotes": "word/footnotes.xml" in names,
        "endnotes": "word/endnotes.xml" in names,
        "numbering": "word/numbering.xml" in names,
    }
    ct = pkg.content_types
    if ct is not None:
        sig["content_type_overrides"] = {
            el.get("PartName"): el.get("ContentType")
            for el in ct.xpath("/*[local-name()='Types']/"
                               "*[local-name()='Override']")
            if el.get("PartName") == "/docProps/custom.xml"
        }
    rels = pkg.root_rels
    if rels is not None:
        sig["root_rel_types"] = sorted(
            el.get("Type").rsplit("/", 1)[-1] for el in rels)
    custom = pkg.custom_props
    if custom is not None:
        props = []
        for p in custom.xpath("//*[local-name()='property']"):
            v = p.xpath("*[local-name()='lpwstr']")
            props.append({
                "name_form": re.sub(r"\d+", "N", p.get("name") or ""),
                "fmtid": p.get("fmtid"),
                "pid_relative_order": int(p.get("pid") or 0),
                "value_is_255_or_last": len(v[0].text or "") if v else None,
            })
        sig["custom_props"] = props
    sig["styles.Bibliography"] = style_definition_signature(
        pkg.styles, "Bibliography")
    fields = find_zotero_fields(pkg.document, DOCUMENT_PART)
    sig["field_count_by_kind"] = {
        "ITEM": sum(1 for f in fields if f.kind == "ITEM"),
        "BIBL": sum(1 for f in fields if f.kind == "BIBL"),
    }
    sig["fields"] = [field_structure(f) for f in fields]
    return sig


def diff_signatures(a: Any, b: Any, path: str = "") -> list[str]:
    """Recursive structural diff.  Returns human-readable difference lines
    ('' when the structures are equal)."""
    out: list[str] = []
    if isinstance(a, dict) and isinstance(b, dict):
        for k in sorted(set(a) | set(b), key=str):
            p = f"{path}.{k}" if path else str(k)
            if k not in a:
                out.append(f"{p}: MISSING in A, B={b[k]!r}")
            elif k not in b:
                out.append(f"{p}: A={a[k]!r}, MISSING in B")
            else:
                out.extend(diff_signatures(a[k], b[k], p))
    elif isinstance(a, list) and isinstance(b, list):
        if len(a) != len(b):
            out.append(f"{path}: list length A={len(a)} B={len(b)}")
        for i, (x, y) in enumerate(zip(a, b)):
            out.extend(diff_signatures(x, y, f"{path}[{i}]"))
    else:
        if a != b:
            out.append(f"{path}: A={a!r} B={b!r}")
    return out
