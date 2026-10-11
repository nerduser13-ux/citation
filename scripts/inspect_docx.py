#!/usr/bin/env python3
"""
scripts/inspect_docx.py -- Phase 0 inspection tool for .docx files (Zotero-aware).

READ-ONLY: the input file is opened with zipfile in read mode and is never
written to, renamed, or touched in any way. All parsing happens in memory.

Usage:
    python scripts/inspect_docx.py FILE.docx

Sections printed:
  A. ZIP part list (name, uncompressed size, compressed size) + a raw scan of
     every part for the byte strings ZOTERO_ITEM / ZOTERO_BIBL / ZOTERO_PREF /
     fldChar, so fields hiding in unexpected parts are not missed.
  B. Every complete Zotero field in word/document.xml, word/footnotes.xml and
     word/endnotes.xml (from the w:fldChar type="begin" run to the matching
     w:fldChar type="end" run), pretty-printed, with the instrText payload and
     its JSON pretty-printed separately.
  C. The full contents of docProps/custom.xml (raw, as stored in the zip, which
     shows the XML-escaped property values) and every custom property, with the
     ZOTERO_PREF* values also shown unescaped (and pretty-printed when the
     unescaped value is itself XML).
  D. The [Content_Types].xml entry and the _rels/.rels relationship that point
     to docProps/custom.xml.
  E. The paragraph properties/styles around the bibliography field (paragraph
     style, indents, the heading paragraph before it) and the run properties of
     citation preview runs (rStyle etc.), with the referenced style definitions
     from word/styles.xml.
  F. Whether the ZOTERO_PREF value is split across ZOTERO_PREF_1, _2, ... and
     exactly how (chunk lengths, boundary characters, whether the concatenation
     is well-formed XML).

Requires: lxml.  Python 3.11+.
"""

from __future__ import annotations

import argparse
import copy
import json
import re
import sys
import zipfile
from xml.sax.saxutils import escape as xml_escape

from lxml import etree

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
CUST_NS = "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties"

NS = {"w": W_NS, "ct": CT_NS, "rel": REL_NS, "cust": CUST_NS}

FIELD_PARTS = ("word/document.xml", "word/footnotes.xml", "word/endnotes.xml")

BAR = "=" * 78
SUBBAR = "-" * 78


def w(tag: str) -> str:
    """Clark-notation tag in the wordprocessingml namespace."""
    return f"{{{W_NS}}}{tag}"


def clean(el):
    """
    Rebuild an element for DISPLAY keeping only the namespace declarations
    actually used in its subtree (lxml otherwise drags all ~40 root
    declarations onto every serialized fragment). Read-only helper; the
    source tree is not modified.
    """
    if el is None:
        return None
    inscope = dict(el.nsmap)  # prefix -> uri in scope (includes inherited)
    used_uris = set()
    for node in el.iter():
        if isinstance(node.tag, str) and node.tag.startswith("{"):
            used_uris.add(node.tag[1:].split("}", 1)[0])
        for k in node.attrib:
            if k.startswith("{"):
                used_uris.add(k[1:].split("}", 1)[0])
    uri2pref: dict[str, str] = {}
    for p, u in inscope.items():
        if p is not None and u in used_uris and u not in uri2pref:
            uri2pref[u] = p
    n = 0
    for u in used_uris:
        if u not in uri2pref:
            uri2pref[u] = f"ns{n}"
            n += 1
    nsmap = {
        p: u
        for u, p in uri2pref.items()
        if u != "http://www.w3.org/XML/1998/namespace"  # 'xml' is implicit
    }

    def rebuild(src):
        new = etree.Element(src.tag, nsmap=nsmap) if src is el else etree.Element(src.tag)
        for k, v in src.attrib.items():
            new.set(k, v)
        new.text = src.text
        new.tail = src.tail
        for ch in src:
            new.append(rebuild(ch))
        return new

    return rebuild(el)


def pretty(el) -> str:
    """Pretty-print an lxml element, namespace noise stripped (display only)."""
    if el is None:
        return "(none)"
    return etree.tostring(clean(el), pretty_print=True, encoding="unicode").rstrip("\n")


def inline(el) -> str:
    """Serialize an lxml element on one line, namespace noise stripped."""
    if el is None:
        return "(none)"
    return etree.tostring(clean(el), encoding="unicode")


def header(title: str) -> None:
    print()
    print(BAR)
    print(title)
    print(BAR)


def sub(title: str) -> None:
    print()
    print(SUBBAR)
    print(title)
    print(SUBBAR)


def paragraph_text(par) -> str:
    """Visible text of a paragraph (w:t descendants only; skips instrText)."""
    return "".join(t.text or "" for t in par.xpath(".//w:t", namespaces=NS))


def paragraph_index(par) -> int:
    """0-based index of this w:p among all w:p in its part."""
    return len(par.xpath("preceding::w:p", namespaces=NS))


def paragraph_style(par):
    """w:val of the paragraph's w:pStyle, or None."""
    vals = par.xpath("./w:pPr/w:pStyle/@w:val", namespaces=NS)
    return vals[0] if vals else None


# ---------------------------------------------------------------------------
# Field discovery
# ---------------------------------------------------------------------------

class Field:
    """One Word field: begin fldChar -> instrText runs -> separate -> end."""

    def __init__(self, begin_fldchar):
        self.begin_fldchar = begin_fldchar
        self.instr_elems: list = []
        self.separate_fldchar = None
        self.end_fldchar = None

    @property
    def begin_run(self):
        return self.begin_fldchar.getparent()

    @property
    def end_run(self):
        return self.end_fldchar.getparent() if self.end_fldchar is not None else None

    @property
    def instruction(self) -> str:
        return "".join((el.text or "") for el in self.instr_elems)

    @property
    def closed(self) -> bool:
        return self.end_fldchar is not None


def find_fields(root) -> list[Field]:
    """All fields in a part, in document order, using a stack (nesting-safe)."""
    fields: list[Field] = []
    stack: list[Field] = []
    for el in root.iter():
        if el.tag == w("fldChar"):
            ftype = el.get(w("fldCharType"))
            if ftype == "begin":
                stack.append(Field(el))
            elif ftype == "separate":
                if stack:
                    stack[-1].separate_fldchar = el
            elif ftype == "end":
                if stack:
                    f = stack.pop()
                    f.end_fldchar = el
                    fields.append(f)
        elif el.tag == w("instrText") and stack:
            stack[-1].instr_elems.append(el)
    # Unclosed fields (malformed documents) are reported too.
    for f in reversed(stack):
        fields.append(f)
    return fields


def slice_field_range(root, start_run, end_run):
    """
    Return a synthetic <w:field-range> element containing a structure-
    preserving copy of everything in document order from start_run (the run
    holding fldChar begin) through end_run (the run holding fldChar end),
    inclusive of end_run's whole subtree.
    """
    order = list(root.iter())
    try:
        i_start = order.index(start_run)
        i_end = order.index(end_run)
    except ValueError:
        return None
    # Extend the end index over end_run's descendants (pre-order traversal).
    i_end_last = i_end
    j = i_end + 1
    while j < len(order) and any(a is end_run for a in order[j].iterancestors()):
        i_end_last = j
        j += 1

    container = etree.Element(w("field-range"), nsmap=dict(root.nsmap))
    copied: dict[int, object] = {}
    for el in order[i_start : i_end_last + 1]:
        if any(id(a) in copied for a in el.iterancestors()):
            continue  # already included inside a deep-copied ancestor
        new = copy.deepcopy(el)
        parent = el.getparent()
        if parent is not None and id(parent) in copied:
            copied[id(parent)].append(new)
        else:
            container.append(new)
        copied[id(el)] = new
    return container


def extract_json(instruction: str):
    """Parse the JSON payload of an ADDIN ZOTERO_ITEM CSL_CITATION instruction."""
    idx = instruction.find("{")
    if idx == -1:
        return None, None
    try:
        obj, end = json.JSONDecoder().raw_decode(instruction, idx)
        trailing = instruction[end:].strip()
        return obj, trailing
    except json.JSONDecodeError as exc:
        return None, f"JSON parse error: {exc}"


def result_runs(field: Field, root) -> list:
    """w:r elements between the separate fldChar and the end fldChar."""
    if field.separate_fldchar is None or field.end_fldchar is None:
        return []
    order = list(root.iter())
    try:
        i0 = order.index(field.separate_fldchar)
        i1 = order.index(field.end_fldchar)
    except ValueError:
        return []
    runs = []
    for el in order[i0 : i1 + 1]:
        if el.tag == w("r"):
            runs.append(el)
    # First run is the one holding 'separate'; it is not a preview run.
    return [r for r in runs if not r.xpath("./w:fldChar", namespaces=NS)]


# ---------------------------------------------------------------------------
# Sections
# ---------------------------------------------------------------------------

def section_a(zf: zipfile.ZipFile) -> None:
    header("SECTION A: ZIP PARTS + raw byte scan for Zotero markers")
    print(f"{'part':<45} {'size':>8} {'compressed':>10}")
    for info in zf.infolist():
        print(f"{info.filename:<45} {info.file_size:>8} {info.compress_size:>10}")

    sub("A.2 raw occurrence counts per part (bytes, before any XML parsing)")
    markers = (b"ZOTERO_ITEM", b"ZOTERO_BIBL", b"ZOTERO_PREF", b"fldChar")
    print(f"{'part':<45} " + " ".join(f"{m.decode():>12}" for m in markers))
    for info in zf.infolist():
        data = zf.read(info.filename)
        counts = [data.count(m) for m in markers]
        if any(counts):
            print(f"{info.filename:<45} " + " ".join(f"{c:>12}" for c in counts))
    print("(parts with zero occurrences of all markers are omitted)")


def dump_field(part: str, root, field: Field, number: int) -> None:
    instruction = field.instruction
    sub(
        f"B.{number} Zotero field in {part} | instr starts: "
        f"{instruction.strip()[:60]!r}"
    )
    if not field.closed:
        print("!! FIELD IS NOT CLOSED (no matching fldChar end) -- malformed")
        return

    begin_par = field.begin_run.xpath("ancestor::w:p[1]", namespaces=NS)
    end_par = field.end_run.xpath("ancestor::w:p[1]", namespaces=NS)
    begin_par = begin_par[0] if begin_par else None
    end_par = end_par[0] if end_par else None
    if begin_par is not None:
        end_desc = (
            f"paragraph #{paragraph_index(end_par)}"
            if end_par is not None
            else "(no enclosing w:p found)"
        )
        print(
            f"location: begin run in paragraph #{paragraph_index(begin_par)} "
            f"(style={paragraph_style(begin_par)!r}), end run in {end_desc}"
        )
        print(f"paragraph text preview: {paragraph_text(begin_par)[:120]!r}")

    print()
    print("fldChar elements (verbatim):")
    print("  begin   : " + inline(field.begin_fldchar))
    if field.separate_fldchar is not None:
        print("  separate: " + inline(field.separate_fldchar))
    else:
        print("  separate: (none)")
    print("  end     : " + inline(field.end_fldchar))

    print()
    print(f"instrText runs: {len(field.instr_elems)}")
    for i, el in enumerate(field.instr_elems, 1):
        space = el.get("{http://www.w3.org/XML/1998/namespace}space")
        print(f"  instrText fragment {i} (xml:space={space!r}): {(el.text or '')!r}")

    print()
    print("instrText concatenated (raw):")
    print(instruction)

    obj, extra = extract_json(instruction)
    print()
    if obj is not None:
        print("instrText JSON payload (pretty-printed):")
        print(json.dumps(obj, indent=2, ensure_ascii=False))
        if extra:
            print(f"trailing text after JSON: {extra!r}")
    elif extra:
        print(f"no JSON payload in instruction. ({extra})")
    else:
        print("no JSON payload in instruction.")

    print()
    print("complete field XML (begin fldChar run -> end fldChar run;")
    print("intermediate whole paragraphs appear as <w:p>; the wrapper")
    print("<w:field-range> is synthetic, added by this script for display):")
    rng = slice_field_range(root, field.begin_run, field.end_run)
    print(pretty(rng))


def section_b(zf: zipfile.ZipFile) -> dict[str, tuple]:
    header("SECTION B: ZOTERO FIELDS (word/document.xml, footnotes.xml, endnotes.xml)")
    parsed: dict[str, tuple] = {}
    number = 0
    for part in FIELD_PARTS:
        if part not in zf.namelist():
            print(f"\n{part}: PART NOT PRESENT in this docx")
            continue
        root = etree.fromstring(zf.read(part))
        parsed[part] = (root, zf.read(part))
        fields = find_fields(root)
        zot_fields = [f for f in fields if "ZOTERO" in f.instruction]
        other = [f for f in fields if "ZOTERO" not in f.instruction]
        print(f"\n{part}: {len(fields)} field(s) total, {len(zot_fields)} Zotero field(s)")
        for f in other:
            print(f"  non-Zotero field, instr={f.instruction.strip()[:60]!r}")
        for f in zot_fields:
            number += 1
            dump_field(part, root, f, number)
    if number == 0:
        print("\nNo Zotero fields found in any of the scanned parts.")
    return parsed


def read_custom_props(zf: zipfile.ZipFile):
    """Return (raw_bytes, list of dicts) for docProps/custom.xml, or (None, [])."""
    if "docProps/custom.xml" not in zf.namelist():
        return None, []
    raw = zf.read("docProps/custom.xml")
    root = etree.fromstring(raw)
    props = []
    for p in root.findall(f"{{{CUST_NS}}}property"):
        value_el = p[0] if len(p) else None
        props.append(
            {
                "name": p.get("name"),
                "fmtid": p.get("fmtid"),
                "pid": p.get("pid"),
                "vt_tag": etree.QName(value_el).localname if value_el is not None else None,
                "text": (value_el.text or "") if value_el is not None else "",
            }
        )
    return raw, props


def section_c(zf: zipfile.ZipFile):
    header("SECTION C: docProps/custom.xml (full contents, escaped + unescaped)")
    raw, props = read_custom_props(zf)
    if raw is None:
        print("docProps/custom.xml: PART NOT PRESENT in this docx")
        return props
    print("C.1 raw file contents exactly as stored in the zip")
    print("    (ZOTERO_PREF values appear here in XML-ESCAPED form):")
    print(raw.decode("utf-8"))

    sub("C.2 every custom property (values UNESCAPED)")
    for p in props:
        print(f"\nname={p['name']!r}  fmtid={p['fmtid']}  pid={p['pid']}  "
              f"vt-type={p['vt_tag']}  value-length={len(p['text'])}")
        print(f"  value (unescaped): {p['text']}")
        if p["name"] and p["name"].startswith("ZOTERO_PREF"):
            print("  value (re-escaped, as stored inside <vt:lpwstr>):")
            print("    " + xml_escape(p["text"]))
            try:
                inner = etree.fromstring(p["text"].encode("utf-8"))
                print("  value parses as XML; pretty-printed:")
                for line in pretty(inner).splitlines():
                    print("    " + line)
            except etree.XMLSyntaxError as exc:
                print(f"  value alone does NOT parse as XML: {exc}")
                print("  (this is expected when the value is a CHUNK of a larger")
                print("   document -- see SECTION F for the concatenation test)")
    return props


def section_d(zf: zipfile.ZipFile) -> None:
    header("SECTION D: [Content_Types].xml entry and _rels/.rels relationship for custom.xml")
    ct = etree.fromstring(zf.read("[Content_Types].xml"))
    matches = [
        el
        for el in ct
        if "custom.xml" in (el.get("PartName") or "")
        or "custom-properties" in (el.get("ContentType") or "")
    ]
    print("[Content_Types].xml entries pointing at docProps/custom.xml:")
    for el in matches:
        print("  " + etree.tostring(el, encoding="unicode"))
    if not matches:
        print("  (NONE FOUND)")

    rels = etree.fromstring(zf.read("_rels/.rels"))
    matches = [el for el in rels if "custom.xml" in (el.get("Target") or "")]
    print()
    print("_rels/.rels relationships pointing at docProps/custom.xml:")
    for el in matches:
        print("  " + etree.tostring(el, encoding="unicode"))
    if not matches:
        print("  (NONE FOUND)")


def lookup_style(styles_root, style_id: str):
    for st in styles_root.xpath("//w:style[@w:styleId=$sid]", sid=style_id, namespaces=NS):
        return st
    return None


def section_e(zf: zipfile.ZipFile, parsed: dict[str, tuple]) -> None:
    header("SECTION E: paragraph properties/styles around the bibliography field; "
           "run properties of citation previews")
    styles_root = None
    if "word/styles.xml" in zf.namelist():
        styles_root = etree.fromstring(zf.read("word/styles.xml"))
    wanted_styles: set[str] = set()

    any_field = False
    for part, (root, _raw) in parsed.items():
        for field in find_fields(root):
            if "ZOTERO" not in field.instruction or not field.closed:
                continue
            any_field = True
            is_bibl = "ZOTERO_BIBL" in field.instruction
            kind = "ZOTERO_BIBL" if is_bibl else "ZOTERO_ITEM"
            sub(f"E: {kind} field in {part}")

            begin_par = field.begin_run.xpath("ancestor::w:p[1]", namespaces=NS)[0]
            end_par = field.end_run.xpath("ancestor::w:p[1]", namespaces=NS)[0]
            body = begin_par.getparent()
            par_list = body.xpath(".//w:p", namespaces=NS)
            try:
                i_begin = par_list.index(begin_par)
                i_end = par_list.index(end_par)
            except ValueError:
                print("  (begin/end paragraphs are not siblings under one parent;"
                      " showing begin paragraph only)")
                par_list = [begin_par]
                i_begin = i_end = 0

            if i_begin > 0:
                prev_par = par_list[i_begin - 1]
                print("paragraph BEFORE the field begin (usually the heading):")
                print(f"  index=#{i_begin - 1} style={paragraph_style(prev_par)!r} "
                      f"text={paragraph_text(prev_par)[:100]!r}")
                ppr = prev_par.xpath("./w:pPr", namespaces=NS)
                print("  its w:pPr:")
                for line in pretty(ppr[0] if ppr else None).splitlines():
                    print("    " + line)
                if paragraph_style(prev_par):
                    wanted_styles.add(paragraph_style(prev_par))

            print()
            print(f"paragraphs spanned by the field: #{i_begin} .. #{i_end} "
                  f"({i_end - i_begin + 1} paragraph(s))")
            for i in range(i_begin, i_end + 1):
                par = par_list[i]
                st = paragraph_style(par)
                if st:
                    wanted_styles.add(st)
                print(f"\n  paragraph #{i}: style={st!r} text={paragraph_text(par)[:100]!r}")
                ppr = par.xpath("./w:pPr", namespaces=NS)
                print("  w:pPr:")
                for line in pretty(ppr[0] if ppr else None).splitlines():
                    print("    " + line)

            print()
            print("run properties of the field RESULT runs (preview text between "
                  "separate and end):")
            rruns = result_runs(field, root)
            if not rruns:
                print("  (no result runs)")
            for r in rruns:
                rpr = r.xpath("./w:rPr", namespaces=NS)
                txt = "".join(t.text or "" for t in r.xpath("./w:t", namespaces=NS))
                rpr_str = inline(rpr[0]) if rpr else "(no rPr)"
                print(f"  run text={txt[:60]!r} rPr={rpr_str}")
                for v in r.xpath("./w:rPr/w:rStyle/@w:val", namespaces=NS):
                    wanted_styles.add(v)

            print()
            print("fldChar-carrying runs (begin/separate/end), with their rPr:")
            for label, fc in (("begin", field.begin_fldchar),
                              ("separate", field.separate_fldchar),
                              ("end", field.end_fldchar)):
                if fc is None:
                    print(f"  {label}: (absent)")
                    continue
                run = fc.getparent()
                rpr = run.xpath("./w:rPr", namespaces=NS)
                rpr_str = inline(rpr[0]) if rpr else "(no rPr)"
                print(f"  {label}: run rPr={rpr_str}")
                for v in run.xpath("./w:rPr/w:rStyle/@w:val", namespaces=NS):
                    wanted_styles.add(v)

    if not any_field:
        print("No Zotero fields -> nothing to report for this section.")
        return

    sub("E.2 style definitions referenced above (from word/styles.xml)")
    if styles_root is None:
        print("word/styles.xml NOT PRESENT")
        return
    for sid in sorted(wanted_styles):
        st = lookup_style(styles_root, sid)
        print(f"\nw:styleId={sid!r}:")
        if st is None:
            print("  (style definition NOT FOUND in styles.xml)")
        else:
            for line in pretty(st).splitlines():
                print("  " + line)


def section_f(props) -> None:
    header("SECTION F: ZOTERO_PREF value splitting analysis")
    prefs = [p for p in props if (p["name"] or "").startswith("ZOTERO_PREF")]
    if not prefs:
        print("No ZOTERO_PREF* properties found.")
        return

    def suffix(name: str) -> int:
        m = re.fullmatch(r"ZOTERO_PREF(?:_(\d+))?", name)
        if not m:
            return -1
        return 0 if m.group(1) is None else int(m.group(1))

    prefs.sort(key=lambda p: suffix(p["name"]))
    print(f"{len(prefs)} ZOTERO_PREF* property/properties found:")
    for p in prefs:
        t = p["text"]
        print(f"\n  name={p['name']!r} pid={p['pid']} value-length={len(t)}")
        print(f"    first 70 chars: {t[:70]!r}")
        print(f"    last  70 chars: {t[-70:]!r}")

    lengths = [len(p["text"]) for p in prefs]
    if len(lengths) > 1:
        head = lengths[:-1]
        if len(set(head)) == 1:
            print(f"\nAll chunks except the last have identical length {head[0]} "
                  f"-> fixed-size character split at {head[0]} chars.")
        else:
            print(f"\nChunk lengths differ: {lengths} (not a single fixed size).")
        print("Chunk boundaries (last 40 chars of chunk N | first 40 chars of chunk N+1):")
        for a, b in zip(prefs, prefs[1:]):
            print(f"  {a['name']}: ...{a['text'][-40:]!r}")
            print(f"  {b['name']}: {b['text'][:40]!r}...")
        pids = [int(p["pid"]) for p in prefs if (p["pid"] or "").isdigit()]
        print(f"\npid values in suffix order: {pids} "
              f"({'ascending' if pids == sorted(pids) else 'NOT ascending'})")
    else:
        print("\nSingle chunk -> no splitting needed for this document.")

    concat = "".join(p["text"] for p in prefs)
    print(f"\nConcatenated value length: {len(concat)}")
    try:
        inner = etree.fromstring(concat.encode("utf-8"))
        print("Concatenation parses as well-formed XML. Pretty-printed:")
        for line in pretty(inner).splitlines():
            print("  " + line)
    except etree.XMLSyntaxError as exc:
        print(f"Concatenation does NOT parse as XML: {exc}")


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(
        description="Read-only inspector for Zotero structures inside a .docx."
    )
    ap.add_argument("docx", help="path to the .docx file to inspect")
    args = ap.parse_args(argv)

    try:
        zf = zipfile.ZipFile(args.docx, "r")
    except (OSError, zipfile.BadZipFile) as exc:
        print(f"ERROR: cannot open {args.docx!r}: {exc}", file=sys.stderr)
        return 1

    with zf:
        print(f"Inspecting: {args.docx}")
        section_a(zf)
        parsed = section_b(zf)
        props = section_c(zf)
        section_d(zf)
        section_e(zf, parsed)
        section_f(props)
    return 0


if __name__ == "__main__":
    sys.exit(main())
