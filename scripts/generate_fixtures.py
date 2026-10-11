#!/usr/bin/env python3
"""
scripts/generate_fixtures.py -- build the Phase-1 test fixtures
programmatically (python-docx + lxml), per the spec:

  simple [1]; grouped [2,5]; range [3-6]; superscript; citation split over
  3 runs; two citations in one paragraph; citation in a footnote; numbers in
  a table that must not convert; "Table 2 [2]" exclusion; auto-numbered
  reference list; reference list in a table; missing reference number;
  reference list must not be converted; appendix after references; entries
  wrapped across paragraphs; and (user decision) a PARTLY CONVERTED document:
  one existing ZOTERO_ITEM field + raw [n] citations adjacent to it, an
  existing 1-entry ZOTERO_BIBL field after the typed list, and ZOTERO_PREF_*
  custom properties -- modelled 1:1 on tests/fixtures/real_vancouver.docx and
  the structural facts in docs/zotero_format.md.

Usage:
    python scripts/generate_fixtures.py [OUTDIR]        # default tests/fixtures/generated
Importable API:
    generate_all(outdir) -> dict[name, Path]

Everything is generated from scratch at test time; no hand-maintained binary
fixtures are committed for these scenarios.
"""

from __future__ import annotations

import json
import os
import re
import sys
import zipfile
from pathlib import Path

from docx import Document
from lxml import etree

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
CUST_NS = "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties"
VT_NS = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"
XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"

NS = {"w": W_NS, "r": R_NS, "ct": CT_NS, "rel": REL_NS}


def w(tag: str) -> str:
    return f"{{{W_NS}}}{tag}"


# ---------------------------------------------------------------------------
# zip surgery helpers (operate on a saved .docx path; rewrite the whole zip)
# ---------------------------------------------------------------------------

def _read_parts(path):
    infos, parts = [], {}
    with zipfile.ZipFile(path, "r") as zf:
        for info in zf.infolist():
            infos.append(info)
            parts[info.filename] = zf.read(info.filename)
    return infos, parts


def _write_parts(path, infos, parts):
    tmp = str(path) + ".tmp"
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        written = set()
        for info in infos:
            ni = zipfile.ZipInfo(info.filename, date_time=info.date_time)
            ni.compress_type = info.compress_type
            ni.external_attr = info.external_attr
            zout.writestr(ni, parts[info.filename])
            written.add(info.filename)
        for name, data in parts.items():
            if name not in written:
                ni = zipfile.ZipInfo(name, date_time=(1980, 1, 1, 0, 0, 0))
                ni.compress_type = zipfile.ZIP_DEFLATED
                zout.writestr(ni, data)
    os.replace(tmp, path)


def _edit_part(path, part_name, fn):
    infos, parts = _read_parts(path)
    root = etree.fromstring(parts[part_name])
    fn(root)
    parts[part_name] = etree.tostring(
        root, xml_declaration=True, encoding="UTF-8", standalone=True
    )
    _write_parts(path, infos, parts)


def _add_part(path, part_name, data: bytes, content_type: str | None,
              rel_type: str | None, rels_part: str = "_rels/.rels"):
    """Add a new part + optional content-type Override + optional relationship."""
    infos, parts = _read_parts(path)
    parts[part_name] = data
    if content_type is not None:
        ct = etree.fromstring(parts["[Content_Types].xml"])
        ov = etree.SubElement(ct, f"{{{CT_NS}}}Override")
        ov.set("PartName", "/" + part_name)
        ov.set("ContentType", content_type)
        parts["[Content_Types].xml"] = etree.tostring(
            ct, xml_declaration=True, encoding="UTF-8", standalone=True
        )
    if rel_type is not None:
        rels = etree.fromstring(parts[rels_part])
        existing = [
            int(m.group(1))
            for el in rels
            for m in [re.fullmatch(r"rId(\d+)", el.get("Id") or "")]
            if m
        ]
        rid = f"rId{max(existing, default=0) + 1}"
        rel = etree.SubElement(rels, f"{{{REL_NS}}}Relationship")
        rel.set("Id", rid)
        rel.set("Type", rel_type)
        # Target is resolved relative to the SOURCE part of rels_part:
        # root rels -> package path; word/_rels/document.xml.rels -> "word/"-relative
        source_dir = rels_part.rsplit("/_rels/", 1)[0] if "/_rels/" in rels_part else ""
        rel.set("Target", part_name[len(source_dir) + 1:] if source_dir else part_name)
        parts[rels_part] = etree.tostring(
            rels, xml_declaration=True, encoding="UTF-8", standalone=True
        )
    _write_parts(path, infos, parts)
    return parts


def _paragraphs(root):
    return root.xpath("//w:p", namespaces=NS)


def _par_text(par):
    return "".join(t.text or "" for t in par.xpath(".//w:t", namespaces=NS))


def _find_par_by_text(root, text):
    for par in _paragraphs(root):
        if _par_text(par) == text:
            return par
    raise LookupError(f"paragraph with text {text!r} not found")


# ---------------------------------------------------------------------------
# lxml run/paragraph builders (used for Zotero fields and injected structures)
# ---------------------------------------------------------------------------

def R_text(text, superscript=False, nsmap=None):
    r = etree.Element(w("r"), nsmap=nsmap)
    if superscript:
        rpr = etree.SubElement(r, w("rPr"))
        va = etree.SubElement(rpr, w("vertAlign"))
        va.set(w("val"), "superscript")
    t = etree.SubElement(r, w("t"))
    t.set(XML_SPACE, "preserve")
    t.text = text
    return r


def R_fld(kind, nsmap=None):
    r = etree.Element(w("r"), nsmap=nsmap)
    fc = etree.SubElement(r, w("fldChar"))
    fc.set(w("fldCharType"), kind)
    return r


def R_instr(instr, nsmap=None):
    r = etree.Element(w("r"), nsmap=nsmap)
    it = etree.SubElement(r, w("instrText"))
    it.set(XML_SPACE, "preserve")
    it.text = instr
    return r


def zotero_field_runs(instr, preview_segments, nsmap=None):
    """begin / instrText / separate / preview runs / end -- docs/zotero_format.md 3.1."""
    runs = [R_fld("begin", nsmap), R_instr(instr, nsmap), R_fld("separate", nsmap)]
    runs += [R_text(seg, nsmap=nsmap) for seg in preview_segments]
    runs.append(R_fld("end", nsmap))
    return runs


# ---------------------------------------------------------------------------
# shared fixture content
# ---------------------------------------------------------------------------

REFS3 = [
    "1. Author A, Author B. Title one of the study. Journal One. 2020;1(1):1-10. "
    "https://doi.org/10.1000/j1.2020.1",
    "2. Smith J, Jones AB. Title two about things. Journal Two. 2021;2(2):11-20. "
    "https://doi.org/10.1000/j2.2021.2",
    "3. Brown C. Title three of analysis. Journal Three. 2022;3(3):21-30. "
    "https://doi.org/10.1000/j3.2022.3",
]
REFS5 = REFS3 + [
    "4. Doe D, Roe R. Title four on methods. Journal Four. 2023;4(4):31-40. "
    "https://doi.org/10.1000/j4.2023.4",
    "5. Black E. Title five of results. Journal Five. 2023;5(5):41-50. "
    "https://doi.org/10.1000/j5.2023.5",
]
REFS6 = REFS5 + [
    "6. Green F, White G. Title six about discussion. Journal Six. 2024;6(6):51-60. "
    "https://doi.org/10.1000/j6.2024.6",
]


def add_ref_section(doc, refs, heading="References"):
    doc.add_paragraph(heading)          # plain paragraph, like the real docs
    for r in refs:
        doc.add_paragraph(r)


def sup_run(par, text):
    run = par.add_run(text)
    run.font.superscript = True
    return run


# ---------------------------------------------------------------------------
# individual fixture builders
# ---------------------------------------------------------------------------

def build_simple_bracket(path):
    doc = Document()
    doc.add_paragraph("Digital banking reshapes labour markets [1].")
    doc.add_paragraph("Others argue markets(2) respond slowly.")
    add_ref_section(doc, REFS3[:2] + [REFS3[2]])
    doc.save(path)


def build_grouped_bracket(path):
    doc = Document()
    doc.add_paragraph("Several studies agree [2,5] on this point.")
    doc.add_paragraph("More evidence exists [1, 3-5] overall.")
    add_ref_section(doc, REFS5)
    doc.save(path)


def build_range_dash(path):
    doc = Document()
    doc.add_paragraph("Hyphen range [3-6] appears here.")
    doc.add_paragraph("En dash range [3\u20136] appears here.")
    doc.add_paragraph("Em dash range [3\u20146] appears here.")
    add_ref_section(doc, REFS6)
    doc.save(path)


def build_superscript(path):
    doc = Document()
    p = doc.add_paragraph()
    p.add_run("Labour markets shift")
    sup_run(p, "1,")          # split over two adjacent superscript runs
    sup_run(p, "2")
    p.add_run(" and prices fall")
    sup_run(p, "3")
    p.add_run(".")
    p2 = doc.add_paragraph()
    p2.add_run("The ")
    sup_run(p2, "1")          # ordinal decoy: superscript 1 + normal 'st'
    p2.add_run("st decade was different; area ")
    p2.add_run("m")
    sup_run(p2, "2")          # unit decoy: single superscript digit after a letter
    p2.add_run(" per capita.")
    add_ref_section(doc, REFS3)
    doc.save(path)


def build_split_runs(path):
    doc = Document()
    p = doc.add_paragraph()
    p.add_run("Financial services ")
    p.add_run("[2,")           # citation split over three runs
    p.add_run("5")
    p.add_run("] improve.")
    p2 = doc.add_paragraph()
    p2.add_run("Retail markets")
    p2.add_run("(1")           # parenthetical split over two runs
    p2.add_run(") react.")
    add_ref_section(doc, REFS5)
    doc.save(path)


def build_two_in_paragraph(path):
    doc = Document()
    doc.add_paragraph("Prior work [1] and later analysis [2,3] agree.")
    add_ref_section(doc, REFS3)
    doc.save(path)


def build_footnote_citation(path):
    doc = Document()
    doc.add_paragraph("Claims appear here [1] and need FOOTNOTE_MARKER.")
    add_ref_section(doc, REFS3)
    doc.save(path)
    # surgery: real footnote part + footnoteReference run
    fn_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
        f'<w:footnotes xmlns:w="{W_NS}" xmlns:r="{R_NS}">'
        '<w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>'
        '<w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>'
        '<w:footnote w:id="1"><w:p><w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr>'
        '<w:footnoteRef/></w:r><w:r><w:t xml:space="preserve"> See also [2] for details.</w:t></w:r></w:p></w:footnote>'
        "</w:footnotes>"
    )
    _add_part(
        path,
        "word/footnotes.xml",
        fn_xml.encode("utf-8"),
        "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml",
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes",
        rels_part="word/_rels/document.xml.rels",
    )

    def repl(root):
        par = _find_par_by_text(root, "Claims appear here [1] and need FOOTNOTE_MARKER.")
        for t in par.xpath(".//w:t", namespaces=NS):
            if "FOOTNOTE_MARKER" in (t.text or ""):
                t.text = t.text.replace(" FOOTNOTE_MARKER", "")
        run = etree.Element(w("r"))
        rpr = etree.SubElement(run, w("rPr"))
        va = etree.SubElement(rpr, w("vertAlign"))
        va.set(w("val"), "superscript")
        ref = etree.SubElement(run, w("footnoteReference"))
        ref.set(w("id"), "1")
        par.append(run)

    _edit_part(path, "word/document.xml", repl)


def build_numbers_in_table(path):
    doc = Document()
    doc.add_paragraph("The table below summarizes results.")
    tbl = doc.add_table(rows=2, cols=2)
    tbl.cell(0, 0).paragraphs[0].add_run("5")
    cell_paren = tbl.cell(0, 1).paragraphs[0]
    cell_paren.add_run("(2)")
    cell_sup = tbl.cell(1, 0).paragraphs[0]
    sup_run(cell_sup, "3")
    tbl.cell(1, 1).paragraphs[0].add_run("[1]")
    doc.add_paragraph("Plain citation [2] after the table.")
    add_ref_section(doc, REFS3)
    doc.save(path)


def build_table_label_exclusion(path):
    doc = Document()
    doc.add_paragraph("Results are summarized in Table 2 [2].")
    doc.add_paragraph("See Fig. 3 (1) for the detailed diagram.")
    doc.add_paragraph("Other studies [2] confirm the finding.")
    add_ref_section(doc, REFS3)
    doc.save(path)


def build_auto_numbered_refs(path):
    doc = Document()
    doc.add_paragraph("Intro point one BODYLIST")
    doc.add_paragraph("Intro point two BODYLIST")
    doc.add_paragraph("Intro point three BODYLIST")
    doc.add_paragraph("Prior work shows effects [2] and confirms [4].")
    doc.add_paragraph("References")
    doc.add_paragraph("Author A, Author B. Auto title one. Journal One. 2020;1(1):1-10. "
                      "https://doi.org/10.2000/a1")
    doc.add_paragraph("Smith J. Auto title two. Journal Two. 2021;2(2):11-20.")
    doc.add_paragraph("Brown C. Auto title three. Journal Three. 2022;3(3):21-30.")
    doc.add_paragraph("Jones D. Auto title four. Journal Four. 2023;4(4):31-40.")
    doc.save(path)
    # surgery: numbering.xml with TWO separate num instances (body list numId=2,
    # reference list numId=1). A naive paragraph counter would number the
    # references 4..7; the resolver must produce 1..4.
    numbering_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
        f'<w:numbering xmlns:w="{W_NS}">'
        '<w:abstractNum w:abstractNumId="0"><w:multiLevelType w:val="singleLevel"/>'
        '<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/>'
        '<w:lvlText w:val="%1."/><w:lvlJc w:val="left"/></w:lvl></w:abstractNum>'
        '<w:abstractNum w:abstractNumId="1"><w:multiLevelType w:val="singleLevel"/>'
        '<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/>'
        '<w:lvlText w:val="%1."/><w:lvlJc w:val="left"/></w:lvl></w:abstractNum>'
        '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>'
        '<w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num>'
        "</w:numbering>"
    )
    # python-docx's default template ALREADY ships word/numbering.xml with its
    # content-type override and relationship -> replace bytes, never add a part.
    infos, parts = _read_parts(path)
    assert "word/numbering.xml" in parts
    parts["word/numbering.xml"] = numbering_xml.encode("utf-8")
    _write_parts(path, infos, parts)

    def set_numpr(root):
        for par in _paragraphs(root):
            text = _par_text(par)
            num_id = None
            if text.endswith("BODYLIST"):
                num_id = "2"
                for t in par.xpath(".//w:t", namespaces=NS):
                    if t.text and t.text.endswith(" BODYLIST"):
                        t.text = t.text[: -len(" BODYLIST")]
            elif text.startswith(("Author A, Author B. Auto", "Smith J. Auto",
                                  "Brown C. Auto", "Jones D. Auto")):
                num_id = "1"
            if num_id is None:
                continue
            ppr = par.xpath("./w:pPr", namespaces=NS)
            if ppr:
                ppr = ppr[0]
            else:
                ppr = etree.Element(w("pPr"))
                par.insert(0, ppr)
            numpr = etree.Element(w("numPr"))
            ilvl = etree.SubElement(numpr, w("ilvl"))
            ilvl.set(w("val"), "0")
            nid = etree.SubElement(numpr, w("numId"))
            nid.set(w("val"), num_id)
            ppr.insert(0, numpr)

    _edit_part(path, "word/document.xml", set_numpr)


def build_refs_in_table(path):
    doc = Document()
    doc.add_paragraph("Prior work [2] matters here.")
    doc.add_paragraph("References")
    tbl = doc.add_table(rows=3, cols=2)
    rows = [
        ("1", "Author A, Author B. Table title one. Journal One. 2020;1(1):1-10."),
        ("2", "Smith J. Table title two. Journal Two. 2021;2(2):11-20."),
        ("3", "Brown C. Table title three. Journal Three. 2022;3(3):21-30."),
    ]
    for i, (num, text) in enumerate(rows):
        tbl.cell(i, 0).paragraphs[0].add_run(num)
        tbl.cell(i, 1).paragraphs[0].add_run(text)
    doc.save(path)


def build_missing_ref_number(path):
    doc = Document()
    doc.add_paragraph("This cites [3] which does not exist, and [4] which does.")
    doc.add_paragraph("References")
    doc.add_paragraph("1. Author A. Missing number test one. Journal One. 2020;1(1):1-10.")
    doc.add_paragraph("2. Author B. Missing number test two. Journal Two. 2021;2(2):11-20.")
    doc.add_paragraph("4. Author D. Missing number test four. Journal Four. 2023;4(4):31-40.")
    doc.save(path)


def build_reflist_not_converted(path):
    doc = Document()
    doc.add_paragraph("Studies confirm [1] the trend.")
    doc.add_paragraph("References")
    doc.add_paragraph(
        "1. Webby O. Internet things [Internet]. London: Web Publisher; 2020 "
        "[cited 2024 Mar 3]. Available from: https://example.org/thing"
    )
    doc.add_paragraph(
        "2. Numby P. Counts of things. Journal of Numbers. 2023;4(1):47-67. "
        "https://doi.org/10.3000/n1"
    )
    doc.save(path)


def build_appendix_after_refs(path):
    doc = Document()
    doc.add_paragraph("The claim is supported [1].")
    add_ref_section(doc, REFS3)
    doc.add_paragraph("Appendix A: Survey instrument", style="Heading 1")
    doc.add_paragraph("Questions below use scales [1] and [2] from the literature.")
    doc.save(path)


def build_wrapped_entries(path):
    doc = Document()
    doc.add_paragraph("Evidence exists [2] for wrapping.")
    doc.add_paragraph("References")
    doc.add_paragraph("1. Author A, Author B. A very long wrapped title that continues")
    doc.add_paragraph("onto a second paragraph. Journal One. 2020;1(1):1-10.")
    doc.add_paragraph("2. Author C. Another wrapped title ending with a hyphen-")
    doc.add_paragraph("ated continuation. Journal Two. 2021;2(2):11-20.")
    doc.save(path)


# ---------------------------------------------------------------------------
# partly-converted fixture (user decisions 1/2/4; modelled on real_vancouver)
# ---------------------------------------------------------------------------

PREF_DATA = (
    '<data data-version="3" zotero-version="10.0.6">'
    '<session id="FIXTURE1"/>'
    '<style id="http://www.zotero.org/styles/harvard-cite-them-right" '
    'locale="en-US" hasBibliography="1" bibliographyStyleHasBeenSet="1"/>'
    '<prefs>'
    '<pref name="fieldType" value="Field"/>'
    '<pref name="automaticJournalAbbreviations" value="true"/>'
    '</prefs>'
    '</data>'
)

# itemData copied key-for-key from tests/fixtures/zotero_sample.docx (Phase 0
# ground truth): id, type, then case-insensitive alphabetical, author+issued
# last; issued year is a STRING, month/day are ints.
HAVRYK_ITEMDATA = {
    "id": 1,
    "type": "article-journal",
    "abstract": (
        "The article examines the role and impact of artificial intelligence "
        "(AI) in the financial sector to determine its impact on the future "
        "workforce and the US's pivotal role in shaping technology trends "
        "around the world."
    ),
    "container-title": "International Science Journal of Management, Economics & Finance",
    "DOI": "10.46299/j.isjmef.20240303.01",
    "ISSN": "2720-6394",
    "issue": "3",
    "journalAbbreviation": "ISJMEF",
    "license": "https://creativecommons.org/licenses/by/4.0",
    "page": "1-9",
    "source": "DOI.org (Crossref)",
    "title": (
        "Artificial intelligence and its role in the labor market and "
        "financial sector itself: US point of view"
    ),
    "title-short": (
        "Artificial intelligence and its role in the labor market and "
        "financial sector itself"
    ),
    "volume": "3",
    "author": [
        {"family": "Havryk", "given": "Alina"},
        {"family": "Nazarova", "given": "Tetiana"},
    ],
    "issued": {"date-parts": [["2024", 6, 1]]},
}

BIBL_PREVIEW = (
    "Havryk, A., Nazarova, T., 2024. Artificial intelligence and its role in "
    "the labor market and financial sector itself: US point of view. "
    "Int. Sci. J. Manag. Econ. Finance 3, 1\u20139. "
    "https://doi.org/10.46299/j.isjmef.20240303.01"
)


def citation_instr(citation_id, item_data, uri, preview):
    payload = {
        "citationID": citation_id,
        "properties": {
            "unsorted": False,
            "formattedCitation": preview,
            "plainCitation": preview,
            "noteIndex": 0,
        },
        "citationItems": [{"id": 1, "uris": [uri], "itemData": item_data}],
        "schema": "https://github.com/citation-style-language/schema/raw/master/csl-citation.json",
    }
    json_str = json.dumps(payload, ensure_ascii=False, separators=(",", ":"))
    return f" ADDIN ZOTERO_ITEM CSL_CITATION {json_str} "


BIBL_INSTR = ' ADDIN ZOTERO_BIBL {"uncited":[],"omitted":[],"custom":[]} CSL_BIBLIOGRAPHY '
ITEM_URI = "http://zotero.org/users/local/FIXLIB01/items/FIXITEM1"


def build_partly_converted(path):
    doc = Document()
    doc.add_paragraph("Digital Banking Report: Fixture")
    doc.add_paragraph("1. Introduction")
    p = doc.add_paragraph()
    p.add_run("AI reshapes financial services [1]. Blockchain matters [2,3].")
    doc.add_paragraph("FIELD_A_PARAGRAPH")     # replaced by surgery (existing field + raw [4])
    doc.add_paragraph("FIELD_B_PARAGRAPH")     # replaced by surgery (title-only itemData)
    doc.add_paragraph("References")
    doc.add_paragraph(
        "1. Author A, Author B. Title one of the study. Journal One. 2020;1(1):1-10. "
        "https://doi.org/10.1000/j1.2020.1"
    )
    doc.add_paragraph(
        "2. Smith J, Jones AB. Title two about things. Journal Two. 2021;2(2):11-20. "
        "https://doi.org/10.1000/j2.2021.2"
    )
    doc.add_paragraph(
        "3. Brown C. Title three of analysis. Journal Three. 2022;3(3):21-30. "
        "https://doi.org/10.1000/j3.2022.3"
    )
    doc.add_paragraph(
        "4. Havryk A, Nazarova T. Artificial intelligence and its role in the labor "
        "market and financial sector itself: US point of view. International Science "
        "Journal of Management, Economics & Finance. 2024;3(3):1-9. "
        "https://doi.org/10.46299/j.isjmef.20240303.01"
    )
    doc.add_paragraph("")   # empty paragraph after the typed list (like real doc #71)
    doc.save(path)

    infos, parts = _read_parts(path)
    doc_root = etree.fromstring(parts["word/document.xml"])
    nsmap = {"w": W_NS}

    # -- paragraph with existing ZOTERO_ITEM field A (itemData WITH DOI) ------
    par_a = _find_par_by_text(doc_root, "FIELD_A_PARAGRAPH")
    new_a = etree.Element(w("p"), nsmap=nsmap)
    new_a.append(R_text("AI affects labour markets ", nsmap=nsmap))
    for r in zotero_field_runs(
        citation_instr("FIXTUREA", HAVRYK_ITEMDATA, ITEM_URI, "(Havryk and Nazarova, 2024)"),
        ["(Havryk", " and Nazarova, 2024)"],
        nsmap=nsmap,
    ):
        new_a.append(r)
    new_a.append(R_text(" [4].", nsmap=nsmap))
    par_a.addnext(new_a)
    par_a.getparent().remove(par_a)

    # -- paragraph with existing ZOTERO_ITEM field B (itemData WITHOUT DOI) ---
    itemdata_b = {k: v for k, v in HAVRYK_ITEMDATA.items() if k != "DOI"}
    par_b = _find_par_by_text(doc_root, "FIELD_B_PARAGRAPH")
    new_b = etree.Element(w("p"), nsmap=nsmap)
    new_b.append(R_text("Further evidence ", nsmap=nsmap))
    for r in zotero_field_runs(
        citation_instr("FIXTUREB", itemdata_b, ITEM_URI, "(Havryk and Nazarova, 2024)"),
        ["(Havryk and Nazarova, 2024)"],
        nsmap=nsmap,
    ):
        new_b.append(r)
    new_b.append(R_text(" supports [1,4].", nsmap=nsmap))
    par_b.addnext(new_b)
    par_b.getparent().remove(par_b)

    # -- existing 1-entry ZOTERO_BIBL field appended after the typed list -----
    body = doc_root.xpath("//w:body", namespaces=NS)[0]
    sect = body.xpath("./w:sectPr", namespaces=NS)
    bib_par = etree.Element(w("p"), nsmap=nsmap)
    ppr = etree.SubElement(bib_par, w("pPr"))
    pstyle = etree.SubElement(ppr, w("pStyle"))
    pstyle.set(w("val"), "Bibliography")
    bib_runs = [R_fld("begin", nsmap), R_instr(BIBL_INSTR, nsmap), R_fld("separate", nsmap)]
    bib_runs += [R_text("Havryk", nsmap=nsmap), R_text(BIBL_PREVIEW[len("Havryk"):], nsmap=nsmap)]
    for r in bib_runs:
        bib_par.append(r)
    end_par = etree.Element(w("p"), nsmap=nsmap)
    end_par.append(R_fld("end", nsmap))
    anchor = sect[0] if sect else None
    if anchor is not None:
        anchor.addprevious(bib_par)
        anchor.addprevious(end_par)
    else:
        body.append(bib_par)
        body.append(end_par)

    parts["word/document.xml"] = etree.tostring(
        doc_root, xml_declaration=True, encoding="UTF-8", standalone=True
    )

    # -- Bibliography paragraph style (sample definition, docs/zotero_format.md 4.5)
    styles_root = etree.fromstring(parts["word/styles.xml"])
    if not styles_root.xpath("//w:style[@w:styleId='Bibliography']", namespaces=NS):
        st = etree.SubElement(styles_root, w("style"))
        st.set(w("type"), "paragraph")
        st.set(w("styleId"), "Bibliography")
        name = etree.SubElement(st, w("name")); name.set(w("val"), "Bibliography")
        based = etree.SubElement(st, w("basedOn")); based.set(w("val"), "Normal")
        nxt = etree.SubElement(st, w("next")); nxt.set(w("val"), "Normal")
        uip = etree.SubElement(st, w("uiPriority")); uip.set(w("val"), "37")
        etree.SubElement(st, w("unhideWhenUsed"))
        st_ppr = etree.SubElement(st, w("pPr"))
        sp = etree.SubElement(st_ppr, w("spacing"))
        sp.set(w("after"), "0"); sp.set(w("line"), "240"); sp.set(w("lineRule"), "auto")
        ind = etree.SubElement(st_ppr, w("ind"))
        ind.set(w("left"), "720"); ind.set(w("hanging"), "720")
        parts["word/styles.xml"] = etree.tostring(
            styles_root, xml_declaration=True, encoding="UTF-8", standalone=True
        )

    # -- docProps/custom.xml with ZOTERO_PREF_1/2 (255-char chunks, pids 2,3) --
    chunks = [PREF_DATA[i : i + 255] for i in range(0, len(PREF_DATA), 255)]
    props = [
        f'<property fmtid="{{D5CDD505-2E9C-101B-9397-08002B2CF9AE}}" pid="{pid}" '
        f'name="ZOTERO_PREF_{i + 1}"><vt:lpwstr>{escape_xml(chunk)}</vt:lpwstr></property>'
        for i, (chunk, pid) in enumerate(zip(chunks, range(2, 2 + len(chunks))))
    ]
    custom_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
        f'<Properties xmlns="{CUST_NS}" xmlns:vt="{VT_NS}">' + "".join(props) + "</Properties>"
    )
    parts["docProps/custom.xml"] = custom_xml.encode("utf-8")
    ct = etree.fromstring(parts["[Content_Types].xml"])
    if not any(
        el.get("PartName") == "/docProps/custom.xml" for el in ct
    ):
        ov = etree.SubElement(ct, f"{{{CT_NS}}}Override")
        ov.set("PartName", "/docProps/custom.xml")
        ov.set(
            "ContentType",
            "application/vnd.openxmlformats-officedocument.custom-properties+xml",
        )
        parts["[Content_Types].xml"] = etree.tostring(
            ct, xml_declaration=True, encoding="UTF-8", standalone=True
        )
    rels = etree.fromstring(parts["_rels/.rels"])
    if not any(el.get("Target") == "docProps/custom.xml" for el in rels):
        existing = [
            int(m.group(1))
            for el in rels
            for m in [re.fullmatch(r"rId(\d+)", el.get("Id") or "")]
            if m
        ]
        rel = etree.SubElement(rels, f"{{{REL_NS}}}Relationship")
        rel.set("Id", f"rId{max(existing, default=0) + 1}")
        rel.set(
            "Type",
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties",
        )
        rel.set("Target", "docProps/custom.xml")
        parts["_rels/.rels"] = etree.tostring(
            rels, xml_declaration=True, encoding="UTF-8", standalone=True
        )

    _write_parts(path, infos, parts)


def escape_xml(s: str) -> str:
    return (
        s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
    )


# ---------------------------------------------------------------------------
# registry / CLI
# ---------------------------------------------------------------------------

BUILDERS = {
    "simple_bracket.docx": build_simple_bracket,
    "grouped_bracket.docx": build_grouped_bracket,
    "range_dash.docx": build_range_dash,
    "superscript.docx": build_superscript,
    "split_runs.docx": build_split_runs,
    "two_in_paragraph.docx": build_two_in_paragraph,
    "footnote_citation.docx": build_footnote_citation,
    "numbers_in_table.docx": build_numbers_in_table,
    "table_label_exclusion.docx": build_table_label_exclusion,
    "auto_numbered_refs.docx": build_auto_numbered_refs,
    "refs_in_table.docx": build_refs_in_table,
    "missing_ref_number.docx": build_missing_ref_number,
    "reflist_not_converted.docx": build_reflist_not_converted,
    "appendix_after_refs.docx": build_appendix_after_refs,
    "wrapped_entries.docx": build_wrapped_entries,
    "partly_converted.docx": build_partly_converted,
}


def generate_all(outdir) -> dict[str, Path]:
    outdir = Path(outdir)
    outdir.mkdir(parents=True, exist_ok=True)
    paths: dict[str, Path] = {}
    for name, builder in BUILDERS.items():
        p = outdir / name
        builder(str(p))
        paths[name] = p
    return paths


def main(argv=None) -> int:
    outdir = argv[0] if argv else "tests/fixtures/generated"
    paths = generate_all(outdir)
    for name, p in paths.items():
        print(f"generated {p} ({p.stat().st_size} bytes)")
    print(f"{len(paths)} fixtures in {outdir}")
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
