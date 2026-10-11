"""Reference-list location + CSL-JSON parsing + existing-field mapping."""

from __future__ import annotations

import pytest
from lxml import etree

from zoteroconv.docx_io import DOCUMENT_PART, DocxPackage
from zoteroconv.parse_reference_list import (
    NS,
    find_reference_list,
    map_existing_fields_to_references,
    par_text,
    parse_dateish,
    w,
)
from zoteroconv.zotero_fields import find_zotero_fields


def _rl(path):
    pkg = DocxPackage.open(path)
    rl = find_reference_list(pkg)
    assert rl is not None, f"no reference list found in {path}"
    return rl


def test_typed_entries_parsed(gen_fixtures):
    rl = _rl(gen_fixtures["simple_bracket.docx"])
    assert par_text(rl.heading) == "References"
    assert [r.number for r in rl.references] == [1, 2, 3]
    assert all(r.number_source == "typed" for r in rl.references)
    r2 = rl.references[1]
    assert r2.csl["type"] == "article-journal"
    assert r2.csl["title"] == "Title two about things"
    assert r2.csl["author"] == [
        {"family": "Smith", "given": "J"},
        {"family": "Jones", "given": "AB"},
    ]
    assert r2.csl["container-title"] == "Journal Two"
    assert r2.csl["issued"] == {"date-parts": [["2021"]]}
    assert r2.csl["volume"] == "2"
    assert r2.csl["issue"] == "2"
    assert r2.csl["page"] == "11-20"
    assert r2.csl["DOI"] == "10.1000/j2.2021.2"
    assert r2.confidence == "high"
    assert r2.raw.startswith("Smith J, Jones AB.")


def test_auto_numbered_entries_resolved_from_numbering_xml(gen_fixtures):
    rl = _rl(gen_fixtures["auto_numbered_refs.docx"])
    assert [r.number for r in rl.references] == [1, 2, 3, 4]
    assert all(r.number_source == "auto" for r in rl.references)
    # a naive paragraph counter would have produced 4..7 (3 body list items
    # precede the references); the resolver must NOT do that
    assert rl.references[0].csl["title"] == "Auto title one"
    assert rl.references[3].csl["title"] == "Auto title four"


def test_reference_list_in_table(gen_fixtures):
    rl = _rl(gen_fixtures["refs_in_table.docx"])
    assert [r.number for r in rl.references] == [1, 2, 3]
    assert all(r.number_source == "table" for r in rl.references)
    assert rl.references[1].csl["title"] == "Table title two"
    # region must contain the w:tbl element itself
    assert any(el.tag == w("tbl") for el in rl.region_elements)


def test_missing_reference_number(gen_fixtures):
    rl = _rl(gen_fixtures["missing_ref_number.docx"])
    assert [r.number for r in rl.references] == [1, 2, 4]
    assert rl.missing_numbers == [3]
    assert any("missing" in n for n in rl.notes)


def test_wrapped_entries_merge(gen_fixtures):
    rl = _rl(gen_fixtures["wrapped_entries.docx"])
    assert len(rl.references) == 2
    r1 = rl.references[0]
    assert "continues onto a second paragraph" in r1.raw
    assert len(r1.elements) == 2
    assert r1.csl["container-title"] == "Journal One"
    r2 = rl.references[1]
    assert "hyphen- ated continuation" in r2.raw
    assert len(r2.elements) == 2


def test_appendix_not_swallowed(gen_fixtures):
    rl = _rl(gen_fixtures["appendix_after_refs.docx"])
    assert [r.number for r in rl.references] == [1, 2, 3]
    assert rl.stopped_at is not None
    assert par_text(rl.stopped_at).startswith("Appendix A")
    assert any("heading" in n for n in rl.notes)
    # the appendix paragraph is NOT part of the region
    assert rl.stopped_at not in rl.region_elements


def test_reflist_entries_types_and_accessed(gen_fixtures):
    rl = _rl(gen_fixtures["reflist_not_converted.docx"])
    r1, r2 = rl.references
    assert r1.csl["type"] == "webpage"
    assert r1.csl["title"] == "Internet things"       # [Internet] stripped
    assert r1.csl["publisher"] == "Web Publisher"
    assert r1.csl["publisher-place"] == "London"
    assert r1.csl["issued"] == {"date-parts": [["2020"]]}
    assert r1.csl["accessed"] == {"date-parts": [["2024", 3, 3]]}
    assert r1.csl["URL"] == "https://example.org/thing"
    assert r2.csl["type"] == "article-journal"
    assert r2.csl["volume"] == "4" and r2.csl["issue"] == "1"
    assert r2.csl["page"] == "47-67"


def test_title_is_never_the_whole_raw(gen_fixtures, real_docx):
    lists = [_rl(p) for p in gen_fixtures.values()] + [_rl(real_docx)]
    checked = 0
    for rl in lists:
        for r in rl.references:
            title = r.csl.get("title", "")
            assert title != r.raw.strip(), f"title == raw for ref {r.number}"
            assert title != " ".join(r.raw.split()), f"title == raw for ref {r.number}"
            checked += 1
    assert checked >= 50


def test_real_vancouver_reference_list(real_docx):
    rl = _rl(real_docx)
    assert par_text(rl.heading) == "References"
    assert [r.number for r in rl.references] == list(range(1, 17))
    assert rl.missing_numbers == []
    # region = paragraphs #54..#70 (17 elements incl. the empty #55)
    assert len(rl.region_elements) == 17
    # stopped at the existing ZOTERO_BIBL paragraph (decision 2)
    assert any("ZOTERO_BIBL" in n for n in rl.notes)

    r1 = rl.references[0]
    assert r1.csl["author"] == [
        {"family": "Ünvan", "given": "YA"},
        {"family": "Ergenç", "given": "C"},
    ]
    assert r1.csl["volume"] == "4" and r1.csl["page"] == "47-67"

    r4 = rl.references[3]
    assert r4.csl["DOI"] == "10.46299/j.isjmef.20240303.01"
    assert r4.csl["author"] == [
        {"family": "Havryk", "given": "A"},
        {"family": "Nazarova", "given": "T"},
    ]
    assert r4.confidence == "high"

    r6 = rl.references[5]
    assert {"family": "Gonzalez De Lena", "given": "MT"} in r6.csl["author"]
    assert r6.csl["page"] == "19232-41"

    r7 = rl.references[6]
    assert len(r7.csl["author"]) == 6
    assert any("et al" in i for i in r7.issues)
    assert r7.csl["page"] == "e12093"

    r9 = rl.references[8]
    assert r9.csl["author"] == [
        {"name": "Financial Conduct Authority"},
        {"name": "Bank of England"},
    ]

    r11 = rl.references[10]
    assert r11.csl["type"] == "report"
    assert r11.csl["number"] == "BIS Working Papers No 1194"

    r12 = rl.references[11]
    assert r12.csl["type"] == "book"
    assert r12.csl["publisher"] == "OECD Publishing"
    assert r12.csl["publisher-place"] == "Paris"
    assert r12.csl["author"] == [
        {"name": "Organisation for Economic Co-operation and Development"}
    ]

    r13 = rl.references[12]
    assert r13.csl["type"] == "report"
    assert r13.csl["number"] == "NIST AI 100-1"

    r16 = rl.references[15]
    assert r16.csl["type"] == "webpage"
    assert r16.csl["issued"] == {"date-parts": [["2024", 5, 7]]}
    assert r16.csl["author"] == [{"name": "Revolut"}]


def test_real_vancouver_existing_field_maps_by_doi(real_docx):
    pkg = DocxPackage.open(real_docx)
    rl = find_reference_list(pkg)
    fields = find_zotero_fields(pkg.document, DOCUMENT_PART)
    maps = map_existing_fields_to_references(fields, rl.references)
    assert len(maps) == 1
    m = maps[0]
    assert m.citation_id == "xO3paPuA"
    assert m.item_key == "NU469JIQ"
    assert m.ref_number == 4
    assert m.method == "doi"


def test_partly_converted_mappings_doi_and_title(gen_fixtures):
    pkg = DocxPackage.open(gen_fixtures["partly_converted.docx"])
    rl = find_reference_list(pkg)
    assert [r.number for r in rl.references] == [1, 2, 3, 4]
    assert any("ZOTERO_BIBL" in n for n in rl.notes)
    fields = find_zotero_fields(pkg.document, DOCUMENT_PART)
    maps = map_existing_fields_to_references(fields, rl.references)
    by_id = {m.citation_id: m for m in maps}
    assert by_id["FIXTUREA"].ref_number == 4
    assert by_id["FIXTUREA"].method == "doi"
    assert by_id["FIXTUREB"].ref_number == 4
    assert by_id["FIXTUREB"].method == "title"


def test_no_heading_returns_none(tmp_path):
    from docx import Document

    p = tmp_path / "no_refs.docx"
    doc = Document()
    doc.add_paragraph("Just some text [1] without any reference list.")
    doc.save(p)
    pkg = DocxPackage.open(p)
    assert find_reference_list(pkg) is None


def test_parse_dateish_variants():
    assert parse_dateish("2024 Mar 3") == ["2024", 3, 3]
    assert parse_dateish("7 May 2024") == ["2024", 5, 7]
    assert parse_dateish("May 7, 2024") == ["2024", 5, 7]
    assert parse_dateish("Jun 2024") == ["2024", 6]
    assert parse_dateish("2024") == ["2024"]
    assert parse_dateish("2024-06-01") == ["2024", 6, 1]
    assert parse_dateish("not a date") is None
