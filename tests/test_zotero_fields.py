"""Existing Zotero field discovery/parsing (structure per docs/zotero_format.md)."""

from __future__ import annotations

from zoteroconv.docx_io import DOCUMENT_PART, DocxPackage
from zoteroconv.zotero_fields import (
    NS,
    field_runs,
    find_zotero_fields,
    parse_instruction,
    w,
    zotero_field_zones,
)


def test_sample_has_one_item_and_one_bibl_field(sample_docx):
    pkg = DocxPackage.open(sample_docx)
    fields = find_zotero_fields(pkg.document, DOCUMENT_PART)
    assert [f.kind for f in fields] == ["ITEM", "BIBL"]
    assert all(f.closed for f in fields)

    item = fields[0]
    assert item.citation_id == "7yuwLzUY"
    assert list(item.payload.keys()) == [
        "citationID",
        "properties",
        "citationItems",
        "schema",
    ]
    assert item.payload["schema"] == (
        "https://github.com/citation-style-language/schema/raw/master/"
        "csl-citation.json"
    )
    assert list(item.payload["properties"].keys()) == [
        "unsorted",
        "formattedCitation",
        "plainCitation",
        "noteIndex",
    ]
    ci = item.payload["citationItems"][0]
    assert list(ci.keys()) == ["id", "uris", "itemData"]
    assert ci["uris"] == ["http://zotero.org/users/local/XXnzLBe9/items/NU469JIQ"]
    data = ci["itemData"]
    assert data["issued"]["date-parts"] == [["2024", 6, 1]]  # year str, month/day int
    assert data["DOI"] == "10.46299/j.isjmef.20240303.01"

    bibl = fields[1]
    assert bibl.payload == {"uncited": [], "omitted": [], "custom": []}
    assert bibl.instr.endswith("CSL_BIBLIOGRAPHY ")
    assert bibl.instr.startswith(" ADDIN ZOTERO_BIBL ")


def test_real_doc_field(sample_docx, real_docx):
    pkg = DocxPackage.open(real_docx)
    fields = find_zotero_fields(pkg.document, DOCUMENT_PART)
    assert [f.kind for f in fields] == ["ITEM", "BIBL"]
    assert fields[0].citation_id == "xO3paPuA"
    assert fields[0].citation_items()[0]["uris"] == [
        "http://zotero.org/users/local/XXnzLBe9/items/NU469JIQ"
    ]


def test_partly_converted_has_three_fields(gen_fixtures):
    pkg = DocxPackage.open(gen_fixtures["partly_converted.docx"])
    fields = find_zotero_fields(pkg.document, DOCUMENT_PART)
    assert [(f.kind, f.citation_id) for f in fields] == [
        ("ITEM", "FIXTUREA"),
        ("ITEM", "FIXTUREB"),
        ("BIBL", None),
    ]
    # field B carries the same item WITHOUT DOI (title-match path, decision 1)
    assert "DOI" in fields[0].citation_items()[0]["itemData"]
    assert "DOI" not in fields[1].citation_items()[0]["itemData"]


def test_generated_plain_doc_has_no_fields(gen_fixtures):
    pkg = DocxPackage.open(gen_fixtures["simple_bracket.docx"])
    assert find_zotero_fields(pkg.document, DOCUMENT_PART) == []


def test_field_runs_zone_covers_preview_only(sample_docx):
    pkg = DocxPackage.open(sample_docx)
    root = pkg.document
    fields = find_zotero_fields(root, DOCUMENT_PART)
    item = fields[0]
    zone = field_runs(root, item)
    par = item.begin_par
    runs = par.xpath(".//w:r", namespaces=NS)
    texts = [
        "".join(t.text or "" for t in r.xpath("./w:t", namespaces=NS)) for r in runs
    ]
    # the plain body run before the field is NOT in the zone
    assert runs[0] not in zone
    assert texts[0].startswith("Impact on employment")
    # preview runs ARE in the zone
    preview = [r for r, t in zip(runs, texts) if t in ("(", "Havryk", " and Nazarova, 2024)")]
    assert preview and all(r in zone for r in preview)


def test_zones_union_and_fields_returned(sample_docx):
    pkg = DocxPackage.open(sample_docx)
    zone, fields = zotero_field_zones(pkg.document, DOCUMENT_PART)
    assert len(fields) == 2
    assert len(zone) > 0


def test_parse_instruction_non_zotero():
    kind, payload, err = parse_instruction(" PAGE  \\* MERGEFORMAT ")
    assert kind is None and payload is None and err is None


def test_parse_instruction_broken_json():
    kind, payload, err = parse_instruction(" ADDIN ZOTERO_ITEM CSL_CITATION {broken ")
    assert kind == "ITEM"
    assert payload is None
    assert "JSON parse error" in err
