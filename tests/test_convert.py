"""Phase 3: field building, bibliography replacement, prefs, idempotency and
structural equivalence with the real Zotero sample."""

from __future__ import annotations

import hashlib
import random
import zipfile
from pathlib import Path

import pytest
from lxml import etree

from zoteroconv.build_fields import (
    ZOTERO_BIBL_INSTR_PREFIX,
    ZOTERO_BIBL_INSTR_SUFFIX,
    ZOTERO_ITEM_INSTR_PREFIX,
    bibliography_paragraphs,
    bibl_instr,
    build_pref_data,
    citation_field_runs,
    citation_instr,
    fld_run,
    make_bibl_payload,
    ordered_item_data,
    read_prefs,
    rpr_without_superscript,
    split_pref_chunks,
    write_prefs,
)
from zoteroconv.convert import convert_document
from zoteroconv.detect_citations import detect_citations
from zoteroconv.docx_io import DOCUMENT_PART, DocxPackage
from zoteroconv.parse_reference_list import find_reference_list, par_text
from zoteroconv.preview import PreviewRenderer
from zoteroconv.resolve_items import resolve_document
from zoteroconv.structure_diff import (
    diff_signatures,
    itemdata_order_class,
    package_signature,
    run_pattern,
)
from zoteroconv.styles import StyleError, resolve_style_id
from zoteroconv.zotero_fields import NS, find_zotero_fields, w, zotero_field_zones

from test_resolve_items import (  # shared fakes
    SECRET_KEY,
    FakeTransport,
    crossref_route,
    dead_local,
    lib_item,
    local_routes,
    web_routes,
)

REPO = Path(__file__).resolve().parents[1]
HAVE_CSL = (REPO / "csl" / "elsevier-harvard.csl").is_file()
SILENT = lambda s: None  # noqa: E731


# ---------------------------------------------------------------------------
# helpers
# ---------------------------------------------------------------------------

def lib3():
    return [
        lib_item("KEY1", type="article-journal", title="Title one of the study",
                 DOI="10.1000/j1.2020.1",
                 author=[{"family": "Author", "given": "A"},
                         {"family": "Author", "given": "B"}],
                 issued={"date-parts": [["2020"]]},
                 **{"container-title": "Journal One", "volume": "1",
                    "issue": "1", "page": "1-10"}),
        lib_item("KEY2", type="article-journal", title="Title two about things",
                 DOI="10.1000/j2.2021.2",
                 author=[{"family": "Smith", "given": "J"},
                         {"family": "Jones", "given": "AB"}],
                 issued={"date-parts": [["2021"]]},
                 **{"container-title": "Journal Two", "volume": "2",
                    "issue": "2", "page": "11-20"}),
        lib_item("KEY3", type="article-journal", title="Title three of analysis",
                 DOI="10.1000/j3.2022.3",
                 author=[{"family": "Brown", "given": "C"}],
                 issued={"date-parts": [["2022"]]},
                 **{"container-title": "Journal Three", "volume": "3",
                    "issue": "3", "page": "21-30"}),
    ]


def lib5():
    return lib3() + [
        lib_item("KEY4", type="article-journal", title="Title four on methods",
                 DOI="10.1000/j4.2023.4",
                 author=[{"family": "Doe", "given": "D"},
                         {"family": "Roe", "given": "R"}],
                 issued={"date-parts": [["2023"]]},
                 **{"container-title": "Journal Four", "volume": "4",
                    "issue": "4", "page": "31-40"}),
        lib_item("KEY5", type="article-journal", title="Title five of results",
                 DOI="10.1000/j5.2023.5",
                 author=[{"family": "Black", "given": "E"}],
                 issued={"date-parts": [["2023"]]},
                 **{"container-title": "Journal Five", "volume": "5",
                    "issue": "5", "page": "41-50"}),
    ]


def resolve_local(path, items, **kw):
    t = FakeTransport(local_routes(items) + [crossref_route([])])
    kw.setdefault("emit", SILENT)
    return resolve_document(DocxPackage.open(path), transport=t, **kw)


def sha(path) -> str:
    return hashlib.sha256(Path(path).read_bytes()).hexdigest()


def differing_parts(a, b) -> set[str]:
    za, zb = zipfile.ZipFile(a), zipfile.ZipFile(b)
    na, nb = set(za.namelist()), set(zb.namelist())
    diff = set()
    for n in na | nb:
        da = za.read(n) if n in na else None
        db = zb.read(n) if n in nb else None
        if da != db:
            diff.add(n)
    return diff


def full_text(pkg) -> str:
    return "".join(t.text or ""
                   for t in pkg.document.xpath("//w:t", namespaces=NS))


def field_slice(pattern: list[str]) -> list[str]:
    return pattern[pattern.index("fld:begin"):pattern.index("fld:end") + 1]


def convert_simple(gen_fixtures, tmp_path, name="simple_bracket.docx",
                   items=None, cited=None, style="elsevier-harvard", **kw):
    src = gen_fixtures[name]
    rep = resolve_local(src, items if items is not None else lib3(),
                        cited_numbers=cited)
    out = tmp_path / (Path(name).stem + "_converted.docx")
    pkg = DocxPackage.open(src)
    kw.setdefault("rng", random.Random(42))
    res = convert_document(pkg, rep, style=style, output=str(out),
                           emit=SILENT, **kw)
    return res, out, rep


# ---------------------------------------------------------------------------
# unit tests: builders
# ---------------------------------------------------------------------------

def test_pref_data_matches_fixtures_verbatim():
    # the exact raw strings read from the two real fixtures this session
    assert build_pref_data(
        "http://www.zotero.org/styles/harvard-cite-them-right",
        "en-US", "88ppEdET", "10.0.6") == (
        '<data data-version="3" zotero-version="10.0.6">'
        '<session id="88ppEdET"/>'
        '<style id="http://www.zotero.org/styles/harvard-cite-them-right" '
        'locale="en-US" hasBibliography="1" bibliographyStyleHasBeenSet="1"/>'
        '<prefs><pref name="fieldType" value="Field"/>'
        '<pref name="automaticJournalAbbreviations" value="true"/></prefs>'
        '</data>')
    sample = build_pref_data(
        "http://www.zotero.org/styles/elsevier-harvard", None,
        "GahNCFud", "10.0.6")
    assert len(sample) == 299          # SS2.2: sample concatenation length
    assert 'locale=' not in sample     # sample carries no locale attribute


def test_split_pref_chunks_exact_255():
    data = "x" * 600
    ch = split_pref_chunks(data)
    assert [len(c) for c in ch] == [255, 255, 90]
    assert "".join(ch) == data
    assert split_pref_chunks("") == [""]


def test_ordered_item_data_key_order():
    csl = {"type": "article-journal", "title": "T", "DOI": "10.1/x",
           "volume": "1", "container-title": "J", "page": "1-2",
           "author": [{"family": "A", "given": "B"}],
           "issued": {"date-parts": [[2021, 6]]}, "URL": "http://e.org"}
    d = ordered_item_data(csl, 7)
    keys = list(d.keys())
    assert keys[0] == "id" and keys[1] == "type"
    assert keys[-2:] == ["author", "issued"]
    plain = keys[2:-2]
    assert plain == sorted(plain, key=str.lower)
    assert d["id"] == 7
    assert d["issued"]["date-parts"] == [["2021", 6]]   # year str, month int
    assert itemdata_order_class(d).startswith("id,type")


def test_field_run_builders():
    runs = citation_field_runs(" INSTR ", "preview", None)
    assert len(runs) == 5
    # bare fld runs: no rPr children
    for r in (runs[0], runs[2], runs[4]):
        assert r.find("./w:rPr", namespaces=NS) is None
    it = runs[1].find("./w:instrText", namespaces=NS)
    assert it.get("{http://www.w3.org/XML/1998/namespace}space") == "preserve"
    assert it.text == " INSTR "
    t = runs[3].find("./w:t", namespaces=NS)
    assert t.text == "preview"


def test_rpr_template_strips_superscript_only():
    r = etree.fromstring(
        f'<w:r xmlns:w="{NS["w"]}"><w:rPr><w:b/><w:vertAlign '
        f'w:val="superscript"/></w:rPr><w:t>x</w:t></w:r>')
    tpl = rpr_without_superscript(r)
    assert tpl is not None
    assert tpl.find("./w:b", namespaces=NS) is not None
    assert tpl.find("./w:vertAlign", namespaces=NS) is None
    # original run untouched
    assert r.find("./w:rPr/w:vertAlign", namespaces=NS) is not None
    assert rpr_without_superscript(None) is None


def test_bibliography_paragraph_shapes():
    instr = bibl_instr(make_bibl_payload([]))
    assert instr == (' ADDIN ZOTERO_BIBL {"uncited":[],"omitted":[],"custom":[]}'
                     ' CSL_BIBLIOGRAPHY ')
    pars = bibliography_paragraphs(instr, ["Entry one.", "Entry two."])
    assert len(pars) == 3
    assert run_pattern(pars[0]) == ["fld:begin", "instr", "fld:separate", "t"]
    assert run_pattern(pars[1]) == ["t"]
    assert run_pattern(pars[2]) == ["fld:end"]
    assert pars[0].find("./w:pPr/w:pStyle", namespaces=NS).get(w("val")) == "Bibliography"
    assert pars[1].find("./w:pPr/w:pStyle", namespaces=NS).get(w("val")) == "Bibliography"
    assert pars[2].find("./w:pPr", namespaces=NS) is None      # SS4.2: no pPr
    assert "".join(t.text or "" for t in pars[2].xpath(".//w:t", namespaces=NS)) == ""


def test_styles_resolution_and_errors():
    assert resolve_style_id("harvard") == \
        "http://www.zotero.org/styles/harvard-cite-them-right"
    assert resolve_style_id("apa") == "http://www.zotero.org/styles/apa"
    assert resolve_style_id("vancouver") == \
        "http://www.zotero.org/styles/vancouver"
    assert resolve_style_id("elsevier-harvard") == \
        "http://www.zotero.org/styles/elsevier-harvard"
    assert resolve_style_id("http://www.zotero.org/styles/nature") == \
        "http://www.zotero.org/styles/nature"
    with pytest.raises(StyleError) as ei:
        resolve_style_id("Not A Style!!")
    msg = str(ei.value)
    for name in ("apa", "harvard", "ieee", "vancouver", "chicago",
                 "nature", "mla"):
        assert name in msg


def test_preview_fallback_is_provisional(tmp_path):
    r = PreviewRenderer("http://www.zotero.org/styles/does-not-exist-offline",
                        style_dirs=[tmp_path])
    assert r.provisional
    assert any("PROVISIONAL" in n for n in r.notes)
    items = [{"type": "article-journal", "title": "T",
              "author": [{"family": "Author", "given": "A"}],
              "issued": {"date-parts": [["2020"]]}}]
    inlines, entries = r.render([items], items)
    assert inlines == ["(Author, 2020)"]
    assert entries[0].startswith("Author, A. (2020) T.")


@pytest.mark.skipif(not HAVE_CSL, reason="offline CSL styles not downloaded")
def test_preview_citeproc_renders_sample_string():
    r = PreviewRenderer("http://www.zotero.org/styles/elsevier-harvard")
    assert "citeproc-py" in r.engine
    havryk = [{
        "type": "article-journal",
        "title": "Artificial intelligence and its role in the labor market",
        "author": [{"family": "Havryk", "given": "A"},
                   {"family": "Nazarova", "given": "T"}],
        "issued": {"date-parts": [["2024", 6, 1]]},
        "container-title": "International Science Journal of Management, "
                           "Economics & Finance",
    }]
    inlines, entries = r.render([havryk], havryk)
    # the sample's own preview text, reproduced by the style file
    assert inlines == ["(Havryk and Nazarova, 2024)"]
    assert entries[0].startswith("Havryk, A., Nazarova, T., 2024.")


# ---------------------------------------------------------------------------
# end-to-end conversions on generated fixtures
# ---------------------------------------------------------------------------

def test_convert_simple_end_to_end(gen_fixtures, tmp_path, sample_docx):
    src = gen_fixtures["simple_bracket.docx"]
    src_sha = sha(src)
    res, out, rep = convert_simple(gen_fixtures, tmp_path, cited={1, 2})
    assert Path(res.output_path).is_file()
    assert sha(src) == src_sha                 # input never modified

    # two citation fields: [1] bracket + (2) paren; ref 3 uncited, left out
    assert len(res.fields) == 2
    assert [(f.text, f.numbers) for f in res.fields] in (
        [("[1]", [1]), ("(2)", [2])], [("(2)", [2]), ("[1]", [1])])
    assert res.uncited == [3] and res.uncited_kept is False
    assert any("left OUT" in n for n in res.notes)

    conv = DocxPackage.open(out)
    fields = find_zotero_fields(conv.document, DOCUMENT_PART)
    assert [f.kind for f in fields] == ["ITEM", "ITEM", "BIBL"]

    f1 = fields[0]
    assert f1.instr.startswith(ZOTERO_ITEM_INSTR_PREFIX)
    assert f1.instr.endswith("} ")
    assert list(f1.payload.keys()) == ["citationID", "properties",
                                       "citationItems", "schema"]
    assert list(f1.payload["properties"].keys()) == [
        "unsorted", "formattedCitation", "plainCitation", "noteIndex"]
    assert f1.payload["properties"]["unsorted"] is False
    ci = f1.payload["citationItems"][0]
    assert list(ci.keys()) == ["id", "uris", "itemData"]
    assert ci["id"] == 1 and ci["itemData"]["id"] == 1     # first doc-local id
    assert ci["uris"] == ["http://zotero.org/users/local/TOKENLOC/items/KEY1"]
    assert ci["itemData"]["DOI"] == "10.1000/j1.2020.1"
    assert ci["itemData"]["issued"] == {"date-parts": [["2020"]]}
    assert itemdata_order_class(ci["itemData"]).startswith("id,type")
    assert len(f1.citation_id) == 8

    # bibliography: heading kept, typed refs gone, ONE entry per cited ref
    assert res.bibliography["written"] is True
    assert res.bibliography["entries"] == 2
    pars = conv.document.xpath("//w:p", namespaces=NS)
    texts = [par_text(p) for p in pars]
    assert "References" in texts
    hidx = texts.index("References")
    assert all("Title three of analysis" not in t for t in texts)   # uncited gone
    assert not any(t.startswith("Smith J, Jones AB.") for t in texts)  # typed gone
    bib_pars = pars[hidx + 1:]
    assert bib_pars[0].find("./w:pPr/w:pStyle", namespaces=NS).get(w("val")) \
        == "Bibliography"
    assert run_pattern(bib_pars[0])[0] == "fld:begin"
    # entries paragraphs then a text-empty end paragraph
    assert run_pattern(bib_pars[-1]) == ["fld:end"]
    assert "".join(t.text or "" for t in bib_pars[-1].xpath(".//w:t",
                                                            namespaces=NS)) == ""
    assert "Title one of the study" in par_text(bib_pars[0])

    # Bibliography style injected (python-docx template lacks it)
    assert res.styles_injected is True
    st = conv.styles.xpath("//w:style[@w:styleId='Bibliography']", namespaces=NS)
    assert st and st[0].find("./w:pPr/w:ind", namespaces=NS).get(w("left")) == "720"

    # prefs part created from scratch with exact sample shape
    assert res.prefs.created_part is True
    assert res.prefs.pids == [2, 3]
    assert res.prefs.chunk_count == 2
    sid, zv, style_id = read_prefs(conv)
    assert style_id == "http://www.zotero.org/styles/elsevier-harvard"
    assert zv == "10.0.6" and sid and len(sid) == 8
    props = conv.custom_props.xpath("//*[local-name()='property']")
    assert [p.get("name") for p in props] == ["ZOTERO_PREF_1", "ZOTERO_PREF_2"]
    vals = [p.xpath("*[local-name()='lpwstr']")[0].text for p in props]
    assert len(vals[0]) == 255

    # only the parts that must change did change
    diff = differing_parts(src, out)
    assert diff <= {"word/document.xml", "word/styles.xml",
                    "docProps/custom.xml", "[Content_Types].xml",
                    "_rels/.rels"}
    assert "word/document.xml" in diff and "docProps/custom.xml" in diff
    # untouched core parts byte-identical
    for part in ("word/settings.xml", "word/fontTable.xml", "word/theme/theme1.xml",
                 "word/webSettings.xml", "docProps/core.xml", "docProps/app.xml"):
        assert part not in diff

    # python-docx (strict OPC reader) still opens the result
    from docx import Document
    Document(str(out))


def test_convert_grouped_single_field_ordered(gen_fixtures, tmp_path):
    res, out, _ = convert_simple(gen_fixtures, tmp_path,
                                 name="grouped_bracket.docx", items=lib5())
    assert len(res.fields) == 2
    by_text = {f.text: f for f in res.fields}
    assert by_text["[2,5]"].numbers == [2, 5]
    assert by_text["[1, 3-5]"].numbers == [1, 3, 4, 5]
    assert by_text["[2,5]"].item_keys == ["KEY2", "KEY5"]      # citation order
    assert by_text["[1, 3-5]"].item_keys == ["KEY1", "KEY3", "KEY4", "KEY5"]
    conv = DocxPackage.open(out)
    fields = [f for f in find_zotero_fields(conv.document, DOCUMENT_PART)
              if f.kind == "ITEM"]
    payloads = {len(f.payload["citationItems"]): f.payload for f in fields}
    p2 = payloads[2]
    assert [ci["uris"][0].rsplit("/", 1)[-1] for ci in p2["citationItems"]] \
        == ["KEY2", "KEY5"]
    p4 = payloads[4]
    assert [ci["uris"][0].rsplit("/", 1)[-1] for ci in p4["citationItems"]] \
        == ["KEY1", "KEY3", "KEY4", "KEY5"]
    # doc-local ids increment across the whole document: 1..6
    ids = sorted(ci["id"] for p in (p2, p4) for ci in p["citationItems"])
    assert ids == [1, 2, 3, 4, 5, 6]
    assert any("grouped citation" in e for e in res.extrapolations)
    assert any("multi-entry ZOTERO_BIBL" in e for e in res.extrapolations)
    assert res.bibliography["entries"] == 5 and res.uncited == []


def test_convert_two_in_one_paragraph(gen_fixtures, tmp_path):
    res, out, _ = convert_simple(gen_fixtures, tmp_path,
                                 name="two_in_paragraph.docx")
    assert len(res.fields) == 2
    assert all(f.par_index == 0 for f in res.fields)
    conv = DocxPackage.open(out)
    fields = find_zotero_fields(conv.document, DOCUMENT_PART)
    zones, _ = zotero_field_zones(conv.document, DOCUMENT_PART)
    pars = conv.document.xpath("//w:p", namespaces=NS)
    pattern = run_pattern(pars[0])
    assert pattern.count("fld:begin") == 2 and pattern.count("fld:end") == 2
    # two disjoint zones, both inside paragraph 0
    assert len(zones) > 0


def test_convert_split_runs_keeps_surrounding_runs(gen_fixtures, tmp_path):
    src = gen_fixtures["split_runs.docx"]
    orig = DocxPackage.open(src)
    orig_pars = orig.document.xpath("//w:p", namespaces=NS)
    orig_texts = [par_text(p) for p in orig_pars[:2]]
    res, out, _ = convert_simple(gen_fixtures, tmp_path, name="split_runs.docx",
                                 items=lib5(), cited={1, 2, 5})
    assert len(res.fields) == 2
    conv = DocxPackage.open(out)
    pars = conv.document.xpath("//w:p", namespaces=NS)
    # text outside the fields equals the original minus the citation spans
    zone, _ = zotero_field_zones(conv.document, DOCUMENT_PART)
    for i in (0, 1):
        outside = "".join(
            t.text or ""
            for t in pars[i].xpath(".//w:t", namespaces=NS)
            if t.xpath("ancestor::w:r[1]", namespaces=NS)[0] not in zone)
        stripped = orig_texts[i]
        for m in (("[2,5]",) if i == 0 else ("(1)",)):
            stripped = stripped.replace(m, "", 1)
        assert outside == stripped
    # exact paragraph round-trip: prefix + preview + suffix (split runs kept
    # their surrounding text: " improve." came from the "] improve." run)
    f_25 = [f for f in res.fields if f.text == "[2,5]"][0]
    f_1 = [f for f in res.fields if f.text == "(1)"][0]
    assert par_text(pars[0]) == "Financial services " + f_25.preview + " improve."
    assert par_text(pars[1]) == "Retail markets" + f_1.preview + " react."


def test_convert_superscript_field_not_superscript(gen_fixtures, tmp_path):
    res, out, _ = convert_simple(gen_fixtures, tmp_path,
                                 name="superscript.docx")
    assert len(res.fields) == 2
    assert any("superscript formatting NOT inherited" in e
               for e in res.extrapolations)
    conv = DocxPackage.open(out)
    zone, fields = zotero_field_zones(conv.document, DOCUMENT_PART)
    for r in zone:
        va = r.xpath("./w:rPr/w:vertAlign", namespaces=NS)
        assert va == [], "field run inherited superscript formatting"


def test_convert_footnote_citation(gen_fixtures, tmp_path):
    src = gen_fixtures["footnote_citation.docx"]
    rep = resolve_local(src, lib3(), cited_numbers={1, 2})
    out = tmp_path / "fn.docx"
    pkg = DocxPackage.open(src)
    res = convert_document(pkg, rep, style="elsevier-harvard",
                           output=str(out), rng=random.Random(1), emit=SILENT)
    assert len(res.fields) == 2
    parts = {f.part for f in res.fields}
    assert parts == {"word/document.xml", "word/footnotes.xml"}
    assert any("footnotes/endnotes" in e for e in res.extrapolations)
    conv = DocxPackage.open(out)
    fn_fields = find_zotero_fields(conv.footnotes, "word/footnotes.xml")
    assert len(fn_fields) == 1 and fn_fields[0].kind == "ITEM"
    assert fn_fields[0].payload["properties"]["noteIndex"] == 0
    # decision g: footnotes part was actually saved (bytes changed)
    diff = differing_parts(src, out)
    assert "word/footnotes.xml" in diff
    assert "word/endnotes.xml" not in diff     # untouched -> byte-identical


def test_convert_partly_converted(gen_fixtures, tmp_path):
    src = gen_fixtures["partly_converted.docx"]
    before = DocxPackage.open(src)
    before_fields = find_zotero_fields(before.document, DOCUMENT_PART)
    rep = resolve_local(src, [])           # empty library, no credentials
    out = tmp_path / "partly.docx"
    pkg = DocxPackage.open(src)
    res = convert_document(pkg, rep, style="elsevier-harvard",
                           output=str(out), rng=random.Random(7), emit=SILENT)

    # ref 4 reused from the existing field; refs 1-3 have no items
    assert len(res.fields) == 1
    f4 = res.fields[0]
    assert f4.text == "[4]" and f4.numbers == [4]
    assert f4.item_keys == ["FIXITEM1"]
    assert f4.adjacent_existing is True            # decision c: reported
    assert res.adjacent_fields == [f4]
    assert len(res.unconverted) == 3
    assert all("unresolved" in u.reason for u in res.unconverted)

    conv = DocxPackage.open(out)
    fields = find_zotero_fields(conv.document, DOCUMENT_PART)
    assert [f.kind for f in fields] == ["ITEM", "ITEM", "ITEM", "BIBL"]

    # the pre-existing fields' RUN SEQUENCES (begin..end) are untouched --
    # the paragraph itself changed because the new [4] field lives in it
    def field_run_xml(f):
        out = [etree.tostring(f.begin_run)]
        r = f.begin_run
        while r is not f.end_run:
            r = r.getnext()
            assert r is not None
            out.append(etree.tostring(r))
        return out

    before_by_id = {f.citation_id: f for f in before_fields}
    after_fields = [f for f in fields if f.citation_id in ("FIXTUREA", "FIXTUREB")]
    assert len(after_fields) == 2
    for f in after_fields:
        assert field_run_xml(f) == field_run_xml(before_by_id[f.citation_id])
    # new field got the next doc-local numeric id (existing fields use 1)
    newf = [f for f in fields
            if f.citation_id not in ("FIXTUREA", "FIXTUREB", None)][0]
    ci = newf.payload["citationItems"][0]
    assert ci["id"] == 2 and ci["itemData"]["id"] == 2
    assert ci["uris"] == ["http://zotero.org/users/local/FIXLIB01/items/FIXITEM1"]
    # old BIBL replaced, typed list gone, exactly one BIBL field
    assert res.bibliography["removed_existing_bibl"] is True
    assert res.bibliography["replaced_typed_list"] is True
    assert sum(1 for f in fields if f.kind == "BIBL") == 1
    text = full_text(conv)
    assert "Title two about things" not in text    # typed entries removed
    # prefs updated IN PLACE: session kept, style changed, no duplicates
    assert res.prefs.created_part is False
    assert res.prefs.style_changed is True
    assert res.prefs.old_style_id == \
        "http://www.zotero.org/styles/harvard-cite-them-right"
    sid, zv, style_id = read_prefs(conv)
    assert sid == "FIXTURE1"                       # session preserved
    assert style_id == "http://www.zotero.org/styles/elsevier-harvard"
    props = conv.custom_props.xpath("//*[local-name()='property']")
    names = [p.get("name") for p in props]
    assert names == [f"ZOTERO_PREF_{i+1}" for i in range(len(names))]
    assert len(set(names)) == len(names)
    pids = [int(p.get("pid")) for p in props]
    assert pids == sorted(pids) and pids[0] >= 2
    # custom.xml part existed -> CT/rels untouched
    diff = differing_parts(src, out)
    assert "[Content_Types].xml" not in diff and "_rels/.rels" not in diff


def test_convert_idempotent_second_run(gen_fixtures, tmp_path):
    res1, out1, rep = convert_simple(gen_fixtures, tmp_path, cited={1, 2})
    out2 = tmp_path / "second.docx"
    pkg2 = DocxPackage.open(out1)
    res2 = convert_document(pkg2, rep, style="elsevier-harvard",
                            output=str(out2), rng=random.Random(99),
                            emit=SILENT)
    assert res2.fields == []                       # nothing new to convert
    assert any("idempotent re-run" in n for n in res2.notes)
    assert res2.bibliography["written"] is False
    # document.xml not even re-serialised: byte-identical to first output
    z1, z2 = zipfile.ZipFile(out1), zipfile.ZipFile(out2)
    assert z1.read("word/document.xml") == z2.read("word/document.xml")
    # full structural signature equal
    sig1 = package_signature(DocxPackage.open(out1))
    sig2 = package_signature(DocxPackage.open(out2))
    assert diff_signatures(sig1, sig2) == []


def test_convert_keep_uncited(gen_fixtures, tmp_path):
    # simple_bracket cites [1] and (2); ref 3 is never cited
    res, out, _ = convert_simple(gen_fixtures, tmp_path, cited={1, 2},
                                 keep_uncited=True)
    assert res.uncited == [3] and res.uncited_kept is True
    assert res.bibliography["entries"] == 3        # cited 1,2 + kept 3
    assert any("uncited" in e and "UNVERIFIED" in e for e in res.extrapolations)
    conv = DocxPackage.open(out)
    bibl = [f for f in find_zotero_fields(conv.document, DOCUMENT_PART)
            if f.kind == "BIBL"][0]
    assert bibl.payload["uncited"] == [
        "http://zotero.org/users/local/TOKENLOC/items/KEY3"]
    assert bibl.payload["omitted"] == [] and bibl.payload["custom"] == []


def test_convert_dry_run_writes_nothing(gen_fixtures, tmp_path):
    src = gen_fixtures["simple_bracket.docx"]
    rep = resolve_local(src, lib3(), cited_numbers={1, 2})
    pkg = DocxPackage.open(src)
    out = tmp_path / "never.docx"
    res = convert_document(pkg, rep, style="elsevier-harvard", dry_run=True,
                           rng=random.Random(3), emit=SILENT)
    assert res.dry_run is True and res.output_path is None
    assert not out.exists()
    assert len(res.fields) == 2                    # computed, not written
    assert any("dry-run" in n for n in res.notes)


def test_convert_refuses_input_as_output(gen_fixtures, tmp_path):
    src = gen_fixtures["simple_bracket.docx"]
    rep = resolve_local(src, lib3(), cited_numbers={1, 2})
    pkg = DocxPackage.open(src)
    with pytest.raises(ValueError, match="refusing to save onto the input"):
        convert_document(pkg, rep, style="elsevier-harvard", output=str(src),
                         rng=random.Random(3), emit=SILENT)


def test_convert_unresolved_refs_left_as_text(gen_fixtures, tmp_path):
    # empty library + no credentials -> nothing resolved -> no fields at all
    src = gen_fixtures["simple_bracket.docx"]
    rep = resolve_local(src, [])
    out = tmp_path / "unres.docx"
    pkg = DocxPackage.open(src)
    res = convert_document(pkg, rep, style="elsevier-harvard", output=str(out),
                           rng=random.Random(3), emit=SILENT)
    assert res.fields == []
    assert len(res.unconverted) == 2
    conv = DocxPackage.open(out)
    text = full_text(conv)
    assert "[1]" in text and "(2)" in text         # raw citations preserved


# ---------------------------------------------------------------------------
# the real document, end to end (mocked HTTP)
# ---------------------------------------------------------------------------

def test_convert_real_vancouver_end_to_end(real_docx, tmp_path):
    src_sha = sha(real_docx)
    pkg0 = DocxPackage.open(real_docx)
    det0 = detect_citations(pkg0, find_reference_list(pkg0))
    cited = set().union(*[c.numbers for c in det0.citations])
    assert cited == set(range(4, 17))

    t = FakeTransport(dead_local() + web_routes(items=[]) + [crossref_route([])])
    rep = resolve_document(DocxPackage.open(real_docx), transport=t,
                           api_key=SECRET_KEY, library_id="12345",
                           cited_numbers=cited, emit=SILENT)
    assert rep.uncited == [1, 2, 3]

    out = tmp_path / "real_vancouver_converted.docx"
    pkg = DocxPackage.open(real_docx)
    res = convert_document(pkg, rep, style="vancouver", output=str(out),
                           rng=random.Random(11), emit=SILENT)
    assert sha(real_docx) == src_sha               # input untouched

    # 12 citation fields, all in document.xml
    assert len(res.fields) == 12
    assert all(f.part == "word/document.xml" for f in res.fields)
    grouped = {tuple(f.numbers): f for f in res.fields if len(f.numbers) > 1}
    assert set(grouped) == {(9, 11), (10, 13)}
    # the reused item keeps its sample-form local URI
    f4 = [f for f in res.fields if f.numbers == [4]][0]
    assert f4.item_keys == ["NU469JIQ"]
    assert f4.adjacent_existing is True
    assert res.adjacent_fields == [f4]             # decision c listing

    conv = DocxPackage.open(out)
    fields = find_zotero_fields(conv.document, DOCUMENT_PART)
    assert sum(1 for f in fields if f.kind == "ITEM") == 13   # 12 new + 1 old
    assert sum(1 for f in fields if f.kind == "BIBL") == 1    # decision d: ONE
    old = [f for f in fields if f.citation_id == "xO3paPuA"][0]
    assert "(Havryk and Nazarova, 2024)" in old.instr.replace(
        '"', "") or "Havryk" in etree.tostring(old.begin_par).decode()
    # doc-local numeric ids: existing field uses 1 -> new ids 2..15, unique
    ids = [ci["id"] for f in fields if f.kind == "ITEM" and f.payload
           for ci in f.payload["citationItems"]]
    assert 1 in ids and len(set(ids)) == len(ids) and max(ids) == 15

    # grouped payloads keep citation order (decision b)
    def titles_for(numbers):
        f = [f for f in fields if f.kind == "ITEM" and f.payload
             and f.citation_items() and
             [ci.get("itemData", {}).get("title") for ci in f.payload
              ["citationItems"]] and
             len(f.payload["citationItems"]) == len(numbers) and
             f.payload["citationItems"][0]["itemData"].get("title")
             == titles[numbers[0]]][0]
        return [ci["itemData"]["title"] for ci in f.payload["citationItems"]]

    titles = {9: "Artificial intelligence in UK financial services: 2024 "
                 "survey results",
              10: "The financial stability implications of artificial "
                  "intelligence",
              11: "Intelligent financial system: how AI is transforming "
                  "finance",
              13: "Artificial Intelligence Risk Management Framework "
                  "(AI RMF 1.0)"}
    assert titles_for([9, 11]) == [titles[9], titles[11]]
    assert titles_for([10, 13]) == [titles[10], titles[13]]

    texts = full_text(conv)
    # uncited refs 1-3: typed text gone AND no preview entry (left OUT)
    assert "Ünvan" not in texts
    assert "Unlocking the full potential" not in texts
    assert "Digital transformation leadership characteristics" not in texts
    assert "References" in texts                   # heading kept

    # bibliography: 13 cited entries in first-citation order + end paragraph
    assert res.bibliography["entries"] == 13
    assert res.bibliography["replaced_typed_list"] is True
    assert res.bibliography["removed_existing_bibl"] is True
    pars = conv.document.xpath("//w:p", namespaces=NS)
    hidx = [i for i, p in enumerate(pars) if par_text(p) == "References"][0]
    first_bib = pars[hidx + 1]
    assert run_pattern(first_bib)[0] == "fld:begin"
    assert first_bib.find("./w:pPr/w:pStyle", namespaces=NS).get(w("val")) == \
        "Bibliography"

    # prefs updated in place: harvard-cite-them-right -> vancouver, session kept
    assert res.prefs.created_part is False and res.prefs.style_changed is True
    sid, zv, style_id = read_prefs(conv)
    assert sid == "88ppEdET" and zv == "10.0.6"
    assert style_id == "http://www.zotero.org/styles/vancouver"

    # styles.xml already had Bibliography -> untouched, byte-identical
    diff = differing_parts(real_docx, out)
    assert "word/styles.xml" not in diff
    # ...and the difference from the sample definition is REPORTED
    assert any("Bibliography style" in n and "kept as-is" in n
               for n in res.notes)
    assert "word/footnotes.xml" not in diff and "word/endnotes.xml" not in diff
    assert "[Content_Types].xml" not in diff and "_rels/.rels" not in diff
    assert diff == {"word/document.xml", "docProps/custom.xml"}

    # second detection pass finds nothing left to convert
    det1 = detect_citations(conv, find_reference_list(conv))
    assert det1.citations == []

    # strict OPC reader still opens it
    from docx import Document
    Document(str(out))


# ---------------------------------------------------------------------------
# structural equivalence with the real Zotero sample (decision h)
# ---------------------------------------------------------------------------

def test_structure_matches_zotero_sample(gen_fixtures, tmp_path, sample_docx):
    res, out, _ = convert_simple(gen_fixtures, tmp_path, cited={1, 2})
    sample_sig = package_signature(DocxPackage.open(sample_docx))
    conv_sig = package_signature(DocxPackage.open(out))

    s_item = [f for f in sample_sig["fields"] if f["kind"] == "ITEM"][0]
    c_item = [f for f in conv_sig["fields"] if f["kind"] == "ITEM"][0]
    assert c_item["instr_form"] == s_item["instr_form"] == \
        " ADDIN ZOTERO_ITEM CSL_CITATION {} "
    assert c_item["payload_key_order"] == s_item["payload_key_order"]
    assert c_item["properties_key_order"] == s_item["properties_key_order"]
    assert c_item["citationItem_key_order"] == s_item["citationItem_key_order"]
    assert c_item["itemData_order_class"] == s_item["itemData_order_class"]
    assert c_item["itemData_id_equals_item_id"] is True
    assert c_item["uris_is_list_of_zotero_uris"] is True
    s_pat = field_slice(s_item["paragraph_run_patterns"][0])
    c_pat = field_slice(c_item["paragraph_run_patterns"][0])
    assert c_pat == s_pat == ["fld:begin", "instr", "fld:separate", "t",
                              "fld:end"]

    s_bibl = [f for f in sample_sig["fields"] if f["kind"] == "BIBL"][0]
    c_bibl = [f for f in conv_sig["fields"] if f["kind"] == "BIBL"][0]
    assert c_bibl["instr_form"] == s_bibl["instr_form"] == \
        " ADDIN ZOTERO_BIBL {} CSL_BIBLIOGRAPHY "
    assert c_bibl["payload_key_order"] == s_bibl["payload_key_order"] == \
        ["uncited", "omitted", "custom"]
    assert c_bibl["paragraph_run_patterns"][0] == \
        s_bibl["paragraph_run_patterns"][0]        # begin/instr/separate + entry
    assert c_bibl["paragraph_run_patterns"][-1] == \
        s_bibl["paragraph_run_patterns"][-1] == ["fld:end"]
    assert c_bibl["begin_par_pPr"] == s_bibl["begin_par_pPr"]
    assert c_bibl["begin_par_pPr"]["style"] == "Bibliography"
    assert c_bibl["end_par_pPr"] is None           # SS4.2 end paragraph
    assert c_bibl["end_par_text_empty"] is True

    # custom properties: same names, fmtid, ascending pids from 2, 255-chunks
    assert [p["name_form"] for p in conv_sig["custom_props"]] == \
        [p["name_form"] for p in sample_sig["custom_props"]] == \
        ["ZOTERO_PREF_N", "ZOTERO_PREF_N"]
    assert [p["pid_relative_order"] for p in conv_sig["custom_props"]] == \
        [p["pid_relative_order"] for p in sample_sig["custom_props"]] == [2, 3]
    assert {p["fmtid"] for p in conv_sig["custom_props"]} == \
        {p["fmtid"] for p in sample_sig["custom_props"]}
    assert conv_sig["custom_props"][0]["value_is_255_or_last"] == 255

    # Bibliography style definition identical to the sample's
    assert conv_sig["styles.Bibliography"] == sample_sig["styles.Bibliography"]

    # package wiring for custom.xml identical
    assert conv_sig["content_type_overrides"] == sample_sig["content_type_overrides"]
    assert "custom-properties" in conv_sig["root_rel_types"]
    assert "custom-properties" in sample_sig["root_rel_types"]
