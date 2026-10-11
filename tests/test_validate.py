"""
Phase 4 tests for zoteroconv.validate: structural checks against the real
Zotero-written sample, the loud mock-key library failure the user asked
for, and tamper cases (duplicate citationID, broken prefs, unconverted
input, typed list coexisting with a BIBL field).
"""

from __future__ import annotations

import json
import random
import re

from lxml import etree

from test_convert import SILENT, convert_simple, lib3, resolve_local
from test_resolve_items import FakeTransport, crossref_route, local_routes

from zoteroconv.docx_io import DocxPackage
from zoteroconv.resolve_items import ZoteroAccess
from zoteroconv.validate import main as validate_main
from zoteroconv.validate import validate_document, validate_file

NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
W = NS["w"]


def local_access(items):
    acc = ZoteroAccess(FakeTransport(local_routes(items) + [crossref_route([])]),
                       zotero_read="local")
    acc.detect()
    return acc


def checks(rep):
    return [f.check for f in rep.failures]


# ---------------------------------------------------------------------------
# the real Zotero-written sample must pass the structural checks
# ---------------------------------------------------------------------------

def test_validate_sample_structural(sample_docx):
    rep = validate_file(sample_docx, check_library=False)
    assert rep.ok, rep.summary_lines()
    assert rep.stats["item_fields"] == 1 and rep.stats["bibl_fields"] == 1
    assert rep.stats["pref_chunks"] == 2
    assert rep.library_checked is False
    assert any("SKIPPED" in ln for ln in rep.summary_lines())


# ---------------------------------------------------------------------------
# converted output: structural + library check against the real (fake-routed)
# library
# ---------------------------------------------------------------------------

def test_validate_converted_with_library_ok(gen_fixtures, tmp_path):
    _res, out, _rep = convert_simple(gen_fixtures, tmp_path, cited={1, 2})
    acc = local_access(lib3())
    rep = validate_document(DocxPackage.open(out), access=acc)
    assert rep.ok, rep.summary_lines()
    assert rep.library_checked is True
    assert rep.stats["keys_referenced"] == 2       # KEY1 + KEY2 (ref3 uncited)


def test_validate_fails_loudly_on_mock_keys(gen_fixtures, tmp_path):
    # fields cite KEY1/KEY2 but the "real" library the validator reads is
    # empty -> every key must fail with the do-NOT-ship wording
    _res, out, _rep = convert_simple(gen_fixtures, tmp_path, cited={1, 2})
    acc = local_access([])
    rep = validate_document(DocxPackage.open(out), access=acc)
    assert not rep.ok
    lib_fails = [f for f in rep.failures if f.check == "library-key"]
    assert {f.message.split("'")[1] for f in lib_fails} == {"KEY1", "KEY2"}
    assert all("NOT FOUND" in f.message for f in lib_fails)
    assert all("mock" in f.message for f in lib_fails)
    assert all("do NOT ship" in f.message for f in lib_fails)


def test_validate_library_check_offline_warns(gen_fixtures, tmp_path):
    from test_resolve_items import dead_local
    _res, out, _rep = convert_simple(gen_fixtures, tmp_path, cited={1, 2})
    acc = ZoteroAccess(FakeTransport(dead_local()), zotero_read="auto")
    acc.detect()
    rep = validate_document(DocxPackage.open(out), access=acc)
    assert rep.ok                                  # structural only
    assert rep.library_checked is False
    assert any("unreachable" in w for w in rep.warnings)


# ---------------------------------------------------------------------------
# tamper cases
# ---------------------------------------------------------------------------

def test_validate_duplicate_citation_id(gen_fixtures, tmp_path):
    _res, out, _rep = convert_simple(gen_fixtures, tmp_path,
                                     name="two_in_paragraph.docx")
    pkg = DocxPackage.open(out)
    instrs = [t for t in pkg.document.xpath("//w:instrText", namespaces=NS)
              if "ZOTERO_ITEM" in (t.text or "")]
    assert len(instrs) == 2
    id0 = json.loads(instrs[0].text.split("CSL_CITATION ", 1)[1]).get(
        "citationID")
    id1 = json.loads(instrs[1].text.split("CSL_CITATION ", 1)[1]).get(
        "citationID")
    assert id0 != id1
    instrs[1].text = instrs[1].text.replace(f'"citationID":"{id1}"',
                                            f'"citationID":"{id0}"')
    from zoteroconv.docx_io import DOCUMENT_PART
    pkg.mark_modified(DOCUMENT_PART)
    tampered = tmp_path / "tampered_dup.docx"
    pkg.save(str(tampered))
    rep = validate_file(tampered, check_library=False)
    assert not rep.ok
    assert "citationid-unique" in checks(rep)


def test_validate_broken_prefs(gen_fixtures, tmp_path):
    _res, out, _rep = convert_simple(gen_fixtures, tmp_path, cited={1, 2})
    pkg = DocxPackage.open(out)
    props = pkg.custom_props.xpath("//*[local-name()='property']")
    victim = [p for p in props if p.get("name") == "ZOTERO_PREF_2"][0]
    victim.getparent().remove(victim)
    from zoteroconv.docx_io import CUSTOM_PROPS_PART
    pkg.mark_modified(CUSTOM_PROPS_PART)
    tampered = tmp_path / "tampered_prefs.docx"
    pkg.save(str(tampered))
    rep = validate_file(tampered, check_library=False)
    assert not rep.ok
    assert "prefs-xml" in checks(rep)          # truncated concatenation


def test_validate_unconverted_input_fails(gen_fixtures):
    rep = validate_file(gen_fixtures["simple_bracket.docx"],
                        check_library=False)
    assert not rep.ok
    assert "no-zotero-fields" in checks(rep)
    assert "prefs-missing" in checks(rep)


def test_validate_typed_list_beside_bibl_fails(gen_fixtures):
    # partly_converted: typed refs 1-4 AND an existing ZOTERO_BIBL field
    rep = validate_file(gen_fixtures["partly_converted.docx"],
                        check_library=False)
    assert not rep.ok
    assert "typed-list-coexists" in checks(rep)


def test_validate_real_vancouver_conversion_ok(real_docx, tmp_path):
    """The full mocked real-document run must pass every structural check;
    the library check sees the fake web library (created NEW* keys are
    absent from a re-read) -- so skip it here; the loud-failure behaviour
    for exactly that situation is covered above."""
    from test_resolve_items import (SECRET_KEY, dead_local, web_routes)
    from zoteroconv.convert import convert_document
    from zoteroconv.detect_citations import detect_citations
    from zoteroconv.parse_reference_list import find_reference_list
    from zoteroconv.resolve_items import resolve_document

    pkg0 = DocxPackage.open(real_docx)
    det0 = detect_citations(pkg0, find_reference_list(pkg0))
    cited = set().union(*[c.numbers for c in det0.citations])
    t = FakeTransport(dead_local() + web_routes(items=[]) + [crossref_route([])])
    rep = resolve_document(DocxPackage.open(real_docx), transport=t,
                           api_key=SECRET_KEY, library_id="12345",
                           cited_numbers=cited, emit=SILENT)
    out = tmp_path / "real_conv.docx"
    res = convert_document(DocxPackage.open(real_docx), rep,
                           style="vancouver", output=str(out),
                           rng=random.Random(11), emit=SILENT)
    assert len(res.fields) == 12
    v = validate_document(DocxPackage.open(out), check_library=False)
    assert v.ok, v.summary_lines()
    assert v.stats["item_fields"] == 13 and v.stats["bibl_fields"] == 1
    assert v.stats["pref_chunks"] == 2
    # mock keys must FAIL the real-library check when the library is
    # re-read as empty (exactly the user's Phase-4 requirement)
    acc = ZoteroAccess(FakeTransport(dead_local() + web_routes(items=[])),
                       api_key=SECRET_KEY, library_id="12345",
                       zotero_read="web")
    acc.detect()
    v2 = validate_document(DocxPackage.open(out), access=acc)
    assert not v2.ok
    assert "library-key" in checks(v2)


# ---------------------------------------------------------------------------
# module entry point
# ---------------------------------------------------------------------------

def test_validate_module_main(sample_docx, tmp_path, capsys):
    assert validate_main([str(sample_docx), "--skip-library-check"]) == 0
    out = capsys.readouterr().out
    assert "OK" in out and "library check: SKIPPED" in out
    assert validate_main([str(tmp_path / "nope.docx")]) == 2
    assert validate_main([]) == 2
    bad = validate_main([str(sample_docx), "--zotero-read", "carrier-pigeon"])
    assert bad == 2
