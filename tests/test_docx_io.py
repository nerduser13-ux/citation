"""DocxPackage: part exposure, byte-identical round-trip, input protection."""

from __future__ import annotations

import zipfile

import pytest
from lxml import etree

from zoteroconv.docx_io import (
    CUSTOM_PROPS_PART,
    DOCUMENT_PART,
    DocxPackage,
)
from zoteroconv.zotero_fields import NS


def _members(path) -> dict[str, bytes]:
    with zipfile.ZipFile(path, "r") as zf:
        return {name: zf.read(name) for name in zf.namelist()}


def _member_order(path) -> list[str]:
    with zipfile.ZipFile(path, "r") as zf:
        return zf.namelist()


def test_sample_parts_exposed(sample_docx):
    pkg = DocxPackage.open(sample_docx)
    assert DOCUMENT_PART in pkg.part_names()
    assert CUSTOM_PROPS_PART in pkg.part_names()
    assert pkg.document is not None
    assert pkg.custom_props is not None
    # the sample has no footnotes/endnotes parts (Phase 0, SECTION A)
    assert pkg.footnotes is None
    assert pkg.endnotes is None


def test_real_doc_has_footnotes_endnotes_numbering(real_docx):
    pkg = DocxPackage.open(real_docx)
    assert len(pkg.part_names()) == 33
    assert pkg.footnotes is not None
    assert pkg.endnotes is not None
    assert pkg.numbering is not None
    assert pkg.styles is not None


@pytest.mark.parametrize("which", ["sample_docx", "real_docx"])
def test_save_unmodified_is_byte_identical(which, request, tmp_path):
    path = request.getfixturevalue(which)
    pkg = DocxPackage.open(path)
    out = tmp_path / "roundtrip.docx"
    pkg.save(out)
    before, after = _members(path), _members(out)
    assert set(before) == set(after)
    for name, data in before.items():
        assert after[name] == data, f"part {name} changed on unmodified save"
    assert _member_order(path) == _member_order(out)


def test_save_modified_document_keeps_other_parts_identical(sample_docx, tmp_path):
    pkg = DocxPackage.open(sample_docx)
    first_t = pkg.document.xpath("//w:t", namespaces=NS)[0]
    original = first_t.text
    first_t.text = "MODIFIED-BY-TEST"
    pkg.mark_modified(DOCUMENT_PART)
    out = tmp_path / "modified.docx"
    pkg.save(out)

    before, after = _members(sample_docx), _members(out)
    assert set(before) == set(after)
    changed = [n for n in before if before[n] != after[n]]
    # document.xml is re-serialised: it is NOT byte-identical (spec note)
    assert changed == [DOCUMENT_PART]

    reopened = DocxPackage.open(out)
    texts = [t.text for t in reopened.document.xpath("//w:t", namespaces=NS)]
    assert "MODIFIED-BY-TEST" in texts
    assert original not in texts


def test_save_refuses_to_overwrite_input(sample_docx):
    pkg = DocxPackage.open(sample_docx)
    with pytest.raises(ValueError, match="refusing to save onto the input"):
        pkg.save(pkg.path)


def test_mark_modified_missing_part_raises(sample_docx):
    pkg = DocxPackage.open(sample_docx)
    with pytest.raises(KeyError):
        pkg.mark_modified("word/footnotes.xml")


def test_partly_converted_prefs_part_readable(gen_fixtures):
    pkg = DocxPackage.open(gen_fixtures["partly_converted.docx"])
    root = pkg.custom_props
    assert root is not None
    props = root.xpath(
        "//*[local-name()='property']",
    )
    names = [p.get("name") for p in props]
    pids = [p.get("pid") for p in props]
    assert names == ["ZOTERO_PREF_1", "ZOTERO_PREF_2"]
    assert pids == ["2", "3"]
    values = [
        p.xpath("*[local-name()='lpwstr']")[0].text or "" for p in props
    ]
    assert len(values[0]) == 255  # fixed chunk size from docs/zotero_format.md 2.4
    joined = "".join(values)
    data = etree.fromstring(joined.encode("utf-8"))
    style = data.xpath("//style")[0]
    assert style.get("id") == "http://www.zotero.org/styles/harvard-cite-them-right"
    assert style.get("locale") == "en-US"
