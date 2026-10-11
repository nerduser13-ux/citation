"""
zoteroconv.validate -- post-conversion validation of a .docx.

Structural checks (no network, derived from docs/zotero_format.md and the
sample fixture):

  * at least one Zotero field; instrText forms exactly
    ' ADDIN ZOTERO_ITEM CSL_CITATION {json} ' / ' ADDIN ZOTERO_BIBL {json}
    CSL_BIBLIOGRAPHY '; JSON payload parses
  * payload key order (citationID, properties, citationItems, schema) and
    (uncited, omitted, custom); properties keys (unsorted,
    formattedCitation, plainCitation, noteIndex)
  * citationIDs are unique across document + footnotes + endnotes
  * every citationItem: integer id, non-empty uris list of
    http://zotero.org/... URIs, itemData dict whose id equals the
    citationItem id and which has a type; schema URL exact
  * exactly one ZOTERO_BIBL field; a typed reference list must not coexist
    with it outside the field (a Refresh would duplicate entries)
  * ZOTERO_PREF_* properties: names ZOTERO_PREF_1..N contiguous, pids
    strictly ascending, every chunk but the last exactly 255 chars, the
    concatenation parses as XML and carries <session id> and <style id>
  * the package still opens with the strict python-docx reader

Library check (network; the user's explicit Phase-4 requirement):

  * every item key used in any citation-field URI (and in the BIBL
    uncited/omitted/custom arrays) must exist in the REAL Zotero library,
    read through the same local-first/web access chain as the resolver.
    A key that is not found fails LOUDLY so that output produced against
    mocked transports can never be mistaken for real output.

Exit codes for ``python -m zoteroconv.validate``: 0 valid, 1 validation
failures, 2 usage/IO error.  Credentials are accepted but never printed.
"""

from __future__ import annotations

import json
import re
import sys
from dataclasses import dataclass, field as dc_field
from pathlib import Path
from typing import Callable

from lxml import etree

from zoteroconv.build_fields import (
    PREF_CHUNK_SIZE,
    read_prefs,
)
from zoteroconv.docx_io import (
    DOCUMENT_PART,
    ENDNOTES_PART,
    FOOTNOTES_PART,
    DocxPackage,
)
from zoteroconv.parse_reference_list import find_reference_list
from zoteroconv.zotero_fields import (
    CITATION_INSTR_HEAD,
    CITATION_SCHEMA_URL,
    BIBL_INSTR_HEAD,
    BIBL_INSTR_TAIL,
    ZoteroField,
    find_zotero_fields,
)

ITEM_KEY_ORDER = ["citationID", "properties", "citationItems", "schema"]
BIBL_KEY_ORDER = ["uncited", "omitted", "custom"]
PROPERTIES_KEY_ORDER = ["unsorted", "formattedCitation", "plainCitation",
                        "noteIndex"]
CITATION_ID_RE = re.compile(r"^[A-Za-z0-9]{8}$")
ZOTERO_URI_RE = re.compile(r"^https?://(?:www\.)?zotero\.org/")


@dataclass
class ValidationFailure:
    check: str                   # short stable id, e.g. 'citationid-unique'
    message: str


@dataclass
class ValidationReport:
    path: str | None
    failures: list[ValidationFailure] = dc_field(default_factory=list)
    warnings: list[str] = dc_field(default_factory=list)
    stats: dict = dc_field(default_factory=dict)
    library_checked: bool = False
    library_keys_seen: int | None = None

    @property
    def ok(self) -> bool:
        return not self.failures

    def fail(self, check: str, message: str) -> None:
        self.failures.append(ValidationFailure(check, message))

    def summary_lines(self) -> list[str]:
        lines = [
            f"validation of {self.path or '<in-memory package>'}: "
            f"{'OK' if self.ok else 'FAILED'}",
            (f"  fields: {self.stats.get('item_fields', 0)} ZOTERO_ITEM, "
             f"{self.stats.get('bibl_fields', 0)} ZOTERO_BIBL; "
             f"pref chunks: {self.stats.get('pref_chunks', 0)}; "
             f"item keys referenced: {self.stats.get('keys_referenced', 0)}"),
            ("  library check: performed against the real Zotero library "
             f"({self.library_keys_seen} keys visible)"
             if self.library_checked else
             "  library check: SKIPPED -- item keys were NOT verified "
             "against a real library"),
        ]
        for w in self.warnings:
            lines.append(f"  WARNING: {w}")
        for f in self.failures:
            lines.append(f"  FAIL [{f.check}]: {f.message}")
        return lines


def _key_from_uri(uri: str) -> str:
    return str(uri).rstrip("/").rsplit("/", 1)[-1]


def _fields_of(pkg: DocxPackage) -> list[ZoteroField]:
    out = []
    for part, tree in ((DOCUMENT_PART, pkg.document),
                       (FOOTNOTES_PART, pkg.footnotes),
                       (ENDNOTES_PART, pkg.endnotes)):
        if tree is None:
            continue
        out.extend(find_zotero_fields(tree, part))
    return out


def _check_fields(pkg: DocxPackage, rep: ValidationReport) -> set[str]:
    """Structural per-field checks; returns the set of item keys referenced."""
    fields = _fields_of(pkg)
    items = [f for f in fields if f.kind == "ITEM"]
    bibls = [f for f in fields if f.kind == "BIBL"]
    rep.stats["item_fields"] = len(items)
    rep.stats["bibl_fields"] = len(bibls)
    if not fields:
        rep.fail("no-zotero-fields",
                 "no Zotero field found in document.xml, footnotes.xml or "
                 "endnotes.xml -- this file was not converted (or the "
                 "conversion wrote nothing)")
    keys: set[str] = set()

    seen_ids: dict[str, str] = {}          # citationID -> first location
    seen_numeric: dict[int, str] = {}
    for f in fields:
        loc = f"{f.part}"
        if f.payload is None:
            rep.fail("payload-json",
                     f"{loc}: Zotero field instrText JSON does not parse: "
                     f"{f.payload_error}")
            continue
        if f.kind == "ITEM":
            if not (f.instr.startswith(CITATION_INSTR_HEAD)
                    and f.instr.endswith(" ")):
                rep.fail("instr-form",
                         f"{loc}: ZOTERO_ITEM instrText does not have the "
                         f"exact form {CITATION_INSTR_HEAD!r}+json+' '")
            if list(f.payload) != ITEM_KEY_ORDER:
                rep.fail("payload-key-order",
                         f"{loc}: payload keys {list(f.payload)} != sample "
                         f"order {ITEM_KEY_ORDER}")
            cid = f.payload.get("citationID")
            if not isinstance(cid, str) or not CITATION_ID_RE.match(cid):
                rep.fail("citationid-form",
                         f"{loc}: citationID {cid!r} is not an 8-character "
                         "alphanumeric string")
            elif cid in seen_ids:
                rep.fail("citationid-unique",
                         f"{loc}: duplicate citationID {cid!r} (first used "
                         f"in {seen_ids[cid]}); Zotero identifies fields by "
                         "this id and duplicates corrupt Refresh")
            else:
                seen_ids[cid] = loc
            props = f.payload.get("properties")
            if not isinstance(props, dict) or list(props) != PROPERTIES_KEY_ORDER:
                rep.fail("properties-keys",
                         f"{loc}: properties missing or wrong key order: "
                         f"{list(props) if isinstance(props, dict) else props}")
            else:
                if not isinstance(props.get("unsorted"), bool):
                    rep.fail("properties-types",
                             f"{loc}: properties.unsorted is not a bool")
                for k in ("formattedCitation", "plainCitation"):
                    if not isinstance(props.get(k), str):
                        rep.fail("properties-types",
                                 f"{loc}: properties.{k} is not a string")
                if not isinstance(props.get("noteIndex"), int):
                    rep.fail("properties-types",
                             f"{loc}: properties.noteIndex is not an integer")
            cites = f.payload.get("citationItems")
            if not isinstance(cites, list) or not cites:
                rep.fail("citationitems",
                         f"{loc}: citationItems must be a non-empty list")
                continue
            for ci in cites:
                if not isinstance(ci, dict) or list(ci) != ["id", "uris",
                                                            "itemData"]:
                    rep.fail("citationitem-key-order",
                             f"{loc}: citationItem keys "
                             f"{list(ci) if isinstance(ci, dict) else ci} != "
                             "['id', 'uris', 'itemData']")
                    continue
                num_id = ci["id"]
                uris = ci["uris"]
                data = ci["itemData"]
                if not isinstance(num_id, int) or isinstance(num_id, bool):
                    rep.fail("citationitem-id",
                             f"{loc}: citationItem id {num_id!r} is not an "
                             "integer")
                else:
                    if num_id in seen_numeric:
                        rep.warnings.append(
                            f"{loc}: numeric citationItem id {num_id} also "
                            f"used in {seen_numeric[num_id]} (Zotero's own "
                            "uniqueness rule for these ids is UNVERIFIED)")
                    else:
                        seen_numeric[num_id] = loc
                if not isinstance(uris, list) or not uris or \
                        not all(isinstance(u, str) and ZOTERO_URI_RE.match(u)
                                for u in uris):
                    rep.fail("citationitem-uris",
                             f"{loc}: uris {uris!r} must be a non-empty list "
                             "of http://zotero.org/... URIs")
                else:
                    keys.update(_key_from_uri(u) for u in uris)
                if not isinstance(data, dict) or "type" not in data:
                    rep.fail("itemdata",
                             f"{loc}: itemData must be an object with a "
                             f"'type' key, got {type(data).__name__}")
                elif data.get("id") != num_id:
                    rep.fail("itemdata-id",
                             f"{loc}: itemData.id {data.get('id')!r} != "
                             f"citationItem id {num_id!r} (sample always "
                             "carries the same numeric id in both places)")
            if f.payload.get("schema") != CITATION_SCHEMA_URL:
                rep.fail("schema-url",
                         f"{loc}: schema {f.payload.get('schema')!r} != "
                         f"{CITATION_SCHEMA_URL!r}")
        else:                                   # BIBL
            if not (f.instr.startswith(BIBL_INSTR_HEAD)
                    and f.instr.endswith(BIBL_INSTR_TAIL)):
                rep.fail("instr-form",
                         f"{loc}: ZOTERO_BIBL instrText does not have the "
                         "exact sample form")
            if list(f.payload) != BIBL_KEY_ORDER:
                rep.fail("payload-key-order",
                         f"{loc}: BIBL payload keys {list(f.payload)} != "
                         f"{BIBL_KEY_ORDER}")
            for arr in BIBL_KEY_ORDER:
                v = f.payload.get(arr)
                if not isinstance(v, list):
                    rep.fail("bibl-arrays",
                             f"{loc}: BIBL {arr!r} must be a list")
                else:
                    for u in v:
                        if isinstance(u, str) and ZOTERO_URI_RE.match(u):
                            keys.add(_key_from_uri(u))
                        elif isinstance(u, dict) and u.get("id"):
                            keys.add(str(u["id"]))
                        else:
                            rep.warnings.append(
                                f"{loc}: BIBL {arr!r} entry {u!r} has an "
                                "UNVERIFIED shape (the sample only shows "
                                "empty arrays)")
    if len(bibls) > 1:
        rep.fail("single-bibl",
                 f"{len(bibls)} ZOTERO_BIBL fields found; a document must "
                 "carry exactly one (a Refresh would duplicate the "
                 "bibliography)")

    # typed reference list must not coexist with the bibliography field
    if bibls:
        rl = find_reference_list(pkg)
        if rl is not None and rl.references:
            bibl = bibls[0]
            inside = set()
            if bibl.part == rl.part:
                body = bibl.begin_par.getparent() if bibl.begin_par is not None \
                    else None
                if body is not None and bibl.begin_par is not None \
                        and bibl.end_par is not None \
                        and bibl.end_par.getparent() is body:
                    i0 = body.index(bibl.begin_par)
                    i1 = body.index(bibl.end_par)
                    inside = set(list(body)[i0:i1 + 1])
            outside = [r for r in rl.references
                       if not any(el in inside for el in r.elements)]
            if outside:
                rep.fail("typed-list-coexists",
                         f"ZOTERO_BIBL field present but {len(outside)} typed "
                         "reference-list entr(ies) remain outside it "
                         "(example: "
                         f"{outside[0].raw[:60]!r}); Zotero > Refresh would "
                         "leave a duplicated bibliography")
    rep.stats["keys_referenced"] = len(keys)
    return keys


def _check_prefs(pkg: DocxPackage, rep: ValidationReport) -> None:
    root = pkg.custom_props
    if root is None:
        rep.fail("prefs-missing",
                 "docProps/custom.xml is absent; Zotero stores its session "
                 "and style in ZOTERO_PREF_* properties and needs them on "
                 "Refresh")
        return
    props = []
    for p in root.xpath("//*[local-name()='property']"):
        name = p.get("name") or ""
        m = re.fullmatch(r"ZOTERO_PREF_(\d+)", name)
        if m:
            v = p.xpath("*[local-name()='lpwstr']")
            pid = p.get("pid")
            props.append((int(m.group(1)),
                          int(pid) if pid is not None and pid.lstrip(
                              "-").isdigit() else None,
                          v[0].text if v and v[0].text else ""))
    rep.stats["pref_chunks"] = len(props)
    if not props:
        rep.fail("prefs-missing",
                 "no ZOTERO_PREF_* property in docProps/custom.xml")
        return
    nums = sorted(n for n, _, _ in props)
    if nums != list(range(1, len(nums) + 1)):
        rep.fail("prefs-names",
                 f"ZOTERO_PREF_* names are not contiguous 1..N: got {nums}")
    pids = [pid for _, pid, _ in sorted(props)]
    if any(pid is None for pid in pids):
        rep.fail("prefs-pids", "a ZOTERO_PREF_* property has no numeric pid")
    elif pids != sorted(pids) or len(set(pids)) != len(pids):
        rep.fail("prefs-pids",
                 f"ZOTERO_PREF_* pids must be unique and ascending, got {pids}")
    ordered = sorted(props)
    for i, (_, _, value) in enumerate(ordered):
        last = i == len(ordered) - 1
        if not last and len(value) != PREF_CHUNK_SIZE:
            rep.fail("prefs-chunking",
                     f"ZOTERO_PREF_{i + 1} holds {len(value)} chars; every "
                     f"chunk but the last must be exactly {PREF_CHUNK_SIZE}")
        if last and len(value) > PREF_CHUNK_SIZE:
            rep.fail("prefs-chunking",
                     f"last chunk holds {len(value)} > {PREF_CHUNK_SIZE} chars")
        if not value:
            rep.fail("prefs-chunking", f"ZOTERO_PREF_{i + 1} is empty")
    joined = "".join(v for _, _, v in ordered)
    try:
        data = etree.fromstring(joined.encode("utf-8"))
    except etree.XMLSyntaxError as exc:
        rep.fail("prefs-xml",
                 f"concatenated ZOTERO_PREF_* chunks do not parse as XML: {exc}")
        return
    if not data.tag.endswith("}data") and data.tag != "data":
        rep.fail("prefs-xml", f"prefs root element is {data.tag!r}, not <data>")
        return
    if not data.xpath("//session/@id"):
        rep.fail("prefs-session", "prefs carry no <session id=...>")
    style = data.xpath("//style/@id")
    if not style or not str(style[0]).strip():
        rep.fail("prefs-style", "prefs carry no <style id=...>")
    sid, zv, style_id = read_prefs(pkg)
    rep.stats["prefs_style_id"] = style_id
    rep.stats["prefs_session_id"] = sid
    rep.stats["prefs_zotero_version"] = zv


def _check_bibliography_style(pkg: DocxPackage, rep: ValidationReport) -> None:
    bibls = [f for f in _fields_of(pkg) if f.kind == "BIBL"]
    if not bibls or bibls[0].part != DOCUMENT_PART:
        return
    styles = pkg.styles
    if styles is None:
        rep.fail("styles-part", "word/styles.xml is missing")
        return
    if not styles.xpath("//w:style[@w:styleId='Bibliography']",
                        namespaces={"w": "http://schemas.openxmlformats.org/"
                                    "wordprocessingml/2006/main"}):
        rep.fail("bibliography-style",
                 "the bibliography field's entry paragraphs use pStyle "
                 "'Bibliography' but word/styles.xml does not define it")


def _check_opc(pkg: DocxPackage, rep: ValidationReport) -> None:
    if not pkg.path:
        rep.warnings.append("in-memory package: strict OPC reader check skipped")
        return
    try:
        from docx import Document
        Document(str(pkg.path))
    except Exception as exc:
        rep.fail("opc-reader",
                 f"python-docx (strict OPC reader) cannot open the file: "
                 f"{type(exc).__name__}: {exc}")


def _check_library(keys: set[str], access, rep: ValidationReport) -> None:
    """Fail loudly for every key that the REAL library does not contain."""
    if access is None:
        rep.warnings.append(
            "no Zotero access configured: library check skipped")
        return
    try:
        items = access.fetch_items_csl()
    except Exception as exc:
        rep.warnings.append(
            f"library check skipped: fetching the real library failed "
            f"({type(exc).__name__}: {exc})")
        return
    if access.read_mode == "offline":
        rep.warnings.append(
            "Zotero unreachable (neither local API nor Web API): library "
            "check skipped -- item keys NOT verified")
        return
    real = set()
    for it in items:
        uri = str(it.get("id") or "")
        key = it.get("key")
        if key:
            real.add(str(key))
        if uri:
            real.add(_key_from_uri(uri))
    rep.library_checked = True
    rep.library_keys_seen = len(real)
    missing = sorted(keys - real)
    for k in missing:
        rep.fail(
            "library-key",
            f"item key {k!r} is used by a citation field but was NOT FOUND "
            f"in the real Zotero library (read mode: {access.read_mode}, "
            f"{len(real)} keys visible). This output was produced against a "
            "mocked/placeholder library or the item was deleted since -- "
            "do NOT ship this file; re-run zoteroconv against the real "
            "library.")


def validate_document(pkg: DocxPackage, *, access=None,
                      check_library: bool | None = None,
                      ) -> ValidationReport:
    """Validate an opened package.  ``access`` is a detected
    ``resolve_items.ZoteroAccess`` used for the library-key check; pass
    ``check_library=False`` to skip that check entirely (structural
    validation only)."""
    rep = ValidationReport(path=str(pkg.path) if pkg.path else None)
    keys = _check_fields(pkg, rep)
    _check_prefs(pkg, rep)
    _check_bibliography_style(pkg, rep)
    _check_opc(pkg, rep)
    if check_library is not False and keys:
        _check_library(keys, access, rep)
    elif check_library is not False and not keys:
        rep.warnings.append("no item keys referenced: library check not needed")
    return rep


def validate_file(path: str | Path, *, zotero_read: str = "auto",
                  api_key: str | None = None, library_id: str | None = None,
                  library_type: str | None = None,
                  transport=None,
                  check_library: bool | None = None,
                  ) -> ValidationReport:
    """Open ``path`` and validate it.  Credentials come from arguments or
    the environment (ZT_API_KEY, ZT_LIBRARY_ID, ZT_LIBRARY_TYPE) and are
    never printed."""
    import os

    from zoteroconv.resolve_items import RequestsTransport, ZoteroAccess
    pkg = DocxPackage.open(str(path))
    access = None
    if check_library is not False:
        env = os.environ
        access = ZoteroAccess(
            transport or RequestsTransport(),
            api_key=api_key or env.get("ZT_API_KEY") or None,
            library_id=library_id or env.get("ZT_LIBRARY_ID") or None,
            library_type=(library_type or env.get("ZT_LIBRARY_TYPE")
                          or "user"),
            zotero_read=zotero_read,
        )
        access.detect()
    return validate_document(pkg, access=access, check_library=check_library)


def main(argv: list[str] | None = None) -> int:
    argv = list(sys.argv[1:] if argv is None else argv)
    if not argv or argv[0] in ("-h", "--help"):
        print(__doc__.split("Exit codes")[0].strip())
        print("\nUsage: python -m zoteroconv.validate FILE.docx "
              "[--skip-library-check] [--zotero-read auto|local|web] "
              "[--api-key KEY] [--library-id ID] [--library-type user|group]")
        return 2 if not argv else 0
    path = argv[0]
    check_library = True
    if "--skip-library-check" in argv:
        check_library = False
    zotero_read = "auto"
    if "--zotero-read" in argv:
        zotero_read = argv[argv.index("--zotero-read") + 1]
    if zotero_read not in ("auto", "local", "web"):
        print(f"error: --zotero-read must be auto, local or web, got "
              f"{zotero_read!r}")
        return 2
    kw: dict = {}
    for flag, name in (("--api-key", "api_key"),
                       ("--library-id", "library_id"),
                       ("--library-type", "library_type")):
        if flag in argv:
            kw[name] = argv[argv.index(flag) + 1]
    if not Path(path).is_file():
        print(f"error: no such file: {path}")
        return 2
    try:
        rep = validate_file(path, zotero_read=zotero_read,
                            check_library=check_library, **kw)
    except Exception as exc:
        print(f"error: cannot validate {path}: {type(exc).__name__}: {exc}")
        return 2
    for ln in rep.summary_lines():
        print(ln)
    return 0 if rep.ok else 1


def main_entry() -> None:          # console-script wrapper
    raise SystemExit(main())


if __name__ == "__main__":
    main_entry()
