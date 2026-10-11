"""
Phase 3 -- conversion orchestrator: raw numeric citations -> live Zotero
fields; typed list + old ZOTERO_BIBL -> ONE new ZOTERO_BIBL field; prefs
written/updated in place.  The input file is never modified (DocxPackage
refuses to save onto it); output goes to a separate path.
"""

from __future__ import annotations

import random
from dataclasses import dataclass, field as dc_field
from typing import Callable

from lxml import etree

from zoteroconv.build_fields import (
    BIBLIOGRAPHY_STYLE_XML,
    SAMPLE_ITEMDATA_KEYS,
    PrefsChange,
    bibliography_paragraphs,
    bibl_instr,
    char_map,
    citation_field_runs,
    citation_instr,
    ensure_bibliography_style,
    make_bibl_payload,
    make_citation_payload,
    new_citation_id,
    replace_span_with_runs,
    rpr_without_superscript,
    write_prefs,
)
from zoteroconv.detect_citations import (
    DOCUMENT_PART,
    FOOTNOTES_PART,
    ENDNOTES_PART,
    detect_citations,
)
from zoteroconv.docx_io import DocxPackage
from zoteroconv.parse_reference_list import find_reference_list
from zoteroconv.preview import PreviewRenderer
from zoteroconv.resolve_items import ResolutionReport
from zoteroconv.styles import resolve_style_id


@dataclass
class FieldConversion:
    part: str
    par_index: int
    text: str                    # raw matched text, e.g. '[2,5]'
    numbers: list[int]
    citation_id: str
    item_keys: list[str]
    preview: str
    kind: str
    adjacent_existing: bool = False


@dataclass
class Unconverted:
    part: str
    par_index: int
    text: str
    numbers: list[int]
    reason: str


@dataclass
class ConversionResult:
    style_id: str
    locale: str
    dry_run: bool
    output_path: str | None
    fields: list[FieldConversion] = dc_field(default_factory=list)
    unconverted: list[Unconverted] = dc_field(default_factory=list)
    ambiguous_count: int = 0
    excluded_count: int = 0
    bibliography: dict = dc_field(default_factory=dict)
    uncited: list[int] = dc_field(default_factory=list)
    uncited_kept: bool = False
    prefs: PrefsChange | None = None
    styles_injected: bool = False
    renderer_engine: str = ""
    renderer_notes: list[str] = dc_field(default_factory=list)
    extrapolations: list[str] = dc_field(default_factory=list)
    notes: list[str] = dc_field(default_factory=list)

    @property
    def adjacent_fields(self) -> list[FieldConversion]:
        return [f for f in self.fields if f.adjacent_existing]

    def summary_lines(self) -> list[str]:
        lines = [
            f"style: {self.style_id} (locale {self.locale})",
            f"preview engine: {self.renderer_engine}",
            f"citation fields written: {len(self.fields)}",
            f"unconverted matches: {len(self.unconverted)}",
            f"ambiguous (not converted): {self.ambiguous_count}; "
            f"excluded: {self.excluded_count}",
            f"adjacent to existing Zotero fields (separate field, decision c): "
            f"{len(self.adjacent_fields)}",
            f"bibliography: {self.bibliography}",
            f"uncited references: {self.uncited} "
            f"({'KEPT via uncited array -- UNVERIFIED format' if self.uncited_kept else 'left OUT by default'})",
        ]
        if self.prefs is not None:
            lines.append(
                f"prefs: chunks={self.prefs.chunk_count} pids={self.prefs.pids} "
                f"created_part={self.prefs.created_part} "
                f"style_changed={self.prefs.style_changed}"
            )
        for n in self.renderer_notes:
            lines.append(f"preview note: {n}")
        for e in self.extrapolations:
            lines.append(f"EXTRAPOLATION: {e}")
        for n in self.notes:
            lines.append(f"note: {n}")
        lines.append("dry-run: no file written" if self.dry_run
                     else f"output: {self.output_path}")
        return lines


def _zone_ranges(pt) -> list[tuple[int, int]]:
    ranges = []
    i = 0
    n = len(pt.chars)
    while i < n:
        if pt.chars[i].field_zone:
            j = i
            while j < n and pt.chars[j].field_zone:
                j += 1
            ranges.append((i, j))
            i = j
        else:
            i += 1
    return ranges


def _adjacent_to_existing(pt, m) -> bool:
    """True when only whitespace (or nothing) separates the raw match from an
    existing Zotero field zone in the same paragraph (decision c report)."""
    for zs, ze in _zone_ranges(pt):
        if ze <= m.start:
            gap = pt.text[ze:m.start]
        elif zs >= m.end:
            gap = pt.text[m.end:zs]
        else:
            continue
        if gap.strip() == "":
            return True
    return False


def _remove_par_range(begin_par, end_par) -> None:
    body = begin_par.getparent()
    if end_par.getparent() is not body:
        for p in (begin_par, end_par):
            pp = p.getparent()
            if pp is not None:
                pp.remove(p)
        return
    i0, i1 = body.index(begin_par), body.index(end_par)
    for el in list(body)[i0:i1 + 1]:
        body.remove(el)


def convert_document(
    pkg: DocxPackage,
    resolution: ResolutionReport,
    *,
    style: str,
    locale: str = "en-US",
    keep_uncited: bool = False,
    dry_run: bool = False,
    output: str | None = None,
    renderer: PreviewRenderer | None = None,
    rng: random.Random | None = None,
    limit_paragraphs: int | None = None,
    emit: Callable[[str], None] = print,
) -> ConversionResult:
    style_id = resolve_style_id(style)
    rng = rng or random.Random()

    rl = find_reference_list(pkg)
    det = detect_citations(pkg, rl)
    renderer = renderer or PreviewRenderer(style_id, locale)

    result = ConversionResult(
        style_id=style_id, locale=locale or "en-US", dry_run=dry_run,
        output_path=None,
        ambiguous_count=len(det.ambiguous), excluded_count=len(det.excluded),
        renderer_engine=renderer.engine, renderer_notes=list(renderer.notes),
    )
    result.notes.extend(det.notes)

    item_by_ref = {r.ref_number: r.item
                   for r in resolution.resolutions if r.item is not None}

    # ---- plan clusters -------------------------------------------------------
    clusters: list[tuple[int, object, list]] = []   # (inline_index, match, items)
    for m in det.citations:
        if (limit_paragraphs is not None and m.part == DOCUMENT_PART
                and m.par_index >= limit_paragraphs):
            result.unconverted.append(Unconverted(
                m.part, m.par_index, m.text, m.numbers,
                f"skipped: paragraph #{m.par_index} is beyond "
                f"--limit-paragraphs {limit_paragraphs} (excerpt mode)",
            ))
            continue
        missing = [n for n in m.numbers if n not in item_by_ref]
        if missing:
            result.unconverted.append(Unconverted(
                m.part, m.par_index, m.text, m.numbers,
                f"reference(s) {missing} unresolved (no Zotero item); "
                "left as raw text",
            ))
            continue
        clusters.append((len(clusters), m, [item_by_ref[n] for n in m.numbers]))

    # ---- adjacency flags (on the ORIGINAL text, before any mutation) ---------
    adj_flags = {}
    pts_by_part = {part: {pt.index: pt for pt in pts}
                   for part, pts in det.paragraph_texts.items()}
    for ci, m, _items in clusters:
        pt = pts_by_part.get(m.part, {}).get(m.par_index)
        adj_flags[ci] = _adjacent_to_existing(pt, m) if pt else False

    # ---- document-local ids ---------------------------------------------------
    used_numeric: set[int] = set()
    used_cids: set[str] = set()
    for f in det.existing_fields:
        if f.kind != "ITEM":
            continue
        if f.citation_id:
            used_cids.add(f.citation_id)
        for ci in f.citation_items():
            for cand in (ci.get("id"), (ci.get("itemData") or {}).get("id")):
                if isinstance(cand, int):
                    used_numeric.add(cand)
    next_numeric = max(used_numeric, default=0) + 1

    # ---- bibliography item order: first-citation order, then kept uncited ----
    bib_numbers: list[int] = []
    seen: set[int] = set()
    for _ci, m, _items in clusters:
        for n in m.numbers:
            if n not in seen:
                seen.add(n)
                bib_numbers.append(n)
    kept_uncited: list[int] = []
    if keep_uncited:
        for n in (resolution.uncited or []):
            if n in item_by_ref and n not in seen:
                seen.add(n)
                kept_uncited.append(n)
                bib_numbers.append(n)
    result.uncited = list(resolution.uncited or [])
    result.uncited_kept = bool(kept_uncited)
    bib_items = [item_by_ref[n].csl for n in bib_numbers]

    inlines, bib_entries = renderer.render(
        [[it.csl for it in items] for _ci, _m, items in clusters], bib_items)

    # ---- write citation fields (last-to-first per paragraph, decision a) -----
    parts_touched: set[str] = set()
    work = sorted(clusters, key=lambda c: (c[1].part, c[1].par_index,
                                           -c[1].start))
    for ci, m, items in work:
        par = m.paragraph
        text, chars = char_map(par)
        if text[m.start:m.end] != m.text:
            result.unconverted.append(Unconverted(
                m.part, m.par_index, m.text, m.numbers,
                "span text no longer matches (paragraph changed); left as-is",
            ))
            continue
        rpr = rpr_without_superscript(chars[m.start]["run"])
        entries = []
        for it in items:
            entries.append((next_numeric, it.uri, it.csl))
            next_numeric += 1
        cid = new_citation_id(rng, used_cids)
        payload = make_citation_payload(cid, entries, formatted=inlines[ci])
        runs = citation_field_runs(citation_instr(payload), inlines[ci], rpr)
        replace_span_with_runs(par, m.start, m.end, runs)
        parts_touched.add(m.part)
        result.fields.append(FieldConversion(
            part=m.part, par_index=m.par_index, text=m.text,
            numbers=list(m.numbers), citation_id=cid,
            item_keys=[it.key for it in items], preview=inlines[ci],
            kind=m.kind, adjacent_existing=adj_flags.get(ci, False),
        ))

    # ---- bibliography replacement (decisions b/d + Phase-1 decision 2) -------
    bibl_fields = [f for f in det.existing_fields
                   if f.kind == "BIBL" and f.part == DOCUMENT_PART]
    typed_refs = list(rl.references) if rl is not None else []
    do_bib = bool(typed_refs) or bool(clusters and rl is not None) \
        or bool(bibl_fields and clusters)
    bib_info: dict = {
        "replaced_typed_list": False, "removed_existing_bibl": False,
        "entries": 0, "written": False, "uncited_uris": len(kept_uncited),
    }
    if do_bib and rl is not None and rl.heading is not None and not bib_entries:
        # never trade a real reference list for an empty field (data loss)
        result.notes.append(
            "no bibliography entries could be rendered (no resolved items); "
            "the existing reference list was left UNTOUCHED")
    elif do_bib and rl is not None and rl.heading is not None:
        result.styles_injected = ensure_bibliography_style(pkg)
        if not result.styles_injected:
            from zoteroconv.structure_diff import style_definition_signature
            have = style_definition_signature(pkg.styles, "Bibliography")
            want = style_definition_signature(
                etree.fromstring(BIBLIOGRAPHY_STYLE_XML), "Bibliography")
            if have is not None and have != want:
                result.notes.append(
                    "the document already defines a Bibliography style whose "
                    "formatting differs from the sample's; it is kept as-is "
                    "(existing style definitions are never rewritten)")
        if typed_refs:
            for el in rl.region_elements:
                parent = el.getparent()
                if parent is not None:
                    parent.remove(el)
            bib_info["replaced_typed_list"] = True
            parts_touched.add(rl.part)
        for f in bibl_fields:
            _remove_par_range(f.begin_par, f.end_par)
            bib_info["removed_existing_bibl"] = True
            parts_touched.add(f.part)
        payload = make_bibl_payload(
            [item_by_ref[n].uri for n in kept_uncited])
        pars = bibliography_paragraphs(bibl_instr(payload), bib_entries)
        body = rl.heading.getparent()
        idx = body.index(rl.heading) + 1
        for i, p in enumerate(pars):
            body.insert(idx + i, p)
        parts_touched.add(rl.part)
        bib_info.update(written=True, entries=len(bib_entries))
    elif do_bib:
        result.notes.append(
            "citations were converted but no References heading was found; "
            "bibliography field NOT written"
        )
    elif bibl_fields:
        result.notes.append(
            "bibliography left untouched: already a single ZOTERO_BIBL field "
            "and nothing new to convert (idempotent re-run)"
        )
    result.bibliography = bib_info

    # ---- prefs (decision e) ----------------------------------------------------
    result.prefs = write_prefs(pkg, style_id, locale or None, rng=rng)
    result.notes.extend(result.prefs.notes)

    # ---- uncited report ---------------------------------------------------------
    if result.uncited and not keep_uncited:
        result.notes.append(
            f"references never cited numerically {result.uncited}: left OUT "
            "of the bibliography (use --keep-uncited to add them to the "
            "ZOTERO_BIBL uncited array)"
        )

    # ---- extrapolation flags (decision f) ----------------------------------------
    ex = result.extrapolations
    ex.append("rsid/paraId/w14 attributes are not written (Word session artifacts)")
    if any(len(items) > 1 for _ci, _m, items in clusters):
        ex.append("grouped citation -> one field with several citationItems "
                  "(sample shows single-item fields only, SS6.1)")
    if len(bib_entries) > 1:
        ex.append("multi-entry ZOTERO_BIBL: one paragraph per entry between "
                  "separate and end; end fldChar in a final empty paragraph "
                  "(sample shows one entry, SS6.4)")
    if any(m.part in (FOOTNOTES_PART, ENDNOTES_PART)
           for _ci, m, _items in clusters):
        ex.append("field inside footnotes/endnotes with noteIndex 0 "
                  "(not shown by the fixtures, SS6.3/SS6.8)")
    if kept_uncited:
        ex.append('non-empty "uncited" array in the ZOTERO_BIBL payload -- '
                  "the sample only shows an empty array (UNVERIFIED format)")
    extra_keys: set[str] = set()
    for _ci, _m, items in clusters:
        for it in items:
            extra_keys |= (set(it.csl.keys()) - SAMPLE_ITEMDATA_KEYS)
    if extra_keys:
        ex.append("itemData keys beyond the sample's set: "
                  + ", ".join(sorted(extra_keys)))
    if any(m.kind == "superscript" for _ci, m, _items in clusters):
        ex.append("superscript formatting NOT inherited by the field runs "
                  "(decision a; sample fields carry no superscript)")

    # ---- save --------------------------------------------------------------------
    for part in parts_touched:
        pkg.mark_modified(part)
    if dry_run:
        result.notes.append(
            "dry-run: document modified IN MEMORY only; no file written")
        emit("zoteroconv: dry-run -- "
             f"{len(result.fields)} field(s) would be written")
    else:
        if output is None:
            raise ValueError("output path is required unless dry_run=True")
        result.output_path = pkg.save(output)
        emit(f"zoteroconv: wrote {len(result.fields)} citation field(s) + "
             f"bibliography ({len(bib_entries)} entries) -> {result.output_path}")
    return result
