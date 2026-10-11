"""
zoteroconv.report -- Markdown conversion report (--report report.md).

Everything the phases promised to surface ends up here: resolution
outcomes per reference, created items and their tags, low-confidence
placeholders and candidates that need a human, converted fields,
unconverted / ambiguous matches, adjacency to pre-existing Zotero fields,
uncited references, prefs changes, flagged extrapolations, validation
results -- and the instruction to open the result in Word and click
Zotero > Refresh.  Credentials are never included (only whether they were
provided and through which channel).
"""

from __future__ import annotations

from datetime import date
from typing import Any


def _trunc(s: str, n: int = 70) -> str:
    s = " ".join(str(s).split())
    return s if len(s) <= n else s[:n - 1] + "\u2026"


def _mask(value: str | None) -> str:
    return "provided (never logged)" if value else "not provided"


def build_report_markdown(*, input_path: str, output_path: str | None,
                          style_query: str, resolution: Any,
                          conversion: Any, validation: Any = None,
                          options: dict | None = None,
                          ) -> str:
    options = options or {}
    res, conv = resolution, conversion
    L: list[str] = []
    A = L.append

    A("# zoteroconv conversion report")
    A("")
    A(f"- date: {date.today().isoformat()}")
    A(f"- input: `{input_path}` (never modified)")
    A(f"- output: `{output_path}`" if output_path else
      "- output: none (dry run -- no file was written)")
    A(f"- requested style: `{style_query}` -> `{conv.style_id}` "
      f"(locale `{conv.locale}`)")
    A(f"- dry run: {bool(options.get('dry_run'))}")
    if options.get("limit_paragraphs") is not None:
        A(f"- EXCERPT MODE: only citations in body paragraphs #0.."
          f"#{options['limit_paragraphs'] - 1} of document.xml were "
          "converted; later citations remain raw text and the bibliography "
          "only covers the excerpt")
    if options.get("keep_uncited"):
        A("- `--keep-uncited`: uncited references were added to the "
          "ZOTERO_BIBL `uncited` array (format UNVERIFIED -- the sample "
          "fixture's array is empty)")
    A("")

    A("## Zotero access")
    A("")
    A(f"- {res.mode_line}")
    A(f"- write channel: {res.write_channel or 'none'} "
      f"(local API accepts writes: {res.local_writes})")
    A(f"- API key: {_mask('x' if options.get('api_key_provided') else None)}")
    A(f"- library: type `{options.get('library_type') or 'user'}`, id "
      f"{_mask('x' if options.get('library_id') else None)}")
    A(f"- match threshold: {res.threshold}; max new items: "
      f"{options.get('max_new_items')}")
    if res.collection_name:
        A(f"- collection: `{res.collection_name}` "
          f"(key `{res.collection_key or 'n/a'}`)")
    A("")

    A("## References: resolution outcomes")
    A("")
    A("| # | status | method | item key | title (truncated) |")
    A("|---|--------|--------|----------|-------------------|")
    for r in sorted(res.resolutions, key=lambda r: r.ref_number):
        item = r.item
        key = item.key if item is not None else "-"
        title = ""
        if item is not None:
            title = _trunc(item.csl.get("title") or "")
        else:
            title = _trunc((r.parsed_csl or {}).get("title") or r.raw)
        A(f"| {r.ref_number} | {r.status} | {r.method or '-'} | `{key}` "
          f"| {title} |")
    A("")
    if res.created_keys:
        A(f"Created {len(res.created_keys)} new Zotero item(s), each tagged "
          f"`zoteroconv-created`: "
          + ", ".join(f"`{k}`" for k in res.created_keys))
        A("")
    review = [r for r in res.resolutions
              if r.status.endswith("low-confidence") or r.suggestion]
    if review:
        A("### Needs review (low confidence -- tagged `zoteroconv-review`)")
        A("")
        for r in review:
            A(f"- ref {r.ref_number} ({r.status}): "
              + "; ".join(r.messages) if r.messages else
              f"- ref {r.ref_number} ({r.status})")
        A("")
    ambiguous_matches = [r for r in res.resolutions
                         if len(r.candidates) > 1]
    if ambiguous_matches:
        A("### Multiple candidates above threshold -- NOT auto-picked")
        A("")
        for r in ambiguous_matches:
            A(f"- ref {r.ref_number}: "
              + "; ".join(f"{c.source} {c.key or c.doi or c.title or '?'} "
                          f"score {c.score:.1f}" for c in r.candidates))
        A("")
    if res.notes:
        A("### Resolver notes")
        A("")
        for n in res.notes:
            A(f"- {n}")
        A("")

    A("## Conversion")
    A("")
    A(f"- preview engine: {conv.renderer_engine} (preview texts are "
      "PROVISIONAL: Zotero overwrites them on Refresh)")
    A(f"- citation fields written: {len(conv.fields)}")
    A(f"- ambiguous matches NOT converted: {conv.ambiguous_count}; "
      f"excluded (Table/Figure/years/...): {conv.excluded_count}")
    A(f"- uncited references: {conv.uncited} "
      f"({'kept via uncited array (UNVERIFIED)' if conv.uncited_kept else 'left OUT of the bibliography'})")
    A("")
    if conv.fields:
        A("| part | par | raw text | refs | citationID | item keys | adjacent to existing field |")
        A("|------|-----|----------|------|------------|-----------|----------------------------|")
        for f in conv.fields:
            A(f"| {f.part} | {f.par_index} | `{f.text}` | {f.numbers} "
              f"| `{f.citation_id}` | {', '.join('`%s`' % k for k in f.item_keys)} "
              f"| {'YES (separate field, decision c)' if f.adjacent_existing else '-'} |")
        A("")
    if conv.unconverted:
        A("### Left as raw text (NOT converted)")
        A("")
        for u in conv.unconverted:
            A(f"- {u.part} par #{u.par_index} `{u.text}` refs {u.numbers}: "
              f"{u.reason}")
        A("")
    A(f"- bibliography: {conv.bibliography}")
    if conv.prefs is not None:
        p = conv.prefs
        A(f"- prefs: {p.chunk_count} chunk(s), pids {p.pids}, part created: "
          f"{p.created_part}, style changed: {p.style_changed}"
          + (f" (`{p.old_style_id}` -> `{conv.style_id}`)"
             if p.style_changed and p.old_style_id else ""))
    if conv.styles_injected:
        A("- word/styles.xml: `Bibliography` style injected (sample-verbatim "
          "definition minus w:rsid)")
    A("")
    if conv.extrapolations:
        A("### EXTRAPOLATIONS beyond the sample fixture (flagged, UNVERIFIED)")
        A("")
        for e in conv.extrapolations:
            A(f"- {e}")
        A("")
    if conv.renderer_notes:
        A("### Preview renderer notes")
        A("")
        for n in conv.renderer_notes:
            A(f"- {n}")
        A("")
    if conv.notes:
        A("### Conversion notes")
        A("")
        for n in conv.notes:
            A(f"- {n}")
        A("")

    if validation is not None:
        A("## Validation of the written file")
        A("")
        for ln in validation.summary_lines():
            A(f"- {ln.strip()}")
        A("")

    A("## What to do next")
    A("")
    A("**Open in Word with the Zotero plugin and click Zotero > Refresh.**")
    A("")
    A("Check that (1) every converted citation re-renders in the chosen "
      "style, (2) the bibliography re-renders with one entry per cited "
      "reference, (3) grouped citations like `[2,5]` stay one field, and "
      "(4) no other content changed. If anything fails, send back the exact "
      "Zotero/Word error text and the converted file.")
    A("")
    A("### UNVERIFIED in this toolchain (cannot be tested without Word)")
    A("")
    A("- Word + Zotero plugin acceptance of the generated fields (structure "
      "is reproduced byte-for-byte from a real Zotero-written sample).")
    A("- Grouped multi-item fields, multi-entry ZOTERO_BIBL paragraph "
      "layout, footnote `noteIndex`, and non-empty `uncited` arrays -- the "
      "sample fixture demonstrates only the single-item, single-entry, "
      "empty-array cases.")
    A("- Preview texts (citeproc-py): provisional; Zotero recomputes them "
      "on Refresh.")
    return "\n".join(L) + "\n"
