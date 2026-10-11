# zoteroconv

Convert numeric (Vancouver-style) citations in a `.docx` into **live Zotero
Word fields**, so the document behaves exactly like one written with the
Zotero Word plugin:

* every raw citation (`[1]`, `[2,5]`, `[1, 3-5]`, `(2)`, superscript `1`)
  becomes an `ADDIN ZOTERO_ITEM CSL_CITATION` field carrying the full
  CSL-JSON `itemData`, the item URI and the numeric id — a grouped
  citation becomes **one** field with several `citationItems`;
* the typed reference list (and any pre-existing `ZOTERO_BIBL` field) is
  replaced by a **single** `ADDIN ZOTERO_BIBL ... CSL_BIBLIOGRAPHY` field
  at the reference-list position, heading preserved;
* `ZOTERO_PREF_*` document properties are written/updated in place
  (255-char chunks, ascending pids) with your chosen style and locale.

**Open the result in Word with the Zotero plugin and click
Zotero > Refresh** — Zotero re-renders every citation and the bibliography
in the chosen style from the data embedded in the fields.

The field/payload/prefs structure is reproduced byte-for-byte from
`tests/fixtures/zotero_sample.docx`, a real document written by the Zotero
7 Word plugin (see `docs/zotero_format.md`). The input file is **never**
modified; all work happens on in-memory copies written to `--output`.

## Installation

Python 3.11+.

```bash
python3 -m venv .venv
.venv/bin/pip install -r requirements.txt
.venv/bin/pip install -e .          # provides the `zoteroconv` command
```

Preview texts (the grey field results Word shows before a Refresh) are
rendered with `citeproc-py` if the CSL style files are available locally;
put them in `./csl/` or point `ZOTEROCONV_CSL_DIR` at a directory, e.g.
`elsevier-harvard.csl`, `nlm-citation-sequence.csl` (the preview file for
`vancouver`) and friends from
<https://github.com/citation-style-language/styles>. Without them,
zoteroconv falls back to clearly-marked PROVISIONAL preview text — Zotero
overwrites all previews on Refresh anyway.

## Usage

```
zoteroconv INPUT.docx --style vancouver --output OUT.docx
           [--collection NAME] [--source-style auto] [--dry-run]
           [--report report.md] [--zotero-read auto|local|web]
           [--library-type user|group] [--library-id ID] [--api-key KEY]
           [--max-new-items 20] [--locale en-US]
           [--keep-uncited] [--limit-paragraphs N]
           [--no-validate] [--skip-library-check] [--no-cache] [--yes]
```

* `--style` — `apa`, `mla`, `harvard`, `ieee`, `vancouver`, `chicago`,
  `nature`, any CSL slug (`elsevier-harvard`,
  `harvard-cite-them-right`, ...) or a full
  `http://www.zotero.org/styles/...` URL. Invalid values fail with the
  list of valid options.
* `--zotero-read` — where the library is read: `auto` (local Zotero API at
  `http://localhost:23119` first, then Web API), `local`, `web`. CSL-JSON
  is requested (Accept header, then `format=csljson` fallback).
* **Creating items needs the Web API**: set `--api-key` + `--library-id`
  (or `ZT_API_KEY`, `ZT_LIBRARY_ID`, `ZT_LIBRARY_TYPE`). Without them
  zoteroconv stops before any write and says so; `--dry-run` never writes
  anywhere. Credentials are never printed, logged, or put in the report.
* `--max-new-items N` (default 20) — creating more than N items asks for
  confirmation (`--yes` to auto-confirm; non-interactive runs refuse
  without it).
* New items are tagged `zoteroconv-created`; low-confidence placeholders
  are tagged `zoteroconv-review`. Reference metadata is never fabricated:
  CrossRef/PubMed results are only auto-accepted above very strict
  agreement thresholds, everything else stays a visible placeholder.
* `--keep-uncited` — never-cited references are left OUT of the
  bibliography by default; this flag puts their URIs into the
  `ZOTERO_BIBL` `uncited` array instead (format UNVERIFIED — the sample
  fixture's array is empty).
* `--limit-paragraphs N` — **excerpt mode**: convert only citations in the
  first N body paragraphs of `document.xml` (footnotes always included).
  Everything else stays raw text and is reported; the bibliography then
  covers only the excerpt. Useful to test a short document in Word before
  the full run.
* `--report report.md` — Markdown report: per-reference resolution
  outcomes, created items, candidates needing review, every converted and
  unconverted match, adjacency to pre-existing Zotero fields, uncited
  references, prefs/style changes, flagged extrapolations and the
  validation result.
* `--source-style` — only `auto` is supported; the detector handles
  bracketed, parenthesised and superscript numeric citations, ranges and
  groups, skips table/figure/equation labels, years, dosages and
  enumerations, and leaves doubtful matches **ambiguous** (unconverted,
  reported).

Existing Zotero fields are never modified: they are skipped, mapped back
to typed reference numbers (DOI first, then title) so their items are not
created twice, and a raw `[n]` next to an existing field becomes a
separate new field (listed in the report).

### Exit codes

`0` success · `1` the written file failed validation · `2` usage/input
error (bad style, missing output, input == output, unreadable file).

## Validation

Every real run automatically validates the written file; disable with
`--no-validate`. Standalone:

```
zoteroconv-validate OUT.docx [--skip-library-check]
python -m zoteroconv.validate OUT.docx [--zotero-read auto|local|web]
```

Checks (from the spec and the sample fixture): instrText forms and JSON
payload key order, **unique citationIDs** across document + footnotes +
endnotes, citationItem structure (`id`/`uris`/`itemData`, matching numeric
ids, `http://zotero.org/...` URIs), exactly one `ZOTERO_BIBL`, no typed
reference list coexisting with it, `Bibliography` style defined,
`ZOTERO_PREF_*` chunking (255-char chunks, contiguous names, ascending
pids) with a concatenation that parses and carries session + style, and
the strict python-docx OPC reader.

**Library check (the loud one):** every item key used in any citation
field must exist in the **real** Zotero library, re-read through the same
local-first/web chain. A key that is not found fails with
`NOT FOUND in the real Zotero library ... do NOT ship this file` — output
produced against mocked transports can never be mistaken for real output.
Skip only with `--skip-library-check` (structural validation still runs;
the skip is printed).

## Development

```bash
.venv/bin/python -m pytest -q             # 129 tests, all HTTP mocked
.venv/bin/python -m pytest --run-live -q  # + live local-API test (needs Zotero running)
.venv/bin/python scripts/inspect_docx.py FILE.docx      # read-only OOXML inspector
.venv/bin/python scripts/structural_diff.py A.docx B.docx [--only fields,styles]
```

Module map: `docx_io` (zip/OOXML package, byte-identical untouched parts)
· `parse_reference_list` · `detect_citations` (char map across runs) ·
`zotero_fields` · `http_cache` (SQLite) · `resolve_items` (local/web
access, match cascade DOI→PMID→ISBN→fuzzy, CrossRef/PubMed enrichment,
creation with tags) · `styles` · `build_fields` (sample-verbatim field/
bibliography/prefs builders) · `preview` (citeproc-py, provisional
fallback) · `convert` (orchestrator) · `structure_diff` · `validate` ·
`report` · `cli`.

### Known limits / UNVERIFIED

* Word + Zotero plugin acceptance of generated fields cannot be tested
  here — the structure is sample-verbatim, but the final proof is
  **Zotero > Refresh** in Word.
* The sample demonstrates single-item fields, a single-entry bibliography
  and an empty `uncited` array only; grouped fields, multi-entry
  bibliographies, footnote `noteIndex` and non-empty `uncited` arrays are
  flagged as **EXTRAPOLATIONS** in every report.
* Preview texts are provisional (citeproc-py); Zotero recomputes them.
* `w:rsid`/`paraId` session artifacts are deliberately not written.
