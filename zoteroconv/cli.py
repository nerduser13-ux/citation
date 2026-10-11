"""
zoteroconv.cli -- the Typer command line:

    zoteroconv INPUT.docx --style vancouver --output OUT.docx
               [--collection NAME] [--source-style auto] [--dry-run]
               [--report report.md] [--zotero-read auto|local|web]
               [--library-type user|group] [--library-id ID]
               [--api-key KEY] [--max-new-items 20] [--locale en-US]
               [--keep-uncited] [--limit-paragraphs N] [--no-validate]
               [--skip-library-check] [--no-cache] [--yes]

Credentials come from flags or the environment (ZT_API_KEY, ZT_LIBRARY_ID,
ZT_LIBRARY_TYPE, CROSSREF_MAILTO) and are never printed or written to the
report.  Exit codes: 0 success, 1 validation of the written file failed,
2 usage/input error.

The transport factory ``_make_transport`` exists so tests can inject a
mocked HTTP layer; production code always talks to the real endpoints.
"""

from __future__ import annotations

import os
import sys
from pathlib import Path
from typing import Optional

import typer

from zoteroconv.convert import convert_document
from zoteroconv.detect_citations import DOCUMENT_PART, detect_citations
from zoteroconv.docx_io import DocxPackage
from zoteroconv.http_cache import HttpCache
from zoteroconv.parse_reference_list import find_reference_list
from zoteroconv.report import build_report_markdown
from zoteroconv.resolve_items import RequestsTransport, resolve_document
from zoteroconv.styles import StyleError, resolve_style_id
from zoteroconv.validate import validate_document

app = typer.Typer(add_completion=False, pretty_exceptions_show_locals=False,
                  context_settings={"help_option_names": ["-h", "--help"]})

DEFAULT_CACHE_DB = Path.home() / ".cache" / "zoteroconv" / "http.sqlite3"


def _make_transport() -> RequestsTransport:
    """Production transport; tests monkeypatch this single seam."""
    return RequestsTransport()


def _die(message: str, code: int = 2) -> None:
    typer.echo(f"error: {message}", err=True)
    raise typer.Exit(code)


def _confirm_factory(assume_yes: bool):
    def confirm(question: str) -> bool:
        if assume_yes:
            typer.echo(f"auto-confirmed (--yes): {question}")
            return True
        if not sys.stdin.isatty():
            typer.echo(f"NOT confirmed (non-interactive stdin, no --yes): "
                       f"{question}")
            return False
        return typer.confirm(question)
    return confirm


@app.command()
def convert(
    input: Path = typer.Argument(
        ..., exists=True, dir_okay=False, readable=True,
        help="input .docx with numeric citations (never modified)"),
    style: str = typer.Option(
        ..., "--style",
        help="target CSL style: short name (apa, mla, harvard, ieee, "
             "vancouver, chicago, nature), any Zotero style slug, or a "
             "full http://www.zotero.org/styles/... URL"),
    output: Optional[Path] = typer.Option(
        None, "--output", "-o",
        help="converted .docx to write (required unless --dry-run)"),
    collection: Optional[str] = typer.Option(
        None, "--collection",
        help="Zotero collection for created items "
             "(default: 'Imported - <document name>')"),
    source_style: str = typer.Option(
        "auto", "--source-style",
        help="citation form in the source document; only 'auto' is "
             "supported (the detector handles \\[n], (n), superscripts, "
             "ranges and groups automatically)"),
    dry_run: bool = typer.Option(
        False, "--dry-run",
        help="resolve and plan only: never write to Zotero, never write "
             "the output file"),
    report: Optional[Path] = typer.Option(
        None, "--report",
        help="write a Markdown conversion report to this path"),
    zotero_read: str = typer.Option(
        "auto", "--zotero-read",
        help="where to read the library: auto (local API first, then Web "
             "API), local, or web"),
    library_type: Optional[str] = typer.Option(
        None, "--library-type", help="user or group (env ZT_LIBRARY_TYPE)"),
    library_id: Optional[str] = typer.Option(
        None, "--library-id", help="Zotero library id (env ZT_LIBRARY_ID)"),
    api_key: Optional[str] = typer.Option(
        None, "--api-key",
        help="Zotero Web API key (env ZT_API_KEY); never printed or logged"),
    max_new_items: int = typer.Option(
        20, "--max-new-items",
        help="refuse to create more than this many new Zotero items "
             "without confirmation"),
    locale: str = typer.Option(
        "en-US", "--locale", help="CSL locale written into the prefs"),
    keep_uncited: bool = typer.Option(
        False, "--keep-uncited",
        help="put never-cited references into the ZOTERO_BIBL 'uncited' "
             "array instead of leaving them out (format UNVERIFIED: the "
             "sample fixture's array is empty)"),
    limit_paragraphs: Optional[int] = typer.Option(
        None, "--limit-paragraphs",
        help="EXCERPT MODE: convert only citations in the first N body "
             "paragraphs of document.xml (footnotes always included); "
             "everything else stays raw text and is reported"),
    no_validate: bool = typer.Option(
        False, "--no-validate",
        help="do not validate the written output file"),
    skip_library_check: bool = typer.Option(
        False, "--skip-library-check",
        help="validation: skip the check that every cited item key exists "
             "in the REAL Zotero library"),
    no_cache: bool = typer.Option(
        False, "--no-cache", help="disable the SQLite HTTP cache"),
    cache_db: Optional[Path] = typer.Option(
        None, "--cache-db",
        help=f"SQLite cache path (default {DEFAULT_CACHE_DB})"),
    yes: bool = typer.Option(
        False, "--yes",
        help="answer confirmation prompts (e.g. above --max-new-items) "
             "with yes; without it, non-interactive runs refuse"),
) -> None:
    """Convert numeric .docx citations into live Zotero Word fields."""
    # ---- option validation ------------------------------------------------
    try:
        style_id = resolve_style_id(style)
    except StyleError as exc:
        _die(str(exc))
    if source_style != "auto":
        _die(f"--source-style {source_style!r} is not supported; only "
             "'auto' (the detector recognises bracketed, parenthesised and "
             "superscript numeric citations, ranges and groups)")
    if zotero_read not in ("auto", "local", "web"):
        _die(f"--zotero-read must be auto, local or web, got {zotero_read!r}")
    if library_type is not None and library_type not in ("user", "group"):
        _die(f"--library-type must be user or group, got {library_type!r}")
    if not dry_run and output is None:
        _die("--output is required unless --dry-run is given")
    if output is not None and output.resolve() == input.resolve():
        _die("refusing to use the input file as --output (the input is "
             "never modified)")
    if limit_paragraphs is not None and limit_paragraphs < 0:
        _die(f"--limit-paragraphs must be >= 0, got {limit_paragraphs}")
    if max_new_items < 0:
        _die(f"--max-new-items must be >= 0, got {max_new_items}")

    env = os.environ
    api_key = api_key or env.get("ZT_API_KEY") or None
    library_id = library_id or env.get("ZT_LIBRARY_ID") or None
    library_type = library_type or env.get("ZT_LIBRARY_TYPE") or None
    crossref_mailto = env.get("CROSSREF_MAILTO") or None

    # ---- open + detect (for cited numbers / excerpt mode) ------------------
    try:
        probe = DocxPackage.open(str(input))
    except Exception as exc:
        _die(f"cannot open {input}: {type(exc).__name__}: {exc}")
    rl = find_reference_list(probe)
    det = detect_citations(probe, rl)
    cited: set[int] = set()
    for m in det.citations:
        if (limit_paragraphs is not None and m.part == DOCUMENT_PART
                and m.par_index >= limit_paragraphs):
            continue
        cited.update(m.numbers)
    if not det.citations:
        typer.echo("warning: no numeric citations detected in the body; "
                   "only the bibliography/prefs would change")
    if limit_paragraphs is not None:
        scope = (f"body paragraphs #0..#{limit_paragraphs - 1}"
                 if limit_paragraphs > 0 else "NOTHING (limit is 0)")
        typer.echo(f"excerpt mode: converting citations in {scope} only "
                   f"({len(cited)} reference number(s) in scope)")

    # ---- resolve against Zotero -------------------------------------------
    transport = _make_transport()
    cache = HttpCache(None if no_cache else (cache_db or DEFAULT_CACHE_DB))
    try:
        resolution = resolve_document(
            DocxPackage.open(str(input)),
            threshold=90.0,
            max_new_items=max_new_items,
            dry_run=dry_run,
            collection_name=collection,
            confirm=_confirm_factory(yes),
            transport=transport,
            cache=cache,
            api_key=api_key,
            library_id=library_id,
            library_type=library_type,
            zotero_read=zotero_read,
            crossref_mailto=crossref_mailto,
            emit=typer.echo,
            cited_numbers=cited,
            keep_uncited=keep_uncited,
        )
    except Exception as exc:
        _die(f"resolution failed: {type(exc).__name__}: {exc}")

    # ---- convert -----------------------------------------------------------
    try:
        result = convert_document(
            DocxPackage.open(str(input)), resolution,
            style=style, locale=locale, keep_uncited=keep_uncited,
            dry_run=dry_run,
            output=str(output) if output is not None else None,
            limit_paragraphs=limit_paragraphs,
            emit=typer.echo,
        )
    except StyleError as exc:                # pragma: no cover - checked above
        _die(str(exc))
    except ValueError as exc:
        _die(str(exc))

    for ln in result.summary_lines():
        typer.echo(ln)
    for u in result.unconverted:
        typer.echo(f"  NOT converted: {u.part} par #{u.par_index} "
                   f"{u.text!r} refs {u.numbers}: {u.reason}")

    # ---- validate the written file ----------------------------------------
    validation = None
    exit_code = 0
    if not dry_run and output is not None and not no_validate:
        from zoteroconv.resolve_items import ZoteroAccess
        access = ZoteroAccess(
            transport, api_key=api_key, library_id=library_id,
            library_type=library_type or "user", zotero_read=zotero_read)
        try:
            access.detect()
        except Exception as exc:
            typer.echo(f"warning: Zotero access detection for validation "
                       f"failed ({type(exc).__name__}: {exc})")
        validation = validate_document(
            DocxPackage.open(str(output)), access=access,
            check_library=not skip_library_check)
        for ln in validation.summary_lines():
            typer.echo(ln)
        if not validation.ok:
            exit_code = 1

    # ---- report ------------------------------------------------------------
    if report is not None:
        md = build_report_markdown(
            input_path=str(input),
            output_path=None if dry_run else str(output),
            style_query=style,
            resolution=resolution,
            conversion=result,
            validation=validation,
            options={
                "dry_run": dry_run,
                "limit_paragraphs": limit_paragraphs,
                "keep_uncited": keep_uncited,
                "max_new_items": max_new_items,
                "library_type": library_type,
                "library_id": bool(library_id),
                "api_key_provided": bool(api_key),
                "collection": collection,
            },
        )
        report.parent.mkdir(parents=True, exist_ok=True)
        report.write_text(md, encoding="utf-8")
        typer.echo(f"report written: {report}")

    cache.close()
    raise typer.Exit(exit_code)


def main() -> None:
    app()


if __name__ == "__main__":
    main()
