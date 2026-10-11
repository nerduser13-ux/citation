"""
Phase 2 -- Zotero access + reference resolution.

Reads the Zotero library (local connector API first, Web API fallback) as
CSL-JSON, resolves every parsed reference to a library item with the cascade

    DOI exact  ->  PMID exact  ->  ISBN exact  ->  fuzzy (rapidfuzz >= 90)

references without identifiers get a CrossRef (bibliographic string) or
PubMed (PMID) *suggestion* which is auto-accepted only when title similarity
is > 95 AND the year AND the first-author surname match; weaker suggestions
are reported as low-confidence candidates and the parsed item is kept as a
visible placeholder.  Nothing is ever fabricated.

Unresolved references are created in the collection ``Imported - <doc name>``
(check-before-create, dedup-check before every create, --max-new-items with a
confirmation hook).  Writes require the Web API with ZT_API_KEY and
ZT_LIBRARY_ID unless a runtime probe shows the local API accepts writes; the
pipeline stops before any write when credentials are missing and never writes
in dry-run mode.  Credentials come only from parameters/env and are never
printed.
"""

from __future__ import annotations

import hashlib
import os
import re
import zlib
from dataclasses import dataclass, field as dc_field
from pathlib import Path
from typing import Any, Callable

from rapidfuzz import fuzz

from zoteroconv.docx_io import DOCUMENT_PART, DocxPackage
from zoteroconv.http_cache import HttpCache
from zoteroconv.parse_reference_list import (
    Reference,
    ReferenceList,
    find_reference_list,
    map_existing_fields_to_references,
    norm_doi,
    norm_title,
)
from zoteroconv.zotero_fields import find_zotero_fields

LOCAL_BASE = "http://localhost:23119/api/users/0"
WEB_BASE = "https://api.zotero.org"
CSL_JSON_ACCEPT = "application/vnd.citationstyles.csl+json"
CROSSREF_WORKS_URL = "https://api.crossref.org/works"
PUBMED_ESUMMARY_URL = (
    "https://eutils.ncbi.nlm.nih.gov/entrez/eutils/esummary.fcgi"
)
TOOL_URL = "https://github.com/nerduser13-ux/citation"
DEFAULT_THRESHOLD = 90.0
SUGGEST_TITLE_ACCEPT = 95.0     # decision 5: strictly above 95
SUGGEST_AUTHOR_ACCEPT = 90.0
DEFAULT_MAX_NEW_ITEMS = 20      # per original spec


# ---------------------------------------------------------------------------
# transport (all HTTP goes through here so tests can mock everything)
# ---------------------------------------------------------------------------

@dataclass
class HttpResponse:
    status: int
    json: Any = None
    text: str = ""
    headers: dict = dc_field(default_factory=dict)


class Transport:
    """Minimal HTTP transport interface."""

    def request(
        self, method: str, url: str, *, params: dict | None = None,
        json: Any = None, headers: dict | None = None, timeout: float = 10.0,
    ) -> HttpResponse:
        raise NotImplementedError


class RequestsTransport(Transport):
    """Real HTTP via requests (only used outside tests / in live runs)."""

    def __init__(self, session=None):
        import requests  # lazy: unit tests never need it

        self._requests = requests
        self.session = session or requests.Session()

    def request(self, method, url, *, params=None, json=None, headers=None,
                timeout=10.0) -> HttpResponse:
        r = self.session.request(
            method, url, params=params, json=json, headers=headers,
            timeout=timeout,
        )
        try:
            payload = r.json()
        except ValueError:
            payload = None
        return HttpResponse(
            status=r.status_code, json=payload, text=r.text,
            headers=dict(r.headers),
        )


# ---------------------------------------------------------------------------
# Zotero access: local-first reads, runtime write probe, web writes
# ---------------------------------------------------------------------------

class ZoteroAccess:
    def __init__(
        self,
        transport: Transport,
        *,
        api_key: str | None = None,
        library_id: str | None = None,
        library_type: str = "user",
        zotero_read: str = "auto",       # 'auto' | 'local' | 'web'
        timeout: float = 5.0,
        local_base: str = LOCAL_BASE,
        web_base: str = WEB_BASE,
    ):
        self.transport = transport
        self.api_key = api_key
        self.library_id = library_id
        self.library_type = library_type or "user"
        self.zotero_read = zotero_read
        self.timeout = timeout
        self.local_base = local_base.rstrip("/")
        self.web_base = web_base.rstrip("/")

        self.local_reachable: bool = False
        self.local_write_status: int | None = None
        self.local_writes: bool | None = None
        self.read_mode: str = "offline"        # 'local' | 'web' | 'offline'
        self.write_channel: str | None = None  # 'local' | 'web' | None
        self.local_uri_prefix: str | None = None
        self.notes: list[str] = []

    # -- probing -------------------------------------------------------------

    @property
    def web_credentials(self) -> bool:
        return bool(self.api_key and self.library_id)

    def library_path(self) -> str:
        kind = "groups" if self.library_type == "group" else "users"
        return f"/{kind}/{self.library_id}"

    def probe_local(self) -> bool:
        if self.zotero_read == "web":
            return False
        try:
            resp = self.transport.request(
                "GET", self.local_base + "/items",
                params={"limit": 1},
                headers={"Accept": CSL_JSON_ACCEPT},
                timeout=self.timeout,
            )
        except Exception:
            self.local_reachable = False
            return False
        # any HTTP answer at all means Zotero is running locally
        self.local_reachable = resp.status < 500
        if not self.local_reachable:
            self.notes.append(
                f"local Zotero answered HTTP {resp.status}; treated as unreachable"
            )
        return self.local_reachable

    def probe_local_writes(self) -> bool | None:
        """POST an intentionally invalid empty item: 405/501/403/404 => the
        local API refuses writes; 4xx validation errors => POST is accepted
        (writes work); 2xx must not happen with an empty body but is treated
        as 'writes work' with a warning."""
        try:
            resp = self.transport.request(
                "POST", self.local_base + "/items", json=[{}],
                headers={"Accept": CSL_JSON_ACCEPT}, timeout=self.timeout,
            )
        except Exception:
            self.local_writes = None
            return None
        self.local_write_status = resp.status
        if resp.status in (405, 501, 403, 404):
            self.local_writes = False
        elif resp.status in (400, 422):
            self.local_writes = True
        elif 200 <= resp.status < 300:
            self.local_writes = True
            self.notes.append(
                "WARNING: local write probe returned "
                f"{resp.status} for an empty item body"
            )
        else:
            self.local_writes = False
            self.notes.append(
                f"local write probe returned HTTP {resp.status}; "
                "treated as 'writes refused'"
            )
        return self.local_writes

    def detect(self) -> None:
        self.probe_local()
        if self.local_reachable:
            self.probe_local_writes()
        if self.zotero_read == "web":
            self.read_mode = "web" if self.web_credentials else "offline"
            if not self.web_credentials:
                self.notes.append(
                    "--zotero-read web but ZT_API_KEY/ZT_LIBRARY_ID are missing"
                )
        elif self.zotero_read == "local":
            self.read_mode = "local" if self.local_reachable else "offline"
        elif self.local_reachable:
            self.read_mode = "local"
        elif self.web_credentials:
            self.read_mode = "web"
        else:
            self.read_mode = "offline"
        if self.local_reachable and self.local_writes:
            self.write_channel = "local"
        elif self.web_credentials:
            self.write_channel = "web"
        else:
            self.write_channel = None

    def mode_line(self) -> str:
        reads = {
            "local": f"reads: LOCAL API ({self.local_base}), CSL-JSON",
            "web": "reads: WEB API (api.zotero.org), CSL-JSON",
            "offline": "reads: NONE (local Zotero unreachable, no web credentials)",
        }[self.read_mode]
        if self.local_reachable and self.local_writes is False:
            probe = (
                f"local write probe refused (HTTP {self.local_write_status})"
                if self.local_write_status is not None
                else "local write probe refused"
            )
        elif self.local_writes:
            probe = "local API ACCEPTS writes"
        else:
            probe = "local API not reachable"
        if self.write_channel == "local":
            writes = "writes: LOCAL API"
        elif self.write_channel == "web":
            writes = "writes: WEB API (credentials present)"
        else:
            writes = "writes: IMPOSSIBLE (ZT_API_KEY / ZT_LIBRARY_ID missing)"
        return f"zoteroconv mode -> {reads} | {probe} | {writes}"

    # -- headers ---------------------------------------------------------------

    def _headers(self, csl: bool = True) -> dict:
        h = {}
        if csl:
            h["Accept"] = CSL_JSON_ACCEPT
        if self.write_channel == "web" or self.read_mode == "web":
            h["Zotero-API-Key"] = self.api_key or ""
            h["Zotero-API-Version"] = "3"
        return h

    def _base(self, for_write: bool = False) -> str:
        if for_write:
            return self.local_base if self.write_channel == "local" else (
                self.web_base + self.library_path()
            )
        return self.local_base if self.read_mode == "local" else (
            self.web_base + self.library_path()
        )

    # -- reads -----------------------------------------------------------------

    @staticmethod
    def _classify(items: Any) -> str:
        """'csl' | 'zotero' | 'empty' | 'bad' -- shape detection, because the
        Accept header is not honoured by every Zotero version."""
        if not isinstance(items, list):
            return "bad"
        if not items:
            return "empty"
        first = items[0]
        if not isinstance(first, dict):
            return "bad"
        if isinstance(first.get("data"), dict):
            return "zotero"
        if isinstance(first.get("id"), str):
            return "csl"
        return "bad"

    def _fetch_page(self, url: str, start: int, use_format_param: bool,
                    accept_csl: bool):
        params: dict[str, Any] = {"limit": 100, "start": start}
        if use_format_param:
            params["format"] = "csljson"
        headers = {"Accept": CSL_JSON_ACCEPT} if accept_csl else {}
        if self.write_channel == "web" or self.read_mode == "web":
            headers["Zotero-API-Key"] = self.api_key or ""
            headers["Zotero-API-Version"] = "3"
        resp = self.transport.request("GET", url, params=params,
                                      headers=headers, timeout=self.timeout)
        if resp.status != 200:
            return None, resp.status
        return resp.json, resp.status

    def fetch_items_csl(self) -> list[dict]:
        """Fetch library items as CSL-JSON.  Tries, in order:
        (1) Accept: application/vnd.citationstyles.csl+json
        (2) ?format=csljson          (user-directed fallback)
        (3) whatever the API returns natively -- Zotero JSON gets converted
            to CSL-JSON locally.  The working representation is reported."""
        if self.read_mode == "offline":
            self.notes.append("library not fetched (offline): treating as empty")
            return []
        url = self._base() + "/items"

        configs = [  # (use_format_param, accept_csl)
            (False, True),
            (True, False),
        ]
        chosen = None
        first_payload = None
        for cfg in configs:
            payload, status = self._fetch_page(url, 0, *cfg)
            kind = self._classify(payload)
            if kind in ("csl", "empty"):
                chosen, first_payload = ("csl", cfg), payload
                break
            if kind == "zotero":
                chosen, first_payload = ("zotero", cfg), payload
                break
            # non-200 or unexpected shape: try the next config
        if chosen is None:
            self.notes.append(
                f"item fetch failed with Accept header and format=csljson "
                f"(last HTTP {status}); treating library as empty"
            )
            return []

        rep, cfg = chosen
        if rep == "csl" and cfg == (False, True):
            self.notes.append("items fetched as CSL-JSON via Accept header")
        elif rep == "csl":
            self.notes.append("items fetched as CSL-JSON via ?format=csljson "
                              "(Accept header did not work)")
        else:
            self.notes.append("API returned Zotero item JSON; converted to "
                              "CSL-JSON locally (fallback path)")

        items: list[dict] = list(first_payload or [])
        if len(items) >= 100:  # paginate with the working configuration
            start = 100
            for _ in range(49):
                payload, status = self._fetch_page(url, start, *cfg)
                if not isinstance(payload, list) or not payload:
                    break
                items.extend(payload)
                if len(payload) < 100:
                    break
                start += 100
        if rep == "zotero":
            items = [zotero_json_to_csl(it, self) for it in items]
            items = [it for it in items if it is not None]
        self.learn_uri_prefix(items)
        return items

    def learn_uri_prefix(self, items_csl: list[dict]) -> None:
        """Sample-format URIs are http://zotero.org/users/local/<token>/items/<key>;
        learn <token> from existing local items so created items get the same form."""
        for it in items_csl:
            uri = str(it.get("id") or "")
            if "/users/local/" in uri and "/items/" in uri:
                self.local_uri_prefix = uri.split("/items/")[0]
                return

    def item_uri(self, key: str) -> str:
        if self.write_channel == "web" or self.read_mode == "web":
            kind = "groups" if self.library_type == "group" else "users"
            return f"http://zotero.org/{kind}/{self.library_id}/items/{key}"
        if self.local_uri_prefix:
            return f"{self.local_uri_prefix}/items/{key}"
        self.notes.append(
            f"no existing item URI to learn the local library token from; "
            f"using placeholder prefix for new item {key} (Zotero Refresh repairs it)"
        )
        return f"http://zotero.org/users/local/0/items/{key}"

    def fetch_collections(self) -> list[dict]:
        if self.write_channel is None and self.read_mode == "offline":
            return []
        url = self._base(for_write=True) + "/collections"
        resp = self.transport.request(
            "GET", url, params={"limit": 100},
            headers=self._headers(csl=False), timeout=self.timeout,
        )
        if resp.status == 200 and isinstance(resp.json, list):
            return resp.json
        self.notes.append(f"collection fetch failed (HTTP {resp.status})")
        return []

    # -- writes ------------------------------------------------------------------

    def create_collection(self, name: str) -> str | None:
        resp = self.transport.request(
            "POST", self._base(for_write=True) + "/collections",
            json=[{"name": name}], headers=self._headers(csl=False),
            timeout=self.timeout,
        )
        ok = (resp.json or {}).get("success") or {}
        first = ok.get("0") or {}
        if resp.status in (200, 201) and first.get("key"):
            return first["key"]
        self.notes.append(f"collection create failed (HTTP {resp.status})")
        return None

    def create_items(self, items: list[dict]) -> dict:
        resp = self.transport.request(
            "POST", self._base(for_write=True) + "/items", json=items,
            headers=self._headers(csl=False), timeout=self.timeout,
        )
        if resp.status not in (200, 201) or not isinstance(resp.json, dict):
            return {"failed": {str(i): {"message": f"HTTP {resp.status}"}
                               for i in range(len(items))}}
        return resp.json


# ---------------------------------------------------------------------------
# library items + stable numeric ids
# ---------------------------------------------------------------------------

# reverse of the creation maps: Zotero item JSON -> CSL-JSON (fallback when
# neither the Accept header nor ?format=csljson yields CSL-JSON)
Z2C_TYPE = {
    "journalArticle": "article-journal",
    "magazineArticle": "article-magazine",
    "newspaperArticle": "article-newspaper",
    "book": "book",
    "bookSection": "chapter",
    "conferencePaper": "paper-conference",
    "report": "report",
    "thesis": "thesis",
    "webpage": "webpage",
    "document": "document",
}
Z2C_CONTAINER = {
    "publicationTitle": "container-title",
    "bookTitle": "container-title",
    "websiteTitle": "container-title",
    "proceedingsName": "container-title",
    "seriesTitle": "container-title",
}
Z2C_PUBLISHER = {"publisher": "publisher", "institution": "publisher",
                 "university": "publisher"}
Z2C_SIMPLE = {
    "title": "title", "volume": "volume", "issue": "issue", "pages": "page",
    "DOI": "DOI", "ISBN": "ISBN", "ISSN": "ISSN", "url": "URL",
    "language": "language", "edition": "edition",
    "abstractNote": "abstract", "place": "publisher-place",
    "reportNumber": "number", "number": "number",
}


def zotero_json_to_csl(item: dict, access: "ZoteroAccess") -> dict | None:
    """Convert one Zotero API item ({key, data:{...}}) to CSL-JSON.  Only
    fields present in the source are emitted; nothing is invented."""
    from zoteroconv.parse_reference_list import parse_dateish

    data = item.get("data") or {}
    key = data.get("key") or item.get("key")
    if not key:
        return None
    csl: dict[str, Any] = {
        "id": access.item_uri(key),
        "type": Z2C_TYPE.get(data.get("itemType") or "document", "document"),
    }
    for z_key, csl_key in Z2C_SIMPLE.items():
        if data.get(z_key):
            csl[csl_key] = data[z_key]
    for z_key, csl_key in Z2C_CONTAINER.items():
        if data.get(z_key):
            csl[csl_key] = data[z_key]
            break
    for z_key, csl_key in Z2C_PUBLISHER.items():
        if data.get(z_key):
            csl[csl_key] = data[z_key]
            break
    if data.get("thesisType"):
        csl["genre"] = data["thesisType"]
    creators = data.get("creators") or []
    for creator_type, csl_key in (("author", "author"), ("editor", "editor")):
        people = []
        for c in creators:
            if (c.get("creatorType") or "author") != creator_type:
                continue
            if c.get("name"):
                people.append({"name": c["name"]})
            else:
                p = {}
                if c.get("lastName"):
                    p["family"] = c["lastName"]
                if c.get("firstName"):
                    p["given"] = c["firstName"]
                if p:
                    people.append(p)
        if people:
            csl[csl_key] = people
    for z_key, csl_key in (("date", "issued"), ("accessDate", "accessed")):
        if data.get(z_key):
            parts = parse_dateish(str(data[z_key]))
            if parts:
                csl[csl_key] = {"date-parts": [parts]}
    extra = str(data.get("extra") or "")
    m = re.search(r"PMID:\s*(\d+)", extra)
    if m:
        csl["PMID"] = m.group(1)
    return csl


def stable_numeric_id(seed: str) -> int:
    """Deterministic 8-digit positive int, the *form* the sample uses for
    citationItems[].id.  The numeric id is Zotero-internal: the Word plugin
    rewrites it on Refresh, so a stable synthetic value is safe."""
    return 10_000_000 + (zlib.crc32(seed.encode("utf-8")) % 90_000_000)


@dataclass
class LibraryItem:
    key: str
    uri: str
    csl: dict                 # full CSL-JSON (sample itemData form)
    numeric_id: int           # sample-form numeric id
    source: str               # 'local' | 'web' | 'existing-field' | 'created'


def library_items_from_csl(items: list[dict], source: str) -> list[LibraryItem]:
    out = []
    for it in items:
        uri = str(it.get("id") or "")
        key = uri.rstrip("/").rsplit("/", 1)[-1] if uri else ""
        if not key:
            continue
        out.append(LibraryItem(
            key=key, uri=uri, csl=dict(it),
            numeric_id=stable_numeric_id(uri or key), source=source,
        ))
    return out


# ---------------------------------------------------------------------------
# match cascade
# ---------------------------------------------------------------------------

@dataclass
class Candidate:
    source: str            # 'library' | 'crossref' | 'pubmed'
    score: float
    key: str | None = None
    doi: str | None = None
    title: str | None = None
    year: str | None = None
    surname: str | None = None
    raw: dict | None = None


@dataclass
class MatchOutcome:
    method: str | None                 # 'doi' | 'pmid' | 'isbn' | 'fuzzy' | None
    item: LibraryItem | None
    candidates: list[Candidate] = dc_field(default_factory=list)
    messages: list[str] = dc_field(default_factory=list)


def csl_year(csl: dict) -> str | None:
    try:
        parts = (csl.get("issued") or {}).get("date-parts") or [[None]]
        y = parts[0][0]
        return str(y) if y is not None else None
    except (IndexError, TypeError, AttributeError):
        return None


def csl_first_surname(csl: dict) -> str:
    for a in (csl.get("author") or []):
        if a.get("family"):
            return str(a["family"])
        if a.get("name"):
            return str(a["name"])
    return ""


def norm_isbn(s: str) -> str:
    return re.sub(r"[^0-9Xx]", "", s).upper()


def fuzzy_key(csl: dict) -> str:
    return " ".join(x for x in (
        norm_title(str(csl.get("title") or "")),
        csl_first_surname(csl).lower(),
        csl_year(csl) or "",
    ) if x)


def match_reference(
    csl: dict, library: list[LibraryItem], threshold: float = DEFAULT_THRESHOLD,
) -> MatchOutcome:
    """DOI exact -> PMID exact -> ISBN exact -> fuzzy title+surname+year.
    More than one fuzzy candidate above threshold => never auto-picked."""
    msgs: list[str] = []

    doi = csl.get("DOI")
    if doi:
        hits = [it for it in library
                if it.csl.get("DOI") and norm_doi(str(it.csl["DOI"])) == norm_doi(str(doi))]
        if hits:
            if len(hits) > 1:
                msgs.append(f"{len(hits)} library items share this DOI; using the first")
            return MatchOutcome("doi", hits[0], [], msgs)

    pmid = str(csl.get("PMID") or "").strip()
    if pmid:
        hits = [it for it in library
                if str(it.csl.get("PMID") or "").strip() == pmid]
        if hits:
            return MatchOutcome("pmid", hits[0], [], msgs)

    isbn = csl.get("ISBN")
    if isbn:
        want = norm_isbn(str(isbn))
        hits = [it for it in library
                if it.csl.get("ISBN") and norm_isbn(str(it.csl["ISBN"])) == want]
        if hits:
            return MatchOutcome("isbn", hits[0], [], msgs)

    title = norm_title(str(csl.get("title") or ""))
    if not title:
        msgs.append("no usable title for fuzzy matching")
        return MatchOutcome(None, None, [], msgs)

    q = fuzzy_key(csl)
    scored: list[tuple[float, LibraryItem]] = []
    for it in library:
        k = fuzzy_key(it.csl)
        if not k:
            continue
        scored.append((float(fuzz.token_sort_ratio(q, k)), it))
    above = sorted((p for p in scored if p[0] >= threshold),
                   key=lambda p: -p[0])
    cands = [
        Candidate(source="library", score=round(s, 1), key=it.key,
                  doi=it.csl.get("DOI"), title=it.csl.get("title"),
                  year=csl_year(it.csl), surname=csl_first_surname(it.csl),
                  raw=it.csl)
        for s, it in above
    ]
    if len(above) == 1:
        return MatchOutcome("fuzzy", above[0][1], cands, msgs)
    if len(above) > 1:
        msgs.append(
            f"{len(above)} library candidates score >= {threshold:.0f}; "
            "not auto-picked"
        )
        return MatchOutcome(None, None, cands, msgs)
    return MatchOutcome(None, None, [], msgs)


# ---------------------------------------------------------------------------
# CrossRef / PubMed suggestions (cached in SQLite, polite User-Agent)
# ---------------------------------------------------------------------------

def _ua(mailto: str | None) -> str:
    contact = f"; mailto:{mailto}" if mailto else ""
    return f"zoteroconv/0.2 ({TOOL_URL}{contact})"


def _sha1(s: str) -> str:
    return hashlib.sha1(s.encode("utf-8")).hexdigest()


def work_fields_crossref(work: dict) -> dict:
    title = ((work.get("title") or [""])[0] or "")
    year = None
    try:
        year = str(work["issued"]["date-parts"][0][0])
    except (KeyError, IndexError, TypeError):
        pass
    surname = ""
    for a in work.get("author") or []:
        surname = str(a.get("family") or a.get("name") or "")
        if surname:
            break
    return {"title": title, "year": year, "surname": surname,
            "doi": work.get("DOI"), "raw": work}


def evaluate_suggestion(ref_csl: dict, sugg: dict) -> tuple[bool, float, str]:
    """Decision 5: accept only title similarity > 95 AND year AND first-author
    surname matching.  Returns (accepted, title_sim, explanation)."""
    sim = float(fuzz.token_sort_ratio(
        norm_title(str(ref_csl.get("title") or "")),
        norm_title(str(sugg.get("title") or "")),
    ))
    ref_year, sug_year = csl_year(ref_csl), sugg.get("year")
    year_ok = bool(ref_year) and str(sug_year) == str(ref_year)
    ref_sur = csl_first_surname(ref_csl).lower()
    sug_sur = str(sugg.get("surname") or "").lower()
    author_ok = bool(ref_sur) and bool(sug_sur) and (
        fuzz.token_sort_ratio(ref_sur, sug_sur) >= SUGGEST_AUTHOR_ACCEPT
    )
    why = (
        f"title similarity {sim:.1f} (need >{SUGGEST_TITLE_ACCEPT:.0f}), "
        f"year {'ok' if year_ok else f'MISMATCH (ref {ref_year} vs {sug_year})'}, "
        f"first author {'ok' if author_ok else 'MISMATCH'}"
    )
    return (sim > SUGGEST_TITLE_ACCEPT and year_ok and author_ok), sim, why


class CrossRefClient:
    def __init__(self, transport: Transport, cache: HttpCache | None = None,
                 mailto: str | None = None, timeout: float = 10.0):
        self.transport = transport
        self.cache = cache
        self.mailto = mailto
        self.timeout = timeout

    def suggest(self, bibliographic: str, rows: int = 5) -> list[dict]:
        """Query CrossRef with the bibliographic string ONLY (decision 5)."""
        key = _sha1(f"works|{rows}|{bibliographic}")
        if self.cache is not None:
            hit = self.cache.get("crossref", key)
            if hit is not None:
                return hit
        params: dict[str, Any] = {"query.bibliographic": bibliographic, "rows": rows}
        if self.mailto:
            params["mailto"] = self.mailto
        try:
            resp = self.transport.request(
                "GET", CROSSREF_WORKS_URL, params=params,
                headers={"User-Agent": _ua(self.mailto)}, timeout=self.timeout,
            )
        except Exception:
            return []
        works: list[dict] = []
        if resp.status == 200 and isinstance(resp.json, dict):
            works = (resp.json.get("message") or {}).get("items") or []
        if self.cache is not None:
            self.cache.set("crossref", key, works)
        return works


class PubmedClient:
    def __init__(self, transport: Transport, cache: HttpCache | None = None,
                 mailto: str | None = None, timeout: float = 10.0):
        self.transport = transport
        self.cache = cache
        self.mailto = mailto
        self.timeout = timeout

    def suggest_for_pmid(self, pmid: str) -> dict | None:
        pmid = str(pmid).strip()
        if not pmid.isdigit():
            return None
        if self.cache is not None:
            hit = self.cache.get("pubmed", pmid)
            if hit is not None:
                return hit or None
        try:
            resp = self.transport.request(
                "GET", PUBMED_ESUMMARY_URL,
                params={"db": "pubmed", "id": pmid, "retmode": "json"},
                headers={"User-Agent": _ua(self.mailto)}, timeout=self.timeout,
            )
        except Exception:
            return None
        out: dict | None = None
        if resp.status == 200 and isinstance(resp.json, dict):
            rec = (resp.json.get("result") or {}).get(pmid)
            if isinstance(rec, dict):
                surname = ""
                authors = rec.get("authors") or []
                if authors:
                    surname = str(authors[0].get("name") or "").split(" ")[0]
                doi = None
                for aid in rec.get("articleids") or []:
                    if aid.get("idtype") == "doi":
                        doi = aid.get("value")
                out = {
                    "title": rec.get("title") or "",
                    "year": str(rec.get("pubdate") or "")[:4] or None,
                    "surname": surname,
                    "doi": doi,
                    "raw": rec,
                }
        if self.cache is not None:
            self.cache.set("pubmed", pmid, out or {})
        return out


# ---------------------------------------------------------------------------
# CSL-JSON -> Zotero item JSON (for creates; only fields that exist)
# ---------------------------------------------------------------------------

TYPE_MAP = {
    "article-journal": "journalArticle",
    "article-magazine": "magazineArticle",
    "article-newspaper": "newspaperArticle",
    "article": "journalArticle",
    "book": "book",
    "chapter": "bookSection",
    "paper-conference": "conferencePaper",
    "report": "report",
    "thesis": "thesis",
    "webpage": "webpage",
    "document": "document",
}
CONTAINER_FIELD = {
    "journalArticle": "publicationTitle",
    "magazineArticle": "publicationTitle",
    "newspaperArticle": "publicationTitle",
    "bookSection": "bookTitle",
    "webpage": "websiteTitle",
    "conferencePaper": "proceedingsName",
    "report": "seriesTitle",
}
PUBLISHER_FIELD = {
    "book": "publisher",
    "report": "institution",
    "thesis": "university",
    "conferencePaper": "publisher",
}
PLACE_FIELD = {"book": "place", "report": "place", "thesis": "place",
               "conferencePaper": "place"}
PAGED_TYPES = {"journalArticle", "magazineArticle", "newspaperArticle",
               "conferencePaper"}


def _iso_date(parts: list) -> str | None:
    if not parts:
        return None
    p = [str(x) for x in parts if x is not None]
    if not p:
        return None
    if len(p) == 1:
        return p[0]
    if len(p) == 2:
        return f"{p[0]}-{int(p[1]):02d}"
    return f"{p[0]}-{int(p[1]):02d}-{int(p[2]):02d}"


def csl_to_zotero_item(csl: dict) -> tuple[dict, list[str]]:
    """Translate parsed CSL-JSON into a Zotero API item dict.  Only fields
    present in the CSL are transferred; nothing is invented.  Fields with no
    Zotero counterpart for the item type are reported, not dropped silently."""
    notes: list[str] = []
    csl_type = csl.get("type") or "document"
    item_type = TYPE_MAP.get(csl_type, "document")
    if csl_type not in TYPE_MAP:
        notes.append(f"unknown CSL type {csl_type!r}; created as 'document'")
    z: dict[str, Any] = {"itemType": item_type}
    handled = {"type", "id"}

    if csl.get("title"):
        z["title"] = csl["title"]
    handled.add("title")

    for csl_key, creator_type in (("author", "author"), ("editor", "editor")):
        handled.add(csl_key)
        people = csl.get(csl_key) or []
        if not people:
            continue
        creators = z.setdefault("creators", [])
        for p in people:
            if p.get("family") or p.get("given"):
                creators.append({
                    "creatorType": creator_type,
                    "lastName": p.get("family", ""),
                    "firstName": p.get("given", ""),
                })
            elif p.get("name"):
                creators.append({"creatorType": creator_type, "name": p["name"]})

    handled |= {"issued", "accessed"}
    issued = _iso_date(((csl.get("issued") or {}).get("date-parts") or [[]])[0])
    if issued:
        z["date"] = issued
    accessed = _iso_date(((csl.get("accessed") or {}).get("date-parts") or [[]])[0])
    if accessed:
        z["accessDate"] = accessed

    handled.add("container-title")
    if csl.get("container-title"):
        f = CONTAINER_FIELD.get(item_type)
        if f:
            z[f] = csl["container-title"]
        else:
            notes.append(
                f"container-title not transferred (no Zotero field on {item_type})"
            )

    handled.add("publisher")
    if csl.get("publisher"):
        f = PUBLISHER_FIELD.get(item_type)
        if f:
            z[f] = csl["publisher"]
        else:
            notes.append(
                f"publisher not transferred (no Zotero field on {item_type}); "
                "kept in the CSL-JSON only"
            )

    handled.add("publisher-place")
    if csl.get("publisher-place"):
        f = PLACE_FIELD.get(item_type)
        if f:
            z[f] = csl["publisher-place"]
        else:
            notes.append(
                f"publisher-place not transferred (no Zotero field on {item_type})"
            )

    handled |= {"volume", "issue", "page"}
    if item_type in PAGED_TYPES:
        if csl.get("volume"):
            z["volume"] = str(csl["volume"])
        if csl.get("issue"):
            z["issue"] = str(csl["issue"])
        if csl.get("page"):
            z["pages"] = str(csl["page"])
    else:
        for k in ("volume", "issue", "page"):
            if csl.get(k):
                notes.append(f"{k} not transferred (not valid on {item_type})")

    handled.add("number")
    if csl.get("number"):
        if item_type == "report":
            z["reportNumber"] = str(csl["number"])
        elif item_type == "book":
            z["number"] = str(csl["number"])
        else:
            notes.append(f"number not transferred (no Zotero field on {item_type})")

    simple = {
        "DOI": "DOI", "ISBN": "ISBN", "ISSN": "ISSN", "URL": "url",
        "language": "language", "edition": "edition",
        "abstract": "abstractNote",
    }
    handled |= set(simple)
    for csl_key, z_key in simple.items():
        if csl.get(csl_key):
            z[z_key] = str(csl[csl_key])

    handled.add("genre")
    if csl.get("genre"):
        if item_type == "thesis":
            z["thesisType"] = csl["genre"]
        else:
            notes.append(f"genre not transferred (no Zotero field on {item_type})")

    handled.add("PMID")
    if csl.get("PMID"):
        extra = z.get("extra") or ""
        z["extra"] = (extra + "\n" if extra else "") + f"PMID: {csl['PMID']}"

    for k in csl:
        if k not in handled:
            notes.append(f"CSL field {k!r} has no Zotero mapping; not transferred")
    return z, notes


# ---------------------------------------------------------------------------
# resolution pipeline
# ---------------------------------------------------------------------------

CREATE_PLANNED_STATUSES = (
    "create-planned",
    "crossref-accepted",
    "pubmed-accepted",
    "crossref-low-confidence",
    "pubmed-low-confidence",
)


@dataclass
class RefResolution:
    ref_number: int
    raw: str
    parsed_csl: dict
    status: str            # see module docstring / STATUSES below
    method: str | None = None
    item: LibraryItem | None = None
    candidates: list[Candidate] = dc_field(default_factory=list)
    messages: list[str] = dc_field(default_factory=list)
    accepted_doi: str | None = None
    suggestion: dict | None = None

    def effective_csl(self) -> dict:
        csl = dict(self.parsed_csl)
        if self.accepted_doi and not csl.get("DOI"):
            csl["DOI"] = self.accepted_doi
        return csl


# statuses:
#   reused-existing-field  matched  ambiguous
#   crossref-accepted  crossref-low-confidence  pubmed-accepted  pubmed-low-confidence
#   create-planned  created  create-failed  duplicate-skip
#   skipped-max-new-items  create-blocked-no-writes


@dataclass
class ResolutionReport:
    mode: str                       # 'local' | 'web' | 'offline' (reads)
    mode_line: str
    local_writes: bool | None
    write_channel: str | None
    dry_run: bool
    threshold: float
    resolutions: list[RefResolution] = dc_field(default_factory=list)
    collection_name: str | None = None
    collection_key: str | None = None
    created_keys: list[str] = dc_field(default_factory=list)
    uncited: list[int] = dc_field(default_factory=list)
    notes: list[str] = dc_field(default_factory=list)

    def by_ref(self, n: int) -> RefResolution | None:
        for r in self.resolutions:
            if r.ref_number == n:
                return r
        return None

    def counts(self) -> dict[str, int]:
        out: dict[str, int] = {}
        for r in self.resolutions:
            out[r.status] = out.get(r.status, 0) + 1
        return out

    def summary_lines(self) -> list[str]:
        lines = [self.mode_line]
        for n in sorted(self.counts()):
            lines.append(f"  {n}: {self.counts()[n]}")
        for r in self.resolutions:
            extra = ""
            if r.item:
                extra = f" -> {r.item.key} ({r.method or r.status})"
            lines.append(f"  ref {r.ref_number:>3}: {r.status}{extra}")
            for c in r.candidates:
                lines.append(
                    f"        candidate [{c.source}] score={c.score} "
                    f"key={c.key or '-'} doi={c.doi or '-'} "
                    f"title={str(c.title or '')[:60]!r}"
                )
            for m in r.messages:
                lines.append(f"        note: {m}")
        for n in self.notes:
            lines.append(f"  ! {n}")
        return lines


def _pick_best_suggestion(ref_csl: dict, works_fields: list[dict]
                          ) -> tuple[dict | None, float, bool, str]:
    best, best_sim, best_ok, best_why = None, -1.0, False, ""
    for wf in works_fields:
        ok, sim, why = evaluate_suggestion(ref_csl, wf)
        if sim > best_sim:
            best, best_sim, best_ok, best_why = wf, sim, ok, why
    return best, best_sim, best_ok, best_why


def resolve_document(
    pkg: DocxPackage,
    *,
    threshold: float = DEFAULT_THRESHOLD,
    max_new_items: int | None = DEFAULT_MAX_NEW_ITEMS,
    dry_run: bool = False,
    collection_name: str | None = None,
    confirm: Callable[[str], bool] | None = None,
    transport: Transport | None = None,
    cache: HttpCache | None = None,
    api_key: str | None = None,
    library_id: str | None = None,
    library_type: str | None = None,
    zotero_read: str = "auto",
    crossref_mailto: str | None = None,
    doc_name: str | None = None,
    emit: Callable[[str], None] = print,
    reflist: ReferenceList | None = None,
    cited_numbers: set[int] | None = None,
    keep_uncited: bool = False,
) -> ResolutionReport:
    """Resolve every reference of ``pkg`` against Zotero.  Credentials are
    taken from parameters or the environment (ZT_API_KEY, ZT_LIBRARY_ID,
    ZT_LIBRARY_TYPE, CROSSREF_MAILTO) and are never printed."""
    env = os.environ
    api_key = api_key or env.get("ZT_API_KEY") or None
    library_id = library_id or env.get("ZT_LIBRARY_ID") or None
    library_type = library_type or env.get("ZT_LIBRARY_TYPE") or "user"
    crossref_mailto = crossref_mailto or env.get("CROSSREF_MAILTO") or None

    transport = transport or RequestsTransport()
    access = ZoteroAccess(
        transport, api_key=api_key, library_id=library_id,
        library_type=library_type, zotero_read=zotero_read,
    )
    access.detect()
    emit(access.mode_line())

    report = ResolutionReport(
        mode=access.read_mode, mode_line=access.mode_line(),
        local_writes=access.local_writes, write_channel=access.write_channel,
        dry_run=dry_run, threshold=threshold,
    )

    rl = reflist if reflist is not None else find_reference_list(pkg)
    if rl is None:
        report.notes.append("no reference list found in the document; nothing to resolve")
        report.notes = list(access.notes) + report.notes
        return report

    library = library_items_from_csl(access.fetch_items_csl(), access.read_mode)
    if library:
        report.notes.append(
            f"library: {len(library)} items fetched via {access.read_mode} API (CSL-JSON)"
        )

    crossref = CrossRefClient(transport, cache=cache, mailto=crossref_mailto)
    pubmed = PubmedClient(transport, cache=cache, mailto=crossref_mailto)

    # decision 4: reuse items behind existing Zotero fields (never re-created)
    fields = find_zotero_fields(pkg.document, DOCUMENT_PART)
    maps = map_existing_fields_to_references(fields, rl.references)
    mapped_by_ref: dict[int, Any] = {}
    for m in maps:
        if m.ref_number is not None and m.ref_number not in mapped_by_ref:
            mapped_by_ref[m.ref_number] = m

    for ref in rl.references:
        m = mapped_by_ref.get(ref.number)
        if m is not None and m.item_data:
            uri = m.uri or ""
            item = LibraryItem(
                key=m.item_key or "", uri=uri, csl=dict(m.item_data),
                numeric_id=stable_numeric_id(uri or m.item_key or ""),
                source="existing-field",
            )
            report.resolutions.append(RefResolution(
                ref_number=ref.number, raw=ref.raw, parsed_csl=ref.csl,
                status="reused-existing-field",
                method=f"existing-field-{m.method}",
                item=item,
                messages=[
                    f"item {item.key} already cited by Zotero field "
                    f"{m.citation_id!r}; will not be created again"
                ],
            ))
            continue
        if (cited_numbers is not None and ref.number not in cited_numbers
                and not keep_uncited):
            # user directive: uncited refs are left OUT by default -- resolve
            # read-only against the library, but never enrich or create
            outcome = match_reference(ref.csl, library, threshold)
            res = RefResolution(
                ref_number=ref.number, raw=ref.raw, parsed_csl=ref.csl,
                status="matched" if outcome.item else "uncited-not-created",
                method=outcome.method, item=outcome.item,
                candidates=outcome.candidates, messages=list(outcome.messages),
            )
            if outcome.item is None:
                res.messages.append(
                    "reference is never cited numerically; left out of the "
                    "bibliography and NOT created (use --keep-uncited to keep it)"
                )
            report.resolutions.append(res)
            continue
        _resolve_one(ref, library, crossref, pubmed, threshold, report)

    if cited_numbers is not None:
        report.uncited = sorted(
            r.number for r in rl.references
            if r.number not in cited_numbers and r.number not in mapped_by_ref
        )
        if report.uncited:
            report.notes.append(
                f"references never cited numerically: {report.uncited} "
                f"({'kept via uncited array' if keep_uncited else 'left OUT of the bibliography by default; see --keep-uncited'})"
            )

    _create_missing(report, access, rl, library, max_new_items=max_new_items,
                    dry_run=dry_run, confirm=confirm, collection_name=collection_name,
                    doc_name=doc_name or (Path(pkg.path).stem if pkg.path else None),
                    emit=emit)
    report.notes = list(access.notes) + report.notes
    return report


def _resolve_one(ref: Reference, library: list[LibraryItem],
                 crossref: CrossRefClient, pubmed: PubmedClient,
                 threshold: float, report: ResolutionReport) -> None:
    outcome = match_reference(ref.csl, library, threshold)
    res = RefResolution(ref_number=ref.number, raw=ref.raw, parsed_csl=ref.csl,
                        status="", method=outcome.method, item=outcome.item,
                        candidates=outcome.candidates,
                        messages=list(outcome.messages))
    if outcome.item is not None:
        res.status = "matched"
        report.resolutions.append(res)
        return
    if outcome.candidates:
        res.status = "ambiguous"
        report.resolutions.append(res)
        return

    has_ids = any(ref.csl.get(k) for k in ("DOI", "PMID", "ISBN"))
    suggestion: dict | None = None
    sim = -1.0
    accepted = False
    why = ""
    source = ""

    if not has_ids:
        works = crossref.suggest(ref.raw)
        fields_list = [work_fields_crossref(w) for w in works]
        suggestion, sim, accepted, why = _pick_best_suggestion(ref.csl, fields_list)
        source = "crossref"
    elif ref.csl.get("PMID") and not ref.csl.get("DOI"):
        pm_sugg = pubmed.suggest_for_pmid(str(ref.csl["PMID"]))
        if pm_sugg:
            ok, tsim, twhy = evaluate_suggestion(ref.csl, pm_sugg)
            suggestion, sim, accepted, why = pm_sugg, tsim, ok, twhy
            source = "pubmed"

    if suggestion is not None:
        res.suggestion = suggestion.get("raw") or suggestion
        res.candidates.append(Candidate(
            source=source, score=round(sim, 1), doi=suggestion.get("doi"),
            title=suggestion.get("title"), year=suggestion.get("year"),
            surname=suggestion.get("surname"), raw=res.suggestion,
        ))
        if accepted and suggestion.get("doi"):
            # a verified DOI may still exist in the library -> re-check first
            recheck = match_reference({**ref.csl, "DOI": suggestion["doi"]},
                                      library, threshold)
            if recheck.item is not None:
                res.status = "matched"
                res.method = f"{source}-doi->{recheck.method}"
                res.item = recheck.item
                res.messages.append(
                    f"{source} suggestion accepted ({why}); DOI already in library"
                )
                report.resolutions.append(res)
                return
            res.status = f"{source}-accepted"
            res.accepted_doi = suggestion["doi"]
            res.messages.append(
                f"{source} suggestion ACCEPTED ({why}); DOI {suggestion['doi']} "
                "will be added on create"
            )
            report.resolutions.append(res)
            return
        res.status = f"{source}-low-confidence" if source else "create-planned"
        res.messages.append(
            f"{source or 'identifier'} suggestion below acceptance bar ({why}); "
            "parsed item kept as placeholder, no metadata taken from the suggestion"
        )
        report.resolutions.append(res)
        return

    res.status = "create-planned"
    report.resolutions.append(res)


def _create_missing(
    report: ResolutionReport, access: ZoteroAccess, rl: ReferenceList,
    library: list[LibraryItem], *, max_new_items: int | None, dry_run: bool,
    confirm: Callable[[str], bool] | None, collection_name: str | None,
    doc_name: str | None, emit: Callable[[str], None],
) -> None:
    plan = [r for r in report.resolutions if r.status in CREATE_PLANNED_STATUSES]
    if not plan:
        return

    if dry_run:
        report.notes.append(
            f"dry-run: {len(plan)} item(s) would be created; no writes performed"
        )
        return

    if access.write_channel is None:
        report.notes.append(
            "STOPPED BEFORE ANY WRITE: creating items requires the Zotero Web "
            "API with ZT_API_KEY and ZT_LIBRARY_ID (env or flags); the local "
            "API refuses writes. No credentials were found, so "
            f"{len(plan)} item(s) stay planned but uncreated."
        )
        for r in plan:
            r.status = "create-blocked-no-writes"
        emit(report.notes[-1])
        return

    # decision 6: --max-new-items with confirmation above the cap
    if max_new_items is not None and len(plan) > max_new_items:
        question = (
            f"{len(plan)} new items exceed --max-new-items={max_new_items}. "
            "Create all of them anyway?"
        )
        proceed = bool(confirm(question)) if confirm is not None else False
        if not proceed:
            overflow = plan[max_new_items:]
            for r in overflow:
                r.status = "skipped-max-new-items"
                r.messages.append(
                    f"above --max-new-items={max_new_items}; confirmation declined"
                )
            plan = plan[:max_new_items]
            report.notes.append(
                f"--max-new-items={max_new_items} applied; "
                f"{len(overflow)} creation(s) skipped"
            )

    # decision 6: collection "Imported - <doc name>", checked before create
    name = collection_name or f"Imported - {doc_name or 'document'}"
    report.collection_name = name
    coll_key: str | None = None
    existing = access.fetch_collections()
    for c in existing:
        data = c.get("data") or {}
        if data.get("name") == name:
            coll_key = c.get("key")
            break
    if coll_key is None:
        coll_key = access.create_collection(name)
        if coll_key:
            report.notes.append(f"collection {name!r} created (key {coll_key})")
        else:
            report.notes.append(
                f"collection {name!r} could not be created; items will be "
                "created without it"
            )
    else:
        report.notes.append(f"collection {name!r} already exists (key {coll_key})")
    report.collection_key = coll_key

    created_items: list[LibraryItem] = []
    for r in plan:
        # decision 6: dedup check before EVERY create, against the library AND
        # everything created earlier in this same run
        dup = match_reference(r.effective_csl(), library + created_items,
                              report.threshold)
        if dup.item is not None:
            r.status = "duplicate-skip"
            r.item = dup.item
            r.method = f"dedup-before-create:{dup.method}"
            r.messages.append(
                f"duplicate found before create; reusing {dup.item.key}"
            )
            continue
        z_item, conv_notes = csl_to_zotero_item(r.effective_csl())
        r.messages.extend(conv_notes)
        # user directive: tag creates so they can be found and cleaned up
        tags = [{"tag": "zoteroconv-created"}]
        if r.status in ("crossref-low-confidence", "pubmed-low-confidence"):
            tags.append({"tag": "zoteroconv-review"})
        z_item["tags"] = tags
        if coll_key:
            z_item["collections"] = [coll_key]
        resp = access.create_items([z_item])
        ok = (resp.get("success") or {}).get("0")
        failed = (resp.get("failed") or {}).get("0")
        if ok and ok.get("key"):
            key = ok["key"]
            uri = access.item_uri(key)
            csl = r.effective_csl()
            csl["id"] = uri
            item = LibraryItem(key=key, uri=uri, csl=csl,
                               numeric_id=stable_numeric_id(uri),
                               source="created")
            r.status = "created"
            r.method = "created"
            r.item = item
            created_items.append(item)
            report.created_keys.append(key)
        else:
            r.status = "create-failed"
            msg = ((failed or {}).get("message")
                   or "unexpected Zotero API response")
            r.messages.append(f"Zotero API create failed: {msg}")
    if report.created_keys:
        report.notes.append(
            f"created {len(report.created_keys)} item(s) via "
            f"{access.write_channel} API into collection {name!r}"
        )
