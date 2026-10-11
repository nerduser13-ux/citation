# Zotero Word-field format — structural facts (Phase 0)

Source of truth (per project rule 6): the real Zotero-produced fixtures only.

| fixture | provenance | style in ZOTERO_PREF | fields |
|---|---|---|---|
| `tests/fixtures/zotero_sample.docx` | produced by the real Zotero Word plugin (Zotero 10.0.6) | `elsevier-harvard` (no locale attr) | 1 × ZOTERO_ITEM, 1 × ZOTERO_BIBL |
| `tests/fixtures/real_vancouver.docx` | user-provided input document; already contains one Zotero citation + bibliography inserted by Zotero 10.0.6, plus a typed Vancouver reference list and numeric `[n]` citations | `harvard-cite-them-right`, `locale="en-US"` | 1 × ZOTERO_ITEM, 1 × ZOTERO_BIBL |

All facts below were produced by:

```
.venv/bin/python scripts/inspect_docx.py tests/fixtures/zotero_sample.docx   # -> docs/phase0_inspection/zotero_sample_inspection.txt
.venv/bin/python scripts/inspect_docx.py tests/fixtures/real_vancouver.docx  # -> docs/phase0_inspection/real_vancouver_inspection.txt
```

Every fact carries an **Evidence** line quoting the inspection output
(`S` = zotero_sample dump, `R` = real_vancouver dump). The inspector is
strictly read-only; fixture sha256 sums were identical before and after
every run (`3f5784a7…b901a` and `b2fb87e3…e09a9`).

---

## 1. Where Zotero data lives in the package

**1.1** In-text citations and the bibliography live as Word fields inside
`word/document.xml`. No Zotero markers exist anywhere else except
`docProps/custom.xml`.

> Evidence (S, SECTION A.2): `word/document.xml  ZOTERO_ITEM=1  ZOTERO_BIBL=1  ZOTERO_PREF=0  fldChar=12` and `docProps/custom.xml  ZOTERO_PREF=2`; all other parts have zero occurrences.

**1.2** `word/footnotes.xml` and `word/endnotes.xml` may exist and contain
zero fields (real doc) or may not exist at all (sample). Code must scan
all three parts and tolerate missing ones.

> Evidence (S, SECTION B): `word/footnotes.xml: PART NOT PRESENT in this docx`, `word/endnotes.xml: PART NOT PRESENT in this docx`.
> Evidence (R, SECTION B): `word/footnotes.xml: 0 field(s) total, 0 Zotero field(s)`, `word/endnotes.xml: 0 field(s) total, 0 Zotero field(s)`.

**1.3** Preferences live in `docProps/custom.xml`, wired into the package
by a content-type Override and a package-level relationship:

```xml
<Override PartName="/docProps/custom.xml"
          ContentType="application/vnd.openxmlformats-officedocument.custom-properties+xml"/>
<Relationship Id="rId4"   <!-- rId5 in the real doc: Ids vary, allocate next free -->
              Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties"
              Target="docProps/custom.xml"/>
```

> Evidence (S, SECTION D): both elements quoted verbatim (`Id="rId4"`).
> Evidence (R, SECTION D): same elements, `Id="rId5"`.

---

## 2. `ZOTERO_PREF_*` custom properties

**2.1** Properties are `<property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="N" name="ZOTERO_PREF_N"><vt:lpwstr>…</vt:lpwstr></property>`;
the value type element is `vt:lpwstr` (vt = `…/docPropsVTypes`).

> Evidence (S, SECTION C.1 raw file): `<property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="2" name="ZOTERO_PREF_1"><vt:lpwstr>&lt;data data-version="3" …`.

**2.2** The value is an XML-escaped XML document. Concatenating all chunks
in suffix order yields well-formed XML:

```xml
<data data-version="3" zotero-version="10.0.6">
  <session id="GahNCFud"/>
  <style id="http://www.zotero.org/styles/elsevier-harvard" hasBibliography="1" bibliographyStyleHasBeenSet="1"/>
  <prefs>
    <pref name="fieldType" value="Field"/>
    <pref name="automaticJournalAbbreviations" value="true"/>
  </prefs>
</data>
```

> Evidence (S, SECTION F): `Concatenated value length: 299` / `Concatenation parses as well-formed XML` + the pretty-printed tree above.
> Evidence (R, SECTION F): same shape, `zotero-version="10.0.6"`, `<session id="88ppEdET"/>`, `<style id="http://www.zotero.org/styles/harvard-cite-them-right" locale="en-US" hasBibliography="1" bibliographyStyleHasBeenSet="1"/>`, concatenated length 321.

**2.3** Attribute inventory of `<data>` / children (observed values):

- `data-version="3"`, `zotero-version="10.0.6"` (both docs)
- `<session id="…"/>` — 8 random chars, differs per document (`GahNCFud` vs `88ppEdET`)
- `<style id="<CSL style URL>" hasBibliography="1" bibliographyStyleHasBeenSet="1"/>`;
  an optional `locale="en-US"` attribute appears in the real doc but not in the sample
- `<prefs>` contains exactly `<pref name="fieldType" value="Field"/>` and
  `<pref name="automaticJournalAbbreviations" value="true"/>` in both docs

> Evidence: (S, SECTION F) pretty-printed concatenation; (R, SECTION F) pretty-printed concatenation showing `locale="en-US"`.

**2.4** Splitting rule: chunk names are `ZOTERO_PREF_1`, `ZOTERO_PREF_2`, …
(**1-based**; no `ZOTERO_PREF` or `ZOTERO_PREF_0` observed). All chunks
except the last are exactly **255 characters**; the split is a raw
character cut — it can land mid-token (here mid-attribute-name:
`…automaticJournal` + `Abbreviations"…`). `pid` values ascend with the
chunk index and start at **2** in both documents.

> Evidence (S, SECTION F): `name='ZOTERO_PREF_1' pid=2 value-length=255`, `name='ZOTERO_PREF_2' pid=3 value-length=44`, `All chunks except the last have identical length 255 -> fixed-size character split at 255 chars.`, boundary `...'automaticJournal'` / `'Abbreviations" value...'`, `pid values in suffix order: [2, 3] (ascending)`.
> Evidence (R, SECTION F): `ZOTERO_PREF_1 pid=2 value-length=255`, `ZOTERO_PREF_2 pid=3 value-length=66`, same fixed-size conclusion.

---

## 3. Citation field: `ADDIN ZOTERO_ITEM CSL_CITATION`

**3.1** Run layout (all runs are direct children of the same `w:p`, each a
separate `w:r`):

```
w:r[ w:fldChar @fldCharType="begin" ]      <- no rPr observed
w:r[ w:instrText @xml:space="preserve" ]   <- whole instruction in ONE instrText run in both fixtures
w:r[ w:fldChar @fldCharType="separate" ]   <- no rPr observed
w:r… preview runs (w:t)                    <- see 3.5
w:r[ w:fldChar @fldCharType="end" ]        <- no rPr observed
```

Word artifacts (`w:proofErr`, `w:lastRenderedPageBreak`) can appear
between/inside these runs but are not written by Zotero.

> Evidence (S, SECTION B.1 "complete field XML"): begin run, single instrText run, separate run, three preview runs interleaved with `w:proofErr`, end run — all inside `<w:field-range>` for paragraph #0.
> Evidence (R, SECTION B.2): begin run also contains `<w:lastRenderedPageBreak/>` before its `<w:fldChar w:fldCharType="begin"/>`.
> Evidence (S/R, SECTION E "fldChar-carrying runs"): `begin: run rPr=(no rPr)`, `separate: run rPr=(no rPr)`, `end: run rPr=(no rPr)`.

**3.2** The fldChar elements are bare — no `w:fldLock`, no `w:dirty`:
`<w:fldChar w:fldCharType="begin"/>` etc.

> Evidence (S, SECTION B.1 "fldChar elements (verbatim)"): all three shown with only `w:fldCharType`.

**3.3** Instruction text (leading and trailing space are part of it):

```
" ADDIN ZOTERO_ITEM CSL_CITATION {…json…} "
```

> Evidence (S, SECTION B.1): `instrText fragment 1 (xml:space='preserve'): ' ADDIN ZOTERO_ITEM CSL_CITATION {"citationID":"7yuwLzUY",…} '` (single fragment, `instrText runs: 1`).

**3.4** JSON payload — exact key order and value shapes (sample):

```json
{
  "citationID": "7yuwLzUY",
  "properties": {
    "unsorted": false,
    "formattedCitation": "(Havryk and Nazarova, 2024)",
    "plainCitation": "(Havryk and Nazarova, 2024)",
    "noteIndex": 0
  },
  "citationItems": [
    {
      "id": 1,
      "uris": ["http://zotero.org/users/local/XXnzLBe9/items/NU469JIQ"],
      "itemData": { "id": 1, "type": "article-journal", … }
    }
  ],
  "schema": "https://github.com/citation-style-language/schema/raw/master/csl-citation.json"
}
```

Observed rules:

- Top-level key order: `citationID`, `properties`, `citationItems`, `schema`.
- `citationID`: 8-char alphanumeric string (`7yuwLzUY` in S; `xO3paPuA` in R).
- `properties` key order: `unsorted`, `formattedCitation`, `plainCitation`, `noteIndex`;
  `noteIndex` is `0` for body citations; `formattedCitation` == `plainCitation` in both fixtures.
- `citationItems[i]` key order: `id`, `uris`, `itemData`. No `prefix`/`suffix`
  keys when empty.
- `id` is a **document-local integer** (both fixtures: `1` for their only
  citation item; identical item `NU469JIQ` → id is not a library id).
- `uris` form for a local Zotero account:
  `http://zotero.org/users/local/<8-char local library id>/items/<8-char item key>`
  — note the literal `local` segment and the 8-char library id `XXnzLBe9`
  (same in both fixtures).
- `schema` URL: `…/csl-citation.json` (NOT `csl-data.json`).
- JSON is serialized compact (no spaces after `:` / `,`) inside instrText.

> Evidence (S, SECTION B.1 "instrText JSON payload (pretty-printed)"): full JSON quoted above, incl. `"id": 1`, `"uris": ["http://zotero.org/users/local/XXnzLBe9/items/NU469JIQ"]`, `"schema": "https://github.com/citation-style-language/schema/raw/master/csl-citation.json"`.
> Evidence (R, SECTION B.1): `"citationID": "xO3paPuA"`, same `uris`, `"id": 1` (grep lines 83/88 of the R dump), same `schema`.

**3.5** `itemData` is full CSL-JSON. Observed key order: `id`, `type`, then
the remaining variables in **case-insensitive alphabetical order**
(`abstract`, `container-title`, `DOI`, `ISSN`, `issue`, `journalAbbreviation`,
`license`, `page`, `source`, `title`, `title-short`, `volume`), then
`author`, then `issued` last. `itemData.id` equals the citationItem `id`.
`issued.date-parts` mixes types: `[["2024", 6, 1]]` — **year is a string,
month and day are integers**.

> Evidence (S, SECTION B.1 JSON dump): key sequence exactly as listed; `"issued": {"date-parts": [["2024", 6, 1]]}`.

**3.6** Preview runs (between `separate` and `end`): the preview text
`(Havryk and Nazarova, 2024)` is split over three `w:r`
(`"("`, `"Havryk"`, `" and Nazarova, 2024)"`) with `w:proofErr` siblings —
a Word spell-check artifact; each run carries
`<w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri"/></w:rPr>`
and `w:rsidRPr="00B27459"`. **No `w:rStyle` on citation previews in either
fixture.**

> Evidence (S, SECTION E "run properties of the field RESULT runs"): three runs with the Calibri `rPr` shown verbatim.
> Evidence (R, SECTION E): identical three-run preview with identical `rPr`.

---

## 4. Bibliography field: `ADDIN ZOTERO_BIBL`

**4.1** Instruction text (leading/trailing spaces included):

```
" ADDIN ZOTERO_BIBL {\"uncited\":[],\"omitted\":[],\"custom\":[]} CSL_BIBLIOGRAPHY "
```

The JSON `{"uncited":[],"omitted":[],"custom":[]}` sits **between**
`ADDIN ZOTERO_BIBL` and the trailing token `CSL_BIBLIOGRAPHY`. Identical
in both fixtures.

> Evidence (S, SECTION B.2): `instrText concatenated (raw): ' ADDIN ZOTERO_BIBL {"uncited":[],"omitted":[],"custom":[]} CSL_BIBLIOGRAPHY '`, `trailing text after JSON: 'CSL_BIBLIOGRAPHY'`.
> Evidence (R, SECTION B.2): same string verbatim.

**4.2** Placement: the `begin`, `instrText` and `separate` runs **and the
first entry's text runs** live inside the first entry paragraph, which has
`<w:pPr><w:pStyle w:val="Bibliography"/></w:pPr>`. The `end` fldChar sits in
a **following, text-empty paragraph**. (Both fixtures contain exactly one
entry; multi-entry layout is an extrapolation — see §6.)

> Evidence (S, SECTION E): `paragraphs spanned by the field: #3 .. #4`; paragraph #3 `style='Bibliography'`, `w:pPr = <w:pPr><w:pStyle w:val="Bibliography"/></w:pPr>`; paragraph #4 `style=None text=''` with `w:pPr: (none)`; (S, SECTION B.2 field XML) shows `<w:p w14:paraId="4754A276" …><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>`.
> Evidence (R, SECTION E): `paragraphs spanned by the field: #72 .. #73`; #72 `style='Bibliography'`; #73 empty-text paragraph holding the end fldChar (its pPr there has direct `spacing/ind` overrides — Word-environment variance, see 4.5).

**4.3** The paragraph before the bibliography field is the heading
paragraph, plain text `References` with **no style and no pPr** in both
fixtures.

> Evidence (S, SECTION E): `paragraph BEFORE the field begin …: index=#2 style=None text='References'`, `its w:pPr: (none)`.
> Evidence (R): heading `References` at paragraph #53 (`style=-`); the field itself was appended after the typed list, preceded by empty paragraph #71.

**4.4** Bibliography preview runs: no `rPr` at all in the sample
(elsevier-harvard); in the real doc (harvard-cite-them-right) the journal-name
run carries `<w:rPr><w:i/><w:iCs/></w:rPr>` — i.e. **preview run formatting
is whatever the CSL style produces** (italics etc.), not a fixed shape.

> Evidence (S, SECTION E): `run text='Havryk' rPr=(no rPr)`, `run text=', A., Nazarova, T., 2024. …' rPr=(no rPr)`.
> Evidence (R, SECTION E): `run text='International Science Journal of Management, Economics & Fin' rPr=<w:rPr><w:i/><w:iCs/></w:rPr>`, other runs `rPr=(no rPr)`.

**4.5** `Bibliography` style definitions differ between the two Word
environments — both are real Zotero-saved documents:

- sample `styles.xml`: `basedOn Normal`, `next Normal`, `uiPriority 37`,
  `unhideWhenUsed`, `pPr = <w:spacing w:after="0" w:line="240" w:lineRule="auto"/><w:ind w:left="720" w:hanging="720"/>`
- real doc `styles.xml`: same skeleton but `pPr = <w:spacing w:after="240" w:line="240" w:lineRule="auto"/>` (no `ind`),
  and its end paragraph #73 carries direct overrides `<w:spacing w:after="80"/><w:ind w:left="198" w:hanging="198"/>`.

Canonical for Phase 3 (per rule 6, the sample wins): entry paragraphs get
`pStyle Bibliography`; if the style is missing from `styles.xml`, create it
with the sample's definition (hanging indent 720 twips).

> Evidence (S, SECTION E.2): full `<w:style w:type="paragraph" w:styleId="Bibliography">…<w:ind w:left="720" w:hanging="720"/>` definition.
> Evidence (R, SECTION E.2 + B.2): the variant definition and paragraph-level `ind left=198 hanging=198`.

**4.6** Bibliography entry text in the sample ends with a plain-text DOI
(`https://doi.org/10.46299/…`); there are **no `w:hyperlink` elements**
anywhere in the sample's `document.xml`.

> Evidence (S): `hyperlink elements in document.xml: 0` (paragraph-map check run during Phase 0); (S, SECTION B.2 field XML) DOI inside a plain `w:t`.

---

## 5. Document-level layout of the sample (paragraph map)

`word/document.xml` of the sample has exactly 5 paragraphs + `sectPr`:

```
#0 body text with the ZOTERO_ITEM field inline (no pPr)
#1 empty
#2 'References' heading (plain, no style)
#3 bibliography entry paragraph (pStyle Bibliography) — field begin/instr/separate + entry runs
#4 empty paragraph — field end fldChar
body last child: sectPr
```

> Evidence: paragraph map printed during Phase 0 (`#0 … len=216`, `#1 len=0`, `#2 'References'`, `#3 style='Bibliography' len=223`, `#4 len=0`, `body last child tag: sectPr`) + (S, SECTION B/E) location lines.

---

## 6. NOT shown by the fixtures (extrapolations — must not be claimed as sample-verified)

These structures are required by the project spec but absent from both
fixtures. Phase 3 will build them by generalizing the observed shapes;
each is listed here so nothing is silently invented:

1. **Grouped citation** (one field, several `citationItems`, e.g. `[2,5]`):
   extrapolated as repeated `{id, uris, itemData}` objects with `id`
   incrementing 1, 2, … per document (facts 3.4/3.5 give the per-item shape).
2. **`prefix` / `suffix` keys** in citationItems: not observed; Zotero omits
   them when empty — we will omit them too unless a citation needs them.
3. **`noteIndex` > 0** (citation inside a footnote/endnote): not observed;
   both fixtures show `noteIndex: 0` for body citations.
4. **Multi-entry bibliography field** (≥2 entries inside one ZOTERO_BIBL):
   not observed; extrapolated as one `w:p` per entry between `separate` and
   `end`, each with `pStyle Bibliography`, the end fldChar in the final
   empty paragraph (facts 4.2/4.5 give the single-entry shape).
5. **`ZOTERO_PREF_3`+ / chunks beyond two, `ZOTERO_PREF_0`**: not observed;
   numbering starts at `_1`, chunk size 255 (fact 2.4).
6. **`rStyle` on citation previews**: not observed in either fixture (fact 3.6).
7. **`uris` for group libraries** (`http://zotero.org/groups/<gid>/items/KEY`):
   not observed; only the `users/local/<libid>` form is sample-backed (fact 3.4).
8. **Fields inside footnotes.xml/endnotes.xml**: not observed (both fixtures
   have zero fields there — fact 1.2).

---

## 7. Idempotency facts from `real_vancouver.docx`

The input document is a **mixed** document — crucial for the conversion logic:

- 75 paragraphs; heading `References` at #53; typed Vancouver list `1.`–`16.`
  at #54–#70 (with an empty #55 and #71); body has 12 bracketed numeric
  citations: `[8]`(¶2), `[15]`(¶9), `[9,11]`(¶32), `[4]`(¶33), `[12]`(¶33),
  `[10,13]`(¶34), `[7]`(¶37), `[7]`(¶39), `[16]`(¶39), `[5]`(¶44), `[14]`(¶44), `[6]`(¶45).
- One citation was **already converted by Zotero**: paragraph #33 contains a
  ZOTERO_ITEM field rendering `(Havryk and Nazarova, 2024)` — and the typed
  `[4]` **still sits right next to it** in the same paragraph text
  (`"…skill requirements (Havryk and Nazarova, 2024) [4]."`).
- A ZOTERO_BIBL field with that one entry was appended at #72–#74, **after**
  the still-present typed list; `ZOTERO_PREF_*` already set to
  `harvard-cite-them-right`, `locale="en-US"`, session `88ppEdET`.
- Issue numbers inside the typed reference list produce `(1)`/`(3)`/`(2)`/`(6)`
  pattern matches (e.g. `2023;4(1):47-67`) — confirming the detector must
  never scan the reference-list region, and `(n)` in body text needs the
  spec's strict conditions.

> Evidence: (R, SECTION B.1/B.2/C/F) + paragraph map and bracketed-number scan printed during Phase 0 (¶33 text `'…(Havryk and Nazarova, 2024) [4]…'`; list `#54 '1. Ünvan YA…'` … `#70 '16. Revolut…'`).

**Consequences for Phase 1/3:** existing Zotero fields must be detected and
skipped (reported), the surrounding typed `[4]` must still be converted,
prefs must be **updated in place** (names `ZOTERO_PREF_1/2`, pids 2/3), and
the typed list must be replaced by the bibliography field while the
appended one is reconciled (not duplicated).

---

## 8. Exact strings Phase 3 must reproduce (copy-paste constants)

| what | exact value (from fixtures) |
|---|---|
| citation instr prefix | ` ADDIN ZOTERO_ITEM CSL_CITATION ` (+ compact JSON + trailing space) |
| bibliography instr | ` ADDIN ZOTERO_BIBL {"uncited":[],"omitted":[],"custom":[]} CSL_BIBLIOGRAPHY ` |
| citation schema URL | `https://github.com/citation-style-language/schema/raw/master/csl-citation.json` |
| custom-property fmtid | `{D5CDD505-2E9C-101B-9397-08002B2CF9AE}` |
| pref value element | `<vt:lpwstr>` (xmlns:vt = `http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes`) |
| pref data root | `<data data-version="3" zotero-version="…">` |
| content-type for custom.xml | `application/vnd.openxmlformats-officedocument.custom-properties+xml` |
| rel type for custom.xml | `http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties` |
| local-library URI form | `http://zotero.org/users/local/<libraryID>/items/<itemKey>` |
| Bibliography style pPr (sample, canonical) | `<w:spacing w:after="0" w:line="240" w:lineRule="auto"/><w:ind w:left="720" w:hanging="720"/>` |
