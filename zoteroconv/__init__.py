"""zoteroconv -- convert numeric .docx citations into live Zotero Word fields.

Phase 1 modules:
    docx_io               -- DocxPackage: load/save .docx with part tracking
    zotero_fields         -- discovery/parsing of existing Zotero Word fields
    parse_reference_list  -- locate + parse the typed/auto-numbered reference list
    detect_citations      -- find numeric in-text citations with exact spans
"""

__version__ = "0.1.0"
