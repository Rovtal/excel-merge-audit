# Security Policy

## Threat Model

`excel-merge-audit` processes untrusted `.xlsx`, `.csv`, and `.tsv` spreadsheets strictly offline:

- **Offline-Only Execution:** Zero external network calls or telemetry.
- **OpenXML XXE & Archive Bounds:** Direct `zipfile + lxml` parsing uses `XMLParser(resolve_entities=False, no_network=True, huge_tree=False)` with zip-slip path rejection, 50 MB per-file cap, 2,000 archive entry cap, and 150 MB uncompressed limit.
- **Unsupported Legacy Formats Rejected:** Binary `.xls` (`OLE2`) and encrypted workbooks fail safely with clear guidance.
- **CSV / Excel Formula Injection (CWE-1236):** Both `.xlsx` and `--csv-out` exports prefix leading `=`, `+`, `-`, `@`, `\t`, and `\r` characters with `'`, and `XlsxWriter` disables `strings_to_formulas` and `strings_to_urls`.

## Reporting a Vulnerability

Report security bugs privately via GitHub security advisories without uploading sensitive client spreadsheets.
