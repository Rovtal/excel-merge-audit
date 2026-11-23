# excel-merge-audit

[![CI](https://github.com/Rovtal/excel-merge-audit/actions/workflows/ci.yml/badge.svg)](https://github.com/Rovtal/excel-merge-audit/actions/workflows/ci.yml)

Language: **English** | [Türkçe](README.tr.md)

A **100% offline** Python CLI engine that consolidates heterogeneous `.xlsx`, `.csv`, and `.tsv` spreadsheets, auto-maps Turkish/English column schemas, normalizes mixed number and date formats, isolates exact and conflicting duplicate documents, and exports a formula-verified 4-sheet Excel reconciliation workbook.

> Designed for strict local data privacy: spreadsheets never leave your machine. No `openpyxl` or `pandas` dependency is required; OpenXML `.xlsx` files are parsed directly via security-hardened `zipfile + lxml`, and workbooks are written with `XlsxWriter`.

## What It Solves

- **Encoding & Delimiter Chaos:** Auto-detects `UTF-8`, `UTF-8 BOM`, `cp1254` (Turkish Windows ANSI), `ISO-8859-9`, and `cp1252` files with `;`, `,`, `\t`, and `|` delimiters.
- **Top Banners, Sparse Cells & Serial Dates:** Scans rows `1–15` to locate real column headers, skips repeated `TOTAL` / `GRAND TOTAL` summary rows, prevents sparse OpenXML empty-cell column shifting, and converts Excel `1900`/`1904` serial dates to `YYYY-MM-DD`.
- **Strict `Decimal` Monetary Parsing:** Parses `1.250,50 TL`, accounting negatives like `(412,50)`, and `1,250.50 USD` without floating-point drift; flags `Net + VAT != Gross` discrepancies as `Error`.
- **Exact Duplicate & Conflicting Document Detection:** Removes exact duplicates while tracking excluded totals in the reconciliation sheet, and flags same-document rows with mismatched amounts or dates as `Error`.
- **Security Hardening:** Guarded against OpenXML **XXE**, **zip-slip (`../`)**, archive entry/size bombs, and Excel formula injection (**CWE-1236** `=CMD|...`).

## Scope Boundaries

- **Supported:** `.xlsx` (OpenXML), `.csv`, `.tsv` files and nested folders (temporary `~$` lock files are safely skipped).
- **Out of Scope by Design:** Legacy binary `.xls` (`Excel 97-2003`) and password-protected/encrypted `.xlsx` files are intentionally rejected with a clear conversion prompt.

## Installation & Quick Demo

In a standard Python (`>=3.10`) environment:

```bash
python -m pip install -e .
excel-merge demo --lang en --out-dir demo_output
```

Run directly from source without installing (PowerShell / Bash):

```powershell
$env:PYTHONPATH="src"; python -m excel_merge_audit.cli demo --lang en --out-dir demo_output
```

```bash
PYTHONPATH=src python -m excel_merge_audit.cli demo --lang en --out-dir demo_output
```

## Merging Your Own Files

Merge a folder or explicit list of files in English (`--lang en`) or Turkish (`--lang tr`, default for backward compatibility). Use `--csv-out` to also export a CWE-1236 sanitized UTF-8 CSV of clean rows, and `--fail-on-error` to return exit code `2` when critical errors are found:

```bash
excel-merge merge ./files -o merged_report.xlsx --json summary.json --csv-out clean.csv --lang en --fail-on-error
```

Use `--sheet "Sheet1"` to target a single worksheet inside `.xlsx` workbooks, or `--keep-duplicates` to keep exact duplicates in the merged output.

## 4-Sheet Workbook Architecture (`--lang en`)

| Sheet Name | Contents |
| --- | --- |
| `Merge_Summary` | Dynamic `COUNTIF` / `SUM` formulas and zero-drift check (`Raw Gross - Removed Duplicate Gross - Clean Gross = 0.00`) |
| `Clean_Merged` | Normalized dates, document IDs, counterparties, tax IDs, Net/VAT/Gross amounts, currency, and per-row status |
| `Source_Files` | Per-sheet encoding, delimiter, header row index, kept/duplicate counts, and gross totals |
| `Duplicates_and_Warnings` | Excluded exact duplicates, conflicting document groups, scientific-notation ID warnings, and formula error cells |

## License

[MIT](LICENSE) — Rovtal
