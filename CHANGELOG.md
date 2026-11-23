# Changelog

## 0.2.0 - 2026-11-23

- Added `--csv-out` CLI option and `export_clean_csv()` to write CWE-1236 sanitized UTF-8 CSV outputs of `Clean_Merged` rows with localized headers and exact `Decimal` amounts.
- Added `-V` / `--version` CLI flag.
- Added GitHub Actions CI matrix (`Python 3.10–3.12`) running `ruff` strict checks, `pytest`, and EN demo CLI smoke checks.
- Added `SECURITY.md` documenting offline XXE, zip-slip, archive limit, and CWE-1236 protections.

## 0.1.0 - 2026-11-23

- Added bilingual (`--lang en` / `--lang tr`) normalization/export support, English documentation, and `normalizer.py` closure fix.
