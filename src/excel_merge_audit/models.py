"""Typed domain models for multi-file spreadsheet loading, normalization, and audit."""

from dataclasses import dataclass, field
from decimal import Decimal
from typing import Optional


@dataclass(slots=True)
class RawSheetTable:
    """Raw 2D cell grid read from a single spreadsheet sheet or delimited file."""

    source_file: str
    sheet_name: str
    file_format: str
    encoding: str
    delimiter: str
    rows: list[list[str]]
    formula_error_cells: list[tuple[int, int, str]] = field(default_factory=list)


@dataclass(slots=True)
class AuditIssue:
    """Structured finding emitted during parsing, normalization, or deduplication."""

    severity: str  # "HATA" | "UYARI" | "BILGI"
    code: str
    source_file: str
    sheet_name: str
    source_row: int
    belge_no: str
    message: str


@dataclass(slots=True)
class NormalizedRecord:
    """Canonical accounting/tabular record aligned across heterogeneous source schemas."""

    source_file: str
    sheet_name: str
    source_row: int
    tarih: str
    belge_no: str
    taraf: str
    vergi_no: str
    aciklama: str
    net_tutar: Decimal
    kdv_tutar: Decimal
    brut_tutar: Decimal
    para_birimi: str
    durum: str  # "GECTI" | "UYARI" | "HATA"
    dedupe_key: str
    extra_note: str = ""


@dataclass(slots=True)
class SourceFileStat:
    """Per-input-sheet diagnostic and reconciliation counters."""

    source_file: str
    sheet_name: str
    file_format: str
    encoding: str
    delimiter: str
    header_row: int
    raw_data_rows: int
    kept_rows: int
    duplicate_skipped: int
    error_count: int
    warning_count: int
    raw_brut_sum: Decimal
    kept_brut_sum: Decimal
    column_map: dict[str, str] = field(default_factory=dict)


@dataclass(slots=True)
class MergeSummary:
    """Aggregate mathematical reconciliation and quality metrics."""

    source_sheets: int
    total_input_rows: int
    kept_rows: int
    exact_duplicates_removed: int
    conflicting_documents: int
    passed_rows: int
    warning_rows: int
    error_rows: int
    raw_brut_total: Decimal
    kept_net_total: Decimal
    kept_kdv_total: Decimal
    kept_brut_total: Decimal
    duplicate_brut_removed: Decimal


@dataclass(slots=True)
class MergeReport:
    """Complete output of a multi-file merge and audit execution."""

    records: list[NormalizedRecord]
    removed_duplicates: list[NormalizedRecord]
    issues: list[AuditIssue]
    source_stats: list[SourceFileStat]
    summary: MergeSummary
    ignored_files: list[str] = field(default_factory=list)
    fatal_reason: Optional[str] = None
