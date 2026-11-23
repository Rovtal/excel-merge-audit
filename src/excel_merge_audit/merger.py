"""Multi-table merge engine with exact deduplication and conflicting-document audit."""

from decimal import Decimal
from pathlib import Path
from typing import Optional

from .i18n import STATUS_LABELS, t, validate_lang
from .models import (
    AuditIssue,
    MergeReport,
    MergeSummary,
    NormalizedRecord,
    SourceFileStat,
)
from .normalizer import SCIENTIFIC_RE, _normalize_id_token, _normalize_party_token, normalize_table
from .readers import load_input_tables


def merge_and_audit(
    inputs: list[Path],
    only_sheet: Optional[str] = None,
    keep_duplicates: bool = False,
    lang: str = "tr",
) -> MergeReport:
    """Load multiple .xlsx/.csv/.tsv inputs, normalize schemas, deduplicate, and reconcile."""
    active_lang = validate_lang(lang)
    st = STATUS_LABELS[active_lang]
    tables, ignored = load_input_tables(inputs, only_sheet=only_sheet, lang=active_lang)
    if not tables:
        return MergeReport(
            records=[],
            removed_duplicates=[],
            issues=[],
            source_stats=[],
            summary=MergeSummary(
                source_sheets=0,
                total_input_rows=0,
                kept_rows=0,
                exact_duplicates_removed=0,
                conflicting_documents=0,
                passed_rows=0,
                warning_rows=0,
                error_rows=0,
                raw_brut_total=Decimal("0.00"),
                kept_net_total=Decimal("0.00"),
                kept_kdv_total=Decimal("0.00"),
                kept_brut_total=Decimal("0.00"),
                duplicate_brut_removed=Decimal("0.00"),
            ),
            ignored_files=ignored,
            fatal_reason=t(active_lang, "merge_no_tables"),
            lang=active_lang,
        )

    all_issues: list[AuditIssue] = []
    seen_exact: dict[str, NormalizedRecord] = {}
    kept_records: list[NormalizedRecord] = []
    removed_duplicates: list[NormalizedRecord] = []
    source_stats: list[SourceFileStat] = []

    raw_brut_total = Decimal("0.00")
    duplicate_brut_removed = Decimal("0.00")
    total_input_rows = 0

    for table in tables:
        recs, tbl_issues, hdr_row, col_labels = normalize_table(table, lang=active_lang)
        all_issues.extend(tbl_issues)

        file_name = Path(table.source_file).name
        file_raw_brut = Decimal("0.00")
        file_kept_brut = Decimal("0.00")
        file_kept_count = 0
        file_dup_count = 0
        file_err = sum(1 for iss in tbl_issues if iss.severity == st["error"])
        file_warn = sum(1 for iss in tbl_issues if iss.severity == st["warn"])

        for rec in recs:
            total_input_rows += 1
            file_raw_brut += rec.brut_tutar
            raw_brut_total += rec.brut_tutar

            if not keep_duplicates and rec.dedupe_key in seen_exact:
                orig = seen_exact[rec.dedupe_key]
                orig_name = Path(orig.source_file).name
                rec.extra_note = t(
                    active_lang,
                    "merge_dup_note",
                    orig_name=orig_name,
                    sheet_name=orig.sheet_name,
                    source_row=orig.source_row,
                )
                removed_duplicates.append(rec)
                duplicate_brut_removed += rec.brut_tutar
                file_dup_count += 1
                file_warn += 1
                all_issues.append(
                    AuditIssue(
                        severity=st["warn"],
                        code="TAM_MUKERRER_SILINDI",
                        source_file=rec.source_file,
                        sheet_name=rec.sheet_name,
                        source_row=rec.source_row,
                        belge_no=rec.belge_no or t(active_lang, "norm_row_prefix", row=rec.source_row),
                        message=t(
                            active_lang,
                            "merge_dup_msg",
                            brut=rec.brut_tutar,
                            currency=rec.para_birimi,
                            orig_name=orig_name,
                            sheet_name=orig.sheet_name,
                            source_row=orig.source_row,
                        ),
                    )
                )
                continue

            seen_exact.setdefault(rec.dedupe_key, rec)
            kept_records.append(rec)
            file_kept_brut += rec.brut_tutar
            file_kept_count += 1

        source_stats.append(
            SourceFileStat(
                source_file=file_name,
                sheet_name=table.sheet_name,
                file_format=table.file_format,
                encoding=table.encoding,
                delimiter=table.delimiter,
                header_row=hdr_row,
                raw_data_rows=len(recs),
                kept_rows=file_kept_count,
                duplicate_skipped=file_dup_count,
                error_count=file_err,
                warning_count=file_warn,
                raw_brut_sum=file_raw_brut,
                kept_brut_sum=file_kept_brut,
                column_map=col_labels,
            )
        )

    # Detect same belge_no + matching counterparty appearing with conflicting amounts or dates
    doc_groups: dict[tuple[str, str], list[NormalizedRecord]] = {}
    for rec in kept_records:
        doc_tok = _normalize_id_token(rec.belge_no)
        if not doc_tok:
            continue
        party_tok = _normalize_party_token(rec.taraf) or _normalize_id_token(rec.vergi_no)
        doc_groups.setdefault((doc_tok, party_tok), []).append(rec)

    conflicting_docs = 0
    for (doc_tok, _party_tok), group in doc_groups.items():
        if len(group) < 2:
            continue
        # Ignore scientific-notation corrupted VKN values when checking issuer uniqueness
        vkn_set = {
            _normalize_id_token(g.vergi_no)
            for g in group
            if _normalize_id_token(g.vergi_no) and not SCIENTIFIC_RE.match(g.vergi_no)
        }
        if len(vkn_set) > 1:
            continue
        unique_signatures = {(g.tarih, g.brut_tutar, g.para_birimi) for g in group}
        if len(unique_signatures) > 1:
            conflicting_docs += 1
            for g in group:
                g.durum = st["error"]
                all_issues.append(
                    AuditIssue(
                        severity=st["error"],
                        code="CAKISAN_BELGE",
                        source_file=g.source_file,
                        sheet_name=g.sheet_name,
                        source_row=g.source_row,
                        belge_no=g.belge_no,
                        message=t(
                            active_lang,
                            "merge_conflict_msg",
                            doc_tok=doc_tok,
                            tarih=g.tarih,
                            brut=g.brut_tutar,
                            currency=g.para_birimi,
                        ),
                    )
                )

    kept_records.sort(key=lambda r: (r.tarih or "9999-99-99", r.belge_no, Path(r.source_file).name, r.source_row))
    severity_rank = {st["error"]: 0, st["warn"]: 1, st["info"]: 2}
    all_issues.sort(key=lambda i: (severity_rank.get(i.severity, 9), Path(i.source_file).name, i.source_row))

    kept_net_total = sum((r.net_tutar for r in kept_records), Decimal("0.00"))
    kept_kdv_total = sum((r.kdv_tutar for r in kept_records), Decimal("0.00"))
    kept_brut_total = sum((r.brut_tutar for r in kept_records), Decimal("0.00"))

    summary = MergeSummary(
        source_sheets=len(tables),
        total_input_rows=total_input_rows,
        kept_rows=len(kept_records),
        exact_duplicates_removed=len(removed_duplicates),
        conflicting_documents=conflicting_docs,
        passed_rows=sum(1 for r in kept_records if r.durum == st["pass"]),
        warning_rows=sum(1 for r in kept_records if r.durum == st["warn"]),
        error_rows=sum(1 for r in kept_records if r.durum == st["error"]),
        raw_brut_total=raw_brut_total,
        kept_net_total=kept_net_total,
        kept_kdv_total=kept_kdv_total,
        kept_brut_total=kept_brut_total,
        duplicate_brut_removed=duplicate_brut_removed,
    )

    return MergeReport(
        records=kept_records,
        removed_duplicates=removed_duplicates,
        issues=all_issues,
        source_stats=source_stats,
        summary=summary,
        ignored_files=ignored,
        lang=active_lang,
    )
