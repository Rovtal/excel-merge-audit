"""Styled formula-backed .xlsx exporter with CWE-1236 formula injection protection."""

import csv
from pathlib import Path
from typing import Optional

import xlsxwriter

from .i18n import EXPORT_COPY, SHEET_NAMES, STATUS_LABELS, validate_lang
from .models import MergeReport


FORMULA_PREFIXES = ("=", "+", "-", "@", "\t", "\r")


def sanitize_excel_text(value: str) -> str:
    """Neutralize spreadsheet formula injection vectors in untrusted text cells."""
    if not value:
        return ""
    if value[0] in FORMULA_PREFIXES:
        return "'" + value
    return value


def export_merge_workbook(
    report: MergeReport,
    output_path: Path,
    lang: Optional[str] = None,
) -> Path:
    """Generate a 4-sheet formatted and formula-verified .xlsx merge audit workbook."""
    active_lang = validate_lang(lang or report.lang)
    sheets = SHEET_NAMES[active_lang]
    st = STATUS_LABELS[active_lang]
    copy = EXPORT_COPY[active_lang]
    metrics = copy["metrics"]
    clean_sheet = sheets["clean"]

    resolved = output_path.resolve()
    resolved.parent.mkdir(parents=True, exist_ok=True)

    wb = xlsxwriter.Workbook(
        str(resolved),
        {"strings_to_formulas": False, "strings_to_urls": False},
    )
    try:
        hdr_fmt = wb.add_format(
            {
                "bold": True,
                "bg_color": "#0F172A",
                "font_color": "#F8FAFC",
                "border": 1,
                "align": "center",
                "valign": "vcenter",
            }
        )
        title_fmt = wb.add_format({"bold": True, "font_size": 14, "font_color": "#0F172A"})
        label_fmt = wb.add_format({"bold": True, "bg_color": "#F1F5F9", "border": 1})
        num_int_fmt = wb.add_format({"num_format": "#,##0", "border": 1})
        money_fmt = wb.add_format({"num_format": "#,##0.00", "border": 1})
        text_fmt = wb.add_format({"border": 1, "valign": "vcenter"})
        ok_fmt = wb.add_format({"bold": True, "font_color": "#065F46", "bg_color": "#D1FAE5", "border": 1, "align": "center"})
        warn_fmt = wb.add_format({"bold": True, "font_color": "#92400E", "bg_color": "#FEF3C7", "border": 1, "align": "center"})
        err_fmt = wb.add_format({"bold": True, "font_color": "#991B1B", "bg_color": "#FEE2E2", "border": 1, "align": "center"})

        # 1. Summary Sheet
        ws_sum = wb.add_worksheet(sheets["summary"])
        ws_sum.hide_gridlines(2)
        ws_sum.set_column("A:A", 36)
        ws_sum.set_column("B:B", 22)
        ws_sum.set_column("C:C", 46)
        ws_sum.write(0, 0, copy["title"], title_fmt)

        for col_i, header_text in enumerate(copy["sum_headers"]):
            ws_sum.write(2, col_i, header_text, hdr_fmt)

        last_rec_row = len(report.records) + 1
        sum_range_valid = last_rec_row >= 2

        row_map: dict[str, str] = {}
        summary_specs: list[tuple[str, object, object]] = [
            ("source_sheets", report.summary.source_sheets, num_int_fmt),
            ("raw_rows", report.summary.total_input_rows, num_int_fmt),
            (
                "clean_rows",
                f'=COUNTIF({clean_sheet}!M2:M{last_rec_row},"<>")' if sum_range_valid else 0,
                num_int_fmt,
            ),
            ("dup_rows", report.summary.exact_duplicates_removed, num_int_fmt),
            ("conflict_groups", report.summary.conflicting_documents, num_int_fmt),
            (
                "passed_rows",
                f'=COUNTIF({clean_sheet}!M2:M{last_rec_row},"{st["pass"]}")' if sum_range_valid else 0,
                ok_fmt,
            ),
            (
                "warning_rows",
                f'=COUNTIF({clean_sheet}!M2:M{last_rec_row},"{st["warn"]}")' if sum_range_valid else 0,
                warn_fmt,
            ),
            (
                "error_rows",
                f'=COUNTIF({clean_sheet}!M2:M{last_rec_row},"{st["error"]}")' if sum_range_valid else 0,
                err_fmt,
            ),
            ("raw_brut", float(report.summary.raw_brut_total), money_fmt),
            ("dup_brut", float(report.summary.duplicate_brut_removed), money_fmt),
            (
                "clean_net",
                f"=SUM({clean_sheet}!I2:I{last_rec_row})" if sum_range_valid else 0.0,
                money_fmt,
            ),
            (
                "clean_kdv",
                f"=SUM({clean_sheet}!J2:J{last_rec_row})" if sum_range_valid else 0.0,
                money_fmt,
            ),
            (
                "clean_brut",
                f"=SUM({clean_sheet}!K2:K{last_rec_row})" if sum_range_valid else 0.0,
                money_fmt,
            ),
        ]

        start_row = 3
        for idx, (key, _val, _fmt) in enumerate(summary_specs):
            row_map[key] = f"B{start_row + idx + 1}"

        summary_specs.append(
            (
                "recon_diff",
                f"={row_map['raw_brut']}-{row_map['dup_brut']}-{row_map['clean_brut']}",
                money_fmt,
            )
        )

        for offset, (key, val, val_fmt) in enumerate(summary_specs, start=start_row):
            lbl, desc = metrics[key]
            ws_sum.write_string(offset, 0, lbl, label_fmt)
            if isinstance(val, str) and val.startswith("="):
                ws_sum.write_formula(offset, 1, val, val_fmt)
            elif isinstance(val, int):
                ws_sum.write_number(offset, 1, val, val_fmt)
            else:
                ws_sum.write_number(offset, 1, float(val), val_fmt)
            ws_sum.write_string(offset, 2, desc, text_fmt)

        # 2. Clean Merged Records Sheet
        ws_clean = wb.add_worksheet(clean_sheet)
        clean_headers = copy["clean_headers"]
        for c_i, h in enumerate(clean_headers):
            ws_clean.write_string(0, c_i, h, hdr_fmt)
        ws_clean.freeze_panes(1, 0)
        if report.records:
            ws_clean.autofilter(0, 0, len(report.records), len(clean_headers) - 1)

        col_widths_clean = [24, 14, 13, 13, 18, 28, 16, 30, 15, 14, 15, 11, 12]
        for idx, width in enumerate(col_widths_clean):
            ws_clean.set_column(idx, idx, width)

        status_fmt_map = {st["pass"]: ok_fmt, st["warn"]: warn_fmt, st["error"]: err_fmt}
        for r_idx, rec in enumerate(report.records, start=1):
            ws_clean.write_string(r_idx, 0, sanitize_excel_text(Path(rec.source_file).name), text_fmt)
            ws_clean.write_string(r_idx, 1, sanitize_excel_text(rec.sheet_name), text_fmt)
            ws_clean.write_number(r_idx, 2, rec.source_row, num_int_fmt)
            ws_clean.write_string(r_idx, 3, sanitize_excel_text(rec.tarih), text_fmt)
            ws_clean.write_string(r_idx, 4, sanitize_excel_text(rec.belge_no), text_fmt)
            ws_clean.write_string(r_idx, 5, sanitize_excel_text(rec.taraf), text_fmt)
            ws_clean.write_string(r_idx, 6, sanitize_excel_text(rec.vergi_no), text_fmt)
            ws_clean.write_string(r_idx, 7, sanitize_excel_text(rec.aciklama), text_fmt)
            ws_clean.write_number(r_idx, 8, float(rec.net_tutar), money_fmt)
            ws_clean.write_number(r_idx, 9, float(rec.kdv_tutar), money_fmt)
            ws_clean.write_number(r_idx, 10, float(rec.brut_tutar), money_fmt)
            ws_clean.write_string(r_idx, 11, sanitize_excel_text(rec.para_birimi), text_fmt)
            ws_clean.write_string(r_idx, 12, rec.durum, status_fmt_map.get(rec.durum, text_fmt))

        # 3. Source Files Sheet
        ws_src = wb.add_worksheet(sheets["sources"])
        src_headers = copy["source_headers"]
        for c_i, h in enumerate(src_headers):
            ws_src.write_string(0, c_i, h, hdr_fmt)
        ws_src.freeze_panes(1, 0)
        for idx, width in enumerate([26, 14, 10, 14, 10, 13, 12, 13, 17, 10, 10, 18, 18]):
            ws_src.set_column(idx, idx, width)

        for r_idx, st_row in enumerate(report.source_stats, start=1):
            ws_src.write_string(r_idx, 0, sanitize_excel_text(st_row.source_file), text_fmt)
            ws_src.write_string(r_idx, 1, sanitize_excel_text(st_row.sheet_name), text_fmt)
            ws_src.write_string(r_idx, 2, st_row.file_format.upper(), text_fmt)
            ws_src.write_string(r_idx, 3, st_row.encoding, text_fmt)
            ws_src.write_string(r_idx, 4, sanitize_excel_text(st_row.delimiter or "-"), text_fmt)
            ws_src.write_number(r_idx, 5, st_row.header_row, num_int_fmt)
            ws_src.write_number(r_idx, 6, st_row.raw_data_rows, num_int_fmt)
            ws_src.write_number(r_idx, 7, st_row.kept_rows, num_int_fmt)
            ws_src.write_number(r_idx, 8, st_row.duplicate_skipped, num_int_fmt)
            ws_src.write_number(r_idx, 9, st_row.warning_count, num_int_fmt)
            ws_src.write_number(r_idx, 10, st_row.error_count, num_int_fmt)
            ws_src.write_number(r_idx, 11, float(st_row.raw_brut_sum), money_fmt)
            ws_src.write_number(r_idx, 12, float(st_row.kept_brut_sum), money_fmt)

        # 4. Duplicates and Warnings Sheet
        ws_iss = wb.add_worksheet(sheets["issues"])
        iss_headers = copy["issue_headers"]
        for c_i, h in enumerate(iss_headers):
            ws_iss.write_string(0, c_i, h, hdr_fmt)
        ws_iss.freeze_panes(1, 0)
        for idx, width in enumerate([12, 24, 24, 14, 10, 18, 68]):
            ws_iss.set_column(idx, idx, width)

        for r_idx, iss in enumerate(report.issues, start=1):
            sev_fmt = status_fmt_map.get(iss.severity, text_fmt)
            ws_iss.write_string(r_idx, 0, iss.severity, sev_fmt)
            ws_iss.write_string(r_idx, 1, iss.code, text_fmt)
            ws_iss.write_string(r_idx, 2, sanitize_excel_text(Path(iss.source_file).name), text_fmt)
            ws_iss.write_string(r_idx, 3, sanitize_excel_text(iss.sheet_name), text_fmt)
            ws_iss.write_number(r_idx, 4, iss.source_row, num_int_fmt)
            ws_iss.write_string(r_idx, 5, sanitize_excel_text(iss.belge_no), text_fmt)
            ws_iss.write_string(r_idx, 6, sanitize_excel_text(iss.message), text_fmt)
    finally:
        wb.close()

    return resolved


def export_clean_csv(
    report: MergeReport,
    output_path: Path,
    lang: Optional[str] = None,
) -> Path:
    """Export deduplicated Clean_Merged records as CWE-1236 sanitized UTF-8 CSV."""
    active_lang = validate_lang(lang or report.lang)
    clean_headers = EXPORT_COPY[active_lang]["clean_headers"]
    resolved = output_path.resolve()
    resolved.parent.mkdir(parents=True, exist_ok=True)

    with resolved.open("w", encoding="utf-8", newline="") as fh:
        writer = csv.writer(fh)
        writer.writerow(clean_headers)
        for rec in report.records:
            writer.writerow(
                [
                    sanitize_excel_text(Path(rec.source_file).name),
                    sanitize_excel_text(rec.sheet_name),
                    rec.source_row,
                    sanitize_excel_text(rec.tarih),
                    sanitize_excel_text(rec.belge_no),
                    sanitize_excel_text(rec.taraf),
                    sanitize_excel_text(rec.vergi_no),
                    sanitize_excel_text(rec.aciklama),
                    f"{rec.net_tutar:.2f}",
                    f"{rec.kdv_tutar:.2f}",
                    f"{rec.brut_tutar:.2f}",
                    sanitize_excel_text(rec.para_birimi),
                    rec.durum,
                ]
            )
    return resolved
