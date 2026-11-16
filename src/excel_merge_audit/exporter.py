"""Styled formula-backed .xlsx exporter with CWE-1236 formula injection protection."""

from pathlib import Path

import xlsxwriter

from .models import MergeReport


FORMULA_PREFIXES = ("=", "+", "-", "@", "\t", "\r")


def sanitize_excel_text(value: str) -> str:
    """Neutralize spreadsheet formula injection vectors in untrusted text cells."""
    if not value:
        return ""
    if value[0] in FORMULA_PREFIXES:
        return "'" + value
    return value


def export_merge_workbook(report: MergeReport, output_path: Path) -> Path:
    """Generate a 4-sheet formatted and formula-verified .xlsx merge audit workbook."""
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

        # 1. Sheet: Ozet_Birlestirme
        ws_sum = wb.add_worksheet("Ozet_Birlestirme")
        ws_sum.hide_gridlines(2)
        ws_sum.set_column("A:A", 36)
        ws_sum.set_column("B:B", 22)
        ws_sum.set_column("C:C", 42)
        ws_sum.write(0, 0, "Çoklu Excel / CSV Birleştirme ve Denetim Özeti", title_fmt)

        ws_sum.write(2, 0, "Özet Göstergesi", hdr_fmt)
        ws_sum.write(2, 1, "Değer", hdr_fmt)
        ws_sum.write(2, 2, "Açıklama / Doğrulama Kuralı", hdr_fmt)

        last_rec_row = len(report.records) + 1
        sum_range_valid = last_rec_row >= 2

        row_map: dict[str, str] = {}
        summary_rows = [
            ("source_sheets", "Kaynak Tablo Sayısı", report.summary.source_sheets, num_int_fmt, "İşlenen sayfa ve CSV/TSV dosyaları"),
            ("raw_rows", "Ham Girdi Satırı", report.summary.total_input_rows, num_int_fmt, "Mükerrer temizliği öncesi okunan veri satırı"),
            (
                "clean_rows",
                "Temiz Birleştirilmiş Satır",
                f'=COUNTIF(Temiz_Birlestirilmis!M2:M{last_rec_row},"<>")' if sum_range_valid else 0,
                num_int_fmt,
                "Temiz_Birlestirilmis sayfasındaki net kayıt adedi",
            ),
            ("dup_rows", "Çıkarılan Birebir Mükerrer", report.summary.exact_duplicates_removed, num_int_fmt, "Mukerrer_ve_Uyarilar sayfasında listelenir"),
            ("conflict_groups", "Çakışan Belge Grubu", report.summary.conflicting_documents, num_int_fmt, "Aynı belge numarasıyla farklı tutar/tarih içerenler"),
            (
                "passed_rows",
                "Durum: Geçti (En İyi)",
                f'=COUNTIF(Temiz_Birlestirilmis!M2:M{last_rec_row},"Geçti")' if sum_range_valid else 0,
                ok_fmt,
                "Matematiksel ve format doğrulamasından geçen satırlar",
            ),
            (
                "warning_rows",
                "Durum: Uyarı",
                f'=COUNTIF(Temiz_Birlestirilmis!M2:M{last_rec_row},"Uyarı")' if sum_range_valid else 0,
                warn_fmt,
                "Tarih, eksik alan veya bilimsel gösterim uyarısı alanlar",
            ),
            (
                "error_rows",
                "Durum: Hata (Kritik)",
                f'=COUNTIF(Temiz_Birlestirilmis!M2:M{last_rec_row},"Hata")' if sum_range_valid else 0,
                err_fmt,
                "Net+KDV uyumsuzluğu, çakışan belge veya #REF!/#DEĞER! hatası",
            ),
            ("raw_brut", "Ham Girdi Brüt Toplam", float(report.summary.raw_brut_total), money_fmt, "Mükerrer temizliği öncesi brüt toplam"),
            ("dup_brut", "Çıkarılan Mükerrer Brüt", float(report.summary.duplicate_brut_removed), money_fmt, "Silinen mükerrer satırların brüt toplamı"),
            (
                "clean_net",
                "Temiz Net Toplam",
                f"=SUM(Temiz_Birlestirilmis!I2:I{last_rec_row})" if sum_range_valid else 0.0,
                money_fmt,
                "Temiz_Birlestirilmis sütun I toplamı",
            ),
            (
                "clean_kdv",
                "Temiz KDV Toplam",
                f"=SUM(Temiz_Birlestirilmis!J2:J{last_rec_row})" if sum_range_valid else 0.0,
                money_fmt,
                "Temiz_Birlestirilmis sütun J toplamı",
            ),
            (
                "clean_brut",
                "Temiz Brüt Toplam",
                f"=SUM(Temiz_Birlestirilmis!K2:K{last_rec_row})" if sum_range_valid else 0.0,
                money_fmt,
                "Mutabakat: Ham Brüt - Mükerrer Brüt eşit olmalıdır",
            ),
        ]

        start_row = 3
        for idx, (key, _lbl, _val, _fmt, _desc) in enumerate(summary_rows):
            row_map[key] = f"B{start_row + idx + 1}"

        summary_rows.append(
            (
                "recon_diff",
                "Mutabakat Sapması",
                f"={row_map['raw_brut']}-{row_map['dup_brut']}-{row_map['clean_brut']}",
                money_fmt,
                "Beklenen değer 0,00 (Ham - Çıkarılan - Temiz)",
            )
        )

        for offset, (_key, lbl, val, val_fmt, desc) in enumerate(summary_rows, start=start_row):
            ws_sum.write_string(offset, 0, lbl, label_fmt)
            if isinstance(val, str) and val.startswith("="):
                ws_sum.write_formula(offset, 1, val, val_fmt)
            elif isinstance(val, int):
                ws_sum.write_number(offset, 1, val, val_fmt)
            else:
                ws_sum.write_number(offset, 1, float(val), val_fmt)
            ws_sum.write_string(offset, 2, desc, text_fmt)

        # 2. Sheet: Temiz_Birlestirilmis
        ws_clean = wb.add_worksheet("Temiz_Birlestirilmis")
        clean_headers = [
            "Kaynak Dosya",
            "Sayfa",
            "Kaynak Satır",
            "Tarih",
            "Belge No",
            "Taraf",
            "Vergi No",
            "Açıklama",
            "Net Tutar",
            "KDV Tutar",
            "Brüt Tutar",
            "Para Birimi",
            "Durum",
        ]
        for c_i, h in enumerate(clean_headers):
            ws_clean.write_string(0, c_i, h, hdr_fmt)
        ws_clean.freeze_panes(1, 0)
        if report.records:
            ws_clean.autofilter(0, 0, len(report.records), len(clean_headers) - 1)

        col_widths_clean = [24, 14, 13, 13, 18, 28, 16, 30, 15, 14, 15, 11, 12]
        for idx, width in enumerate(col_widths_clean):
            ws_clean.set_column(idx, idx, width)

        status_fmt_map = {"Geçti": ok_fmt, "Uyarı": warn_fmt, "Hata": err_fmt}
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

        # 3. Sheet: Kaynak_Dosyalar
        ws_src = wb.add_worksheet("Kaynak_Dosyalar")
        src_headers = [
            "Dosya",
            "Sayfa",
            "Biçim",
            "Karakter Seti",
            "Ayırıcı",
            "Başlık Satırı",
            "Ham Satır",
            "Alınan Satır",
            "Atlanan Mükerrer",
            "Uyarı",
            "Hata",
            "Ham Brüt Toplam",
            "Alınan Brüt Toplam",
        ]
        for c_i, h in enumerate(src_headers):
            ws_src.write_string(0, c_i, h, hdr_fmt)
        ws_src.freeze_panes(1, 0)
        for idx, width in enumerate([26, 14, 10, 14, 10, 13, 12, 13, 17, 10, 10, 18, 18]):
            ws_src.set_column(idx, idx, width)

        for r_idx, st in enumerate(report.source_stats, start=1):
            ws_src.write_string(r_idx, 0, sanitize_excel_text(st.source_file), text_fmt)
            ws_src.write_string(r_idx, 1, sanitize_excel_text(st.sheet_name), text_fmt)
            ws_src.write_string(r_idx, 2, st.file_format.upper(), text_fmt)
            ws_src.write_string(r_idx, 3, st.encoding, text_fmt)
            ws_src.write_string(r_idx, 4, sanitize_excel_text(st.delimiter or "-"), text_fmt)
            ws_src.write_number(r_idx, 5, st.header_row, num_int_fmt)
            ws_src.write_number(r_idx, 6, st.raw_data_rows, num_int_fmt)
            ws_src.write_number(r_idx, 7, st.kept_rows, num_int_fmt)
            ws_src.write_number(r_idx, 8, st.duplicate_skipped, num_int_fmt)
            ws_src.write_number(r_idx, 9, st.warning_count, num_int_fmt)
            ws_src.write_number(r_idx, 10, st.error_count, num_int_fmt)
            ws_src.write_number(r_idx, 11, float(st.raw_brut_sum), money_fmt)
            ws_src.write_number(r_idx, 12, float(st.kept_brut_sum), money_fmt)

        # 4. Sheet: Mukerrer_ve_Uyarilar
        ws_iss = wb.add_worksheet("Mukerrer_ve_Uyarilar")
        iss_headers = ["Önem", "Kod", "Kaynak Dosya", "Sayfa", "Satır", "Belge No", "Açıklama"]
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
