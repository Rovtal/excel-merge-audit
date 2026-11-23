"""Deterministic dirty multi-format sample dataset generator for demos and tests."""

from pathlib import Path
import xlsxwriter

from .i18n import validate_lang


_DEMO_TEXT: dict[str, dict[str, object]] = {
    "tr": {
        "csv_name": "satis_ocak_bozuk_cp1254.csv",
        "csv_lines": [
            "OCAK 2026 SATIS RAPORU;;;;;;;",
            "Fatura Tarihi;Belge No;Cari Unvan;VKN / TCKN;Açıklama;Matrah;KDV;Genel Toplam",
            "05.01.2026;FT-2026-001;Anadolu Lojistik A.Ş.;1234567890;Yük Taşıma Bedeli;1.000,00 TL;200,00 TL;1.200,00 TL",
            "12.01.2026;FT-2026-002;Güneş Tedarik Ltd. Şti.;9876543210;Paketleme Malzemesi;2.500,00;500,00;3.000,00",
            "12.01.2026;FT-2026-002;Güneş Tedarik Ltd. Şti.;9876543210;Paketleme Malzemesi;2.500,00;500,00;3.000,00",
            "19.01.2026;FT-2026-003;Kuzey Büro Sistemleri;5554443322;=CMD|' /C calc'!A0 Formül Testi;450,00;90,00;540,00",
        ],
        "xlsx_name": "satis_subat_bayi.xlsx",
        "sheet_name": "Subat_Veri",
        "banner_1": "BAYİ SATIŞ AKTARIMI - ŞUBAT 2026",
        "banner_2": "Oluşturma: Otomatik İhracat",
        "headers": ["İşlem Tarihi", "Fatura Numarası", "Müşteri Adı", "Not", "Ara Toplam", "KDV Tutarı", "Toplam Tutar"],
        "xlsx_rows": [
            ("05.01.2026", "FT-2026-001", "Anadolu Lojistik A.Ş.", "Çakışan tutarlı tekrar kayıt", 1000.00, 200.00, 1350.00),
            (46060, "FT-2026-011", "Ege Teknik Ticaret", "Bakım Sözleşmesi", 3200.00, 640.00, 3840.00),
            ("14.02.2026", "FT-2026-012", "Marmara Gıda A.Ş.", "Hatalı KDV toplamlı satır", 500.00, 100.00, 650.00),
        ],
        "tsv_name": "iade_mart.tsv",
        "tsv_lines": [
            "Issue Date\tDocument No\tCounterparty\tTax ID\tDescription\tSubtotal\tVAT\tGross Total\tCurrency",
            "2026-03-03\tFT-2026-021\tDelta Ofis Ltd.\t9,05321E+11\tBilimsel gösterimli vergi no\t800.00\t160.00\t960.00\tTRY",
            "2026-03-09\tIADE-2026-01\tEge Teknik Ticaret\t1112223334\tKısmi İade Faturası\t-200.00\t-40.00\t-240.00\tTRY",
        ],
    },
    "en": {
        "csv_name": "sales_jan_cp1254.csv",
        "csv_lines": [
            "JAN 2026 SALES REPORT;;;;;;;",
            "Invoice Date;Document No;Counterparty;Tax ID;Description;Subtotal;VAT;Gross Total",
            "05.01.2026;INV-2026-001;Anatolia Logistics Inc.;1234567890;Freight Charge;1.000,00 TL;200,00 TL;1.200,00 TL",
            "12.01.2026;INV-2026-002;Sun Supply Ltd.;9876543210;Packaging Materials;2.500,00;500,00;3.000,00",
            "12.01.2026;INV-2026-002;Sun Supply Ltd.;9876543210;Packaging Materials;2.500,00;500,00;3.000,00",
            "19.01.2026;INV-2026-003;North Office Systems;5554443322;=CMD|' /C calc'!A0 Formula Trap;450,00;90,00;540,00",
        ],
        "xlsx_name": "sales_feb_branch.xlsx",
        "sheet_name": "Feb_Data",
        "banner_1": "BRANCH SALES EXPORT - FEB 2026",
        "banner_2": "Generated: Automated Export",
        "headers": ["Date", "Invoice No", "Customer", "Notes", "Net Amount", "VAT", "Total"],
        "xlsx_rows": [
            ("05.01.2026", "INV-2026-001", "Anatolia Logistics Inc.", "Conflicting amount duplicate", 1000.00, 200.00, 1350.00),
            (46060, "INV-2026-011", "Aegean Tech Trade", "Maintenance Contract", 3200.00, 640.00, 3840.00),
            ("14.02.2026", "INV-2026-012", "Marmara Foods Inc.", "Mismatched VAT total row", 500.00, 100.00, 650.00),
        ],
        "tsv_name": "returns_mar.tsv",
        "tsv_lines": [
            "Issue Date\tDocument No\tCounterparty\tTax ID\tDescription\tSubtotal\tVAT\tGross Total\tCurrency",
            "2026-03-03\tINV-2026-021\tDelta Office Ltd.\t9,05321E+11\tScientific notation tax ID\t800.00\t160.00\t960.00\tTRY",
            "2026-03-09\tCRN-2026-01\tAegean Tech Trade\t1112223334\tPartial Return Credit Note\t-200.00\t-40.00\t-240.00\tTRY",
        ],
    },
}


def create_dirty_demo_inputs(target_dir: Path, lang: str = "tr") -> list[Path]:
    """Create 3 realistic messy CSV/XLSX/TSV files containing TR/EN formats, duplicates, and traps."""
    active_lang = validate_lang(lang)
    spec = _DEMO_TEXT[active_lang]
    resolved = target_dir.resolve()
    resolved.mkdir(parents=True, exist_ok=True)

    # 1. cp1254 ANSI CSV with semicolon delimiter, leading title row, exact duplicate, formula payload
    csv_path = resolved / str(spec["csv_name"])
    csv_lines: list[str] = list(spec["csv_lines"])  # type: ignore[arg-type]
    csv_path.write_bytes("\r\n".join(csv_lines).encode("cp1254"))

    # 2. XLSX with 2 leading banner rows, blank column gap, conflicting invoice, Net+VAT mismatch
    xlsx_path = resolved / str(spec["xlsx_name"])
    wb = xlsxwriter.Workbook(str(xlsx_path))
    ws = wb.add_worksheet(str(spec["sheet_name"]))
    date_fmt = wb.add_format({"num_format": "dd.mm.yyyy"})

    ws.write(0, 0, str(spec["banner_1"]))
    ws.write(1, 0, str(spec["banner_2"]))
    headers: list[str] = list(spec["headers"])  # type: ignore[arg-type]
    for c_i, hdr in enumerate(headers):
        col = c_i + 1 if c_i >= 3 else c_i
        ws.write(3, col, hdr)

    xlsx_rows: list[tuple[object, ...]] = list(spec["xlsx_rows"])  # type: ignore[arg-type]
    for idx, (date_val, doc_no, party, note, net_val, vat_val, gross_val) in enumerate(xlsx_rows, start=4):
        if isinstance(date_val, int):
            ws.write_number(idx, 0, date_val, date_fmt)
        else:
            ws.write(idx, 0, str(date_val))
        ws.write(idx, 1, str(doc_no))
        ws.write(idx, 2, str(party))
        ws.write(idx, 4, str(note))
        ws.write(idx, 5, float(net_val))  # type: ignore[arg-type]
        ws.write(idx, 6, float(vat_val))  # type: ignore[arg-type]
        ws.write(idx, 7, float(gross_val))  # type: ignore[arg-type]
    wb.close()

    # 3. UTF-8 BOM TSV file with scientific notation tax ID warning and negative refund
    tsv_path = resolved / str(spec["tsv_name"])
    tsv_lines: list[str] = list(spec["tsv_lines"])  # type: ignore[arg-type]
    tsv_path.write_bytes("\r\n".join(tsv_lines).encode("utf-8-sig"))

    return [csv_path, xlsx_path, tsv_path]
