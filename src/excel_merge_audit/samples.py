"""Deterministic dirty multi-format sample dataset generator for demos and tests."""

from pathlib import Path
import xlsxwriter


def create_dirty_demo_inputs(target_dir: Path) -> list[Path]:
    """Create 3 realistic messy CSV/XLSX/TSV files containing TR/EN formats, duplicates, and traps."""
    resolved = target_dir.resolve()
    resolved.mkdir(parents=True, exist_ok=True)

    # 1. cp1254 Turkish ANSI CSV with semicolon delimiter, leading title row, exact duplicate, formula payload
    csv_path = resolved / "satis_ocak_bozuk_cp1254.csv"
    csv_lines = [
        "OCAK 2026 SATIS RAPORU;;;;;;;",
        "Fatura Tarihi;Belge No;Cari Unvan;VKN / TCKN;Açıklama;Matrah;KDV;Genel Toplam",
        "05.01.2026;FT-2026-001;Anadolu Lojistik A.Ş.;1234567890;Yük Taşıma Bedeli;1.000,00 TL;200,00 TL;1.200,00 TL",
        "12.01.2026;FT-2026-002;Güneş Tedarik Ltd. Şti.;9876543210;Paketleme Malzemesi;2.500,00;500,00;3.000,00",
        "12.01.2026;FT-2026-002;Güneş Tedarik Ltd. Şti.;9876543210;Paketleme Malzemesi;2.500,00;500,00;3.000,00",
        "19.01.2026;FT-2026-003;Kuzey Büro Sistemleri;5554443322;=CMD|' /C calc'!A0 Formül Testi;450,00;90,00;540,00",
    ]
    csv_path.write_bytes("\r\n".join(csv_lines).encode("cp1254"))

    # 2. XLSX with 2 leading banner rows, blank column gap, conflicting invoice FT-2026-001, Net+KDV mismatch
    xlsx_path = resolved / "satis_subat_bayi.xlsx"
    wb = xlsxwriter.Workbook(str(xlsx_path))
    ws = wb.add_worksheet("Subat_Veri")
    date_fmt = wb.add_format({"num_format": "dd.mm.yyyy"})

    ws.write(0, 0, "BAYİ SATIŞ AKTARIMI - ŞUBAT 2026")
    ws.write(1, 0, "Oluşturma: Otomatik İhracat")
    headers = ["İşlem Tarihi", "Fatura Numarası", "Müşteri Adı", "Not", "Ara Toplam", "KDV Tutarı", "Toplam Tutar"]
    for c_i, hdr in enumerate(headers):
        col = c_i + 1 if c_i >= 3 else c_i
        ws.write(3, col, hdr)

    ws.write(4, 0, "05.01.2026")
    ws.write(4, 1, "FT-2026-001")
    ws.write(4, 2, "Anadolu Lojistik A.Ş.")
    ws.write(4, 4, "Çakışan tutarlı tekrar kayıt")
    ws.write(4, 5, 1000.00)
    ws.write(4, 6, 200.00)
    ws.write(4, 7, 1350.00)

    ws.write_number(5, 0, 46060, date_fmt)  # 2026-02-07
    ws.write(5, 1, "FT-2026-011")
    ws.write(5, 2, "Ege Teknik Ticaret")
    ws.write(5, 4, "Bakım Sözleşmesi")
    ws.write(5, 5, 3200.00)
    ws.write(5, 6, 640.00)
    ws.write(5, 7, 3840.00)

    ws.write(6, 0, "14.02.2026")
    ws.write(6, 1, "FT-2026-012")
    ws.write(6, 2, "Marmara Gıda A.Ş.")
    ws.write(6, 4, "Hatalı KDV toplamlı satır")
    ws.write(6, 5, 500.00)
    ws.write(6, 6, 100.00)
    ws.write(6, 7, 650.00)
    wb.close()

    # 3. UTF-8 BOM TSV file with scientific notation VKN warning and negative refund
    tsv_path = resolved / "iade_mart.tsv"
    tsv_lines = [
        "Issue Date\tDocument No\tCounterparty\tTax ID\tDescription\tSubtotal\tVAT\tGross Total\tCurrency",
        "2026-03-03\tFT-2026-021\tDelta Ofis Ltd.\t9,05321E+11\tBilimsel gösterimli vergi no\t800.00\t160.00\t960.00\tTRY",
        "2026-03-09\tIADE-2026-01\tEge Teknik Ticaret\t1112223334\tKısmi İade Faturası\t-200.00\t-40.00\t-240.00\tTRY",
    ]
    tsv_path.write_bytes(("\r\n".join(tsv_lines)).encode("utf-8-sig"))

    return [csv_path, xlsx_path, tsv_path]
