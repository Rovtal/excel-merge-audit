"""Tests for dirty Turkish/global money/date parsing, dedupe reconciliation, and conflict traps."""

from decimal import Decimal
from pathlib import Path
import tempfile
import unittest

import xlsxwriter

from excel_merge_audit.merger import merge_and_audit
from excel_merge_audit.normalizer import parse_date_value, parse_decimal_amount


class NormalizerAndMergerTests(unittest.TestCase):
    """Verify exact Decimal parsing, header detection, deduplication math, and conflicts."""

    def test_tr_and_en_money_parsing_and_garbage_rejection(self) -> None:
        self.assertEqual(parse_decimal_amount("1.250,50 TL"), (Decimal("1250.50"), "TRY"))
        self.assertEqual(parse_decimal_amount("(412,50)"), (Decimal("-412.50"), None))
        self.assertEqual(parse_decimal_amount("1,250.50 USD"), (Decimal("1250.50"), "USD"))
        self.assertEqual(parse_decimal_amount("-80.25"), (Decimal("-80.25"), None))
        self.assertEqual(parse_decimal_amount("0,13"), (Decimal("0.13"), None))

        bad_val, bad_msg = parse_decimal_amount("#DEĞER!")
        self.assertIsNone(bad_val)
        self.assertIn("formül hatası", bad_msg or "")

        garbage_val, _ = parse_decimal_amount("ödenmedi")
        self.assertIsNone(garbage_val)

    def test_date_formats_and_invalid_dates(self) -> None:
        self.assertEqual(parse_date_value("16.11.2026"), ("2026-11-16", None))
        self.assertEqual(parse_date_value("16/11/2026 14:30"), ("2026-11-16", None))
        self.assertEqual(parse_date_value("2026-11-16"), ("2026-11-16", None))
        _, err = parse_date_value("32.13.2026")
        self.assertIsNotNone(err)

    def test_merge_reconciliation_duplicates_conflicts_and_subtotal_skip(self) -> None:
        with tempfile.TemporaryDirectory() as tmp_str:
            tmp = Path(tmp_str)
            f1 = tmp / "ocak.csv"
            f1.write_bytes(
                "\r\n".join(
                    [
                        "RAPOR BASLIK;;;;",
                        "Tarih;Fatura No;Cari Adı;VKN;Açıklama;Matrah;KDV;Genel Toplam",
                        "10.01.2026;FT-100;Örnek A.Ş.;1111111111;Malzeme;1.000,00;200,00;1.200,00",
                        "10.01.2026;FT-100;Örnek A.Ş.;1111111111;Malzeme;1.000,00;200,00;1.200,00",
                        ";GENEL TOPLAM;;;;2.000,00;400,00;2.400,00",
                    ]
                ).encode("cp1254")
            )

            f2 = tmp / "subat.xlsx"
            wb = xlsxwriter.Workbook(str(f2))
            ws = wb.add_worksheet("Sheet1")
            ws.write_row(0, 0, ["Date", "Invoice No", "Customer", "Tax No", "Subtotal", "VAT", "Gross Total"])
            ws.write_row(1, 0, ["2026-01-10", "FT-100", "Örnek A.Ş.", "9,05321E+11", 1000.00, 200.00, 1400.00])
            ws.write_row(2, 0, ["2026-02-01", "FT-200", "Temiz Ltd.", "2222222222", 500.00, 100.00, 600.00])
            wb.close()

            report = merge_and_audit([f1, f2])
            s = report.summary
            self.assertEqual(s.total_input_rows, 4)
            self.assertEqual(s.exact_duplicates_removed, 1)
            self.assertEqual(s.kept_rows, 3)
            self.assertEqual(s.raw_brut_total, Decimal("4400.00"))
            self.assertEqual(s.duplicate_brut_removed, Decimal("1200.00"))
            self.assertEqual(s.kept_brut_total, Decimal("3200.00"))
            self.assertEqual(s.raw_brut_total - s.duplicate_brut_removed, s.kept_brut_total)
            self.assertEqual(s.conflicting_documents, 1)

            codes = {i.code for i in report.issues}
            self.assertIn("TAM_MUKERRER_SILINDI", codes)
            self.assertIn("CAKISAN_BELGE", codes)
            self.assertIn("BILIMSEL_GOSTERIM_VKN", codes)
