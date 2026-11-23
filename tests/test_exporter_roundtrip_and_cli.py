"""End-to-end round-trip .xlsx XML verification, CWE-1236 safety, and CLI exit code tests."""

from decimal import Decimal
import json
from pathlib import Path
import tempfile
import unittest
import zipfile

from lxml import etree

from excel_merge_audit.cli import main
from excel_merge_audit.readers import NS_SS, SAFE_XML_PARSER


class ExporterRoundTripAndCliTests(unittest.TestCase):
    """Verify exported workbook formulas, CWE-1236 neutralization, exact sum reconciliation, and CLI exits."""

    def test_demo_round_trip_xml_sums_and_cwe1236_neutralization(self) -> None:
        with tempfile.TemporaryDirectory() as tmp_str:
            out_dir = Path(tmp_str) / "demo_out"
            rc = main(["demo", "--out-dir", str(out_dir)])
            self.assertEqual(rc, 0)

            xlsx_path = out_dir / "birlestirilmis_denetim_raporu.xlsx"
            json_path = out_dir / "ozet.json"
            self.assertTrue(xlsx_path.is_file())
            self.assertTrue(json_path.is_file())

            payload = json.loads(json_path.read_text(encoding="utf-8"))
            summary = payload["summary"]
            self.assertEqual(summary["total_input_rows"], 9)
            self.assertEqual(summary["exact_duplicates_removed"], 1)
            self.assertEqual(summary["kept_rows"], 8)
            self.assertEqual(summary["conflicting_documents"], 1)
            self.assertEqual(
                Decimal(summary["raw_brut_total"]) - Decimal(summary["duplicate_brut_removed"]),
                Decimal(summary["kept_brut_total"]),
            )

            with zipfile.ZipFile(xlsx_path, "r") as zf:
                wb_xml = etree.fromstring(zf.read("xl/workbook.xml"), parser=SAFE_XML_PARSER)
                sheet_names = [el.get("name") for el in wb_xml.findall(f".//{{{NS_SS}}}sheet")]
                self.assertEqual(
                    sheet_names,
                    ["Ozet_Birlestirme", "Temiz_Birlestirilmis", "Kaynak_Dosyalar", "Mukerrer_ve_Uyarilar"],
                )

                shared_xml = etree.fromstring(zf.read("xl/sharedStrings.xml"), parser=SAFE_XML_PARSER)
                shared_strings = ["".join(si.itertext()) for si in shared_xml.findall(f".//{{{NS_SS}}}si")]

                # CWE-1236: dangerous formula string must stay escaped with leading quote and never become a <f> tag
                escaped_payloads = [s for s in shared_strings if "CMD|" in s]
                self.assertEqual(len(escaped_payloads), 1)
                self.assertTrue(escaped_payloads[0].startswith("'="))

                # Sheet 2 = Temiz_Birlestirilmis: sum actual XML columns I, J, K and verify zero <f> tags
                clean_xml = etree.fromstring(zf.read("xl/worksheets/sheet2.xml"), parser=SAFE_XML_PARSER)
                self.assertEqual(len(clean_xml.findall(f".//{{{NS_SS}}}f")), 0)

                net_sum = Decimal("0.00")
                kdv_sum = Decimal("0.00")
                brut_sum = Decimal("0.00")
                data_rows = clean_xml.findall(f".//{{{NS_SS}}}sheetData/{{{NS_SS}}}row")[1:]
                self.assertEqual(len(data_rows), summary["kept_rows"])
                for row in data_rows:
                    for c in row.findall(f"{{{NS_SS}}}c"):
                        ref = c.get("r", "")
                        v = c.find(f"{{{NS_SS}}}v")
                        if v is None or v.text is None:
                            continue
                        if ref.startswith("I"):
                            net_sum += Decimal(v.text).quantize(Decimal("0.01"))
                        elif ref.startswith("J"):
                            kdv_sum += Decimal(v.text).quantize(Decimal("0.01"))
                        elif ref.startswith("K"):
                            brut_sum += Decimal(v.text).quantize(Decimal("0.01"))

                self.assertEqual(net_sum, Decimal(summary["kept_net_total"]))
                self.assertEqual(kdv_sum, Decimal(summary["kept_kdv_total"]))
                self.assertEqual(brut_sum, Decimal(summary["kept_brut_total"]))

                # Sheet 1 = Ozet_Birlestirme: verify dynamic COUNTIF, SUM, and Mutabakat Sapmasi formula coords
                sum_xml = etree.fromstring(zf.read("xl/worksheets/sheet1.xml"), parser=SAFE_XML_PARSER)
                formulas_by_ref = {}
                for c in sum_xml.findall(f".//{{{NS_SS}}}c"):
                    f_el = c.find(f"{{{NS_SS}}}f")
                    if f_el is not None and f_el.text:
                        formulas_by_ref[c.get("r")] = f_el.text

                last_row = summary["kept_rows"] + 1
                self.assertIn("COUNTIF(Temiz_Birlestirilmis!M2:M9", formulas_by_ref.get("B9", ""))
                self.assertEqual(formulas_by_ref.get("B14"), f"SUM(Temiz_Birlestirilmis!I2:I{last_row})")
                self.assertEqual(formulas_by_ref.get("B15"), f"SUM(Temiz_Birlestirilmis!J2:J{last_row})")
                self.assertEqual(formulas_by_ref.get("B16"), f"SUM(Temiz_Birlestirilmis!K2:K{last_row})")
                self.assertEqual(formulas_by_ref.get("B17"), "B12-B13-B16")

    def test_cli_fail_on_error_and_missing_input_exit_codes(self) -> None:
        with tempfile.TemporaryDirectory() as tmp_str:
            tmp = Path(tmp_str)
            demo_dir = tmp / "demo_src"
            self.assertEqual(main(["demo", "--out-dir", str(demo_dir)]), 0)

            out_xlsx = tmp / "fail_check.xlsx"
            rc_err = main(
                [
                    "merge",
                    str(demo_dir / "ornek_girdiler"),
                    "-o",
                    str(out_xlsx),
                    "--fail-on-error",
                ]
            )
            self.assertEqual(rc_err, 2)
            self.assertTrue(out_xlsx.is_file())

            missing_rc = main(["merge", str(tmp / "yok.csv"), "-o", str(out_xlsx)])
            self.assertEqual(missing_rc, 1)

    def test_english_demo_and_merge_flag_order_and_formulas(self) -> None:
        with tempfile.TemporaryDirectory() as tmp_str:
            tmp = Path(tmp_str)
            demo_dir = tmp / "demo_en"
            rc = main(["demo", "--lang", "en", "--out-dir", str(demo_dir)])
            self.assertEqual(rc, 0)

            xlsx_path = demo_dir / "merged_audit_report.xlsx"
            json_path = demo_dir / "summary.json"
            self.assertTrue(xlsx_path.is_file())
            self.assertTrue(json_path.is_file())

            payload = json.loads(json_path.read_text(encoding="utf-8"))
            self.assertEqual(payload["lang"], "en")
            summary = payload["summary"]
            self.assertEqual(summary["total_input_rows"], 9)
            self.assertEqual(summary["exact_duplicates_removed"], 1)
            self.assertEqual(summary["kept_rows"], 8)
            self.assertEqual(summary["conflicting_documents"], 1)
            self.assertEqual(
                Decimal(summary["raw_brut_total"]) - Decimal(summary["duplicate_brut_removed"]),
                Decimal(summary["kept_brut_total"]),
            )
            statuses = {rec["durum"] for rec in payload["records"]}
            self.assertEqual(statuses, {"Passed", "Warning", "Error"})

            with zipfile.ZipFile(xlsx_path, "r") as zf:
                wb_xml = etree.fromstring(zf.read("xl/workbook.xml"), parser=SAFE_XML_PARSER)
                sheet_names = [el.get("name") for el in wb_xml.findall(f".//{{{NS_SS}}}sheet")]
                self.assertEqual(
                    sheet_names,
                    ["Merge_Summary", "Clean_Merged", "Source_Files", "Duplicates_and_Warnings"],
                )

                sum_xml = etree.fromstring(zf.read("xl/worksheets/sheet1.xml"), parser=SAFE_XML_PARSER)
                formulas_by_ref = {}
                for c in sum_xml.findall(f".//{{{NS_SS}}}c"):
                    f_el = c.find(f"{{{NS_SS}}}f")
                    if f_el is not None and f_el.text:
                        formulas_by_ref[c.get("r")] = f_el.text

                last_row = summary["kept_rows"] + 1
                self.assertIn('COUNTIF(Clean_Merged!M2:M9,"Passed")', formulas_by_ref.get("B9", ""))
                self.assertEqual(formulas_by_ref.get("B14"), f"SUM(Clean_Merged!I2:I{last_row})")
                self.assertEqual(formulas_by_ref.get("B15"), f"SUM(Clean_Merged!J2:J{last_row})")
                self.assertEqual(formulas_by_ref.get("B16"), f"SUM(Clean_Merged!K2:K{last_row})")
                self.assertEqual(formulas_by_ref.get("B17"), "B12-B13-B16")

            # Verify --lang before subcommand also works cleanly and --fail-on-error exits with 2
            merge_out = tmp / "merged_en.xlsx"
            rc_err = main(
                [
                    "--lang",
                    "en",
                    "merge",
                    str(demo_dir / "sample_inputs"),
                    "-o",
                    str(merge_out),
                    "--fail-on-error",
                ]
            )
            self.assertEqual(rc_err, 2)
            self.assertTrue(merge_out.is_file())

            with self.assertRaises(SystemExit) as bad_lang_ctx:
                main(["demo", "--lang", "fr"])
            self.assertEqual(bad_lang_ctx.exception.code, 2)
