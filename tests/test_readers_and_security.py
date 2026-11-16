"""Adversarial tests for offline XLSX/CSV readers, encodings, date serials, and archive security."""

from pathlib import Path
import tempfile
import unittest
import zipfile

import xlsxwriter

from excel_merge_audit.readers import OLE2_MAGIC, load_input_tables, read_delimited_table, read_xlsx_tables


class ReaderAndSecurityTests(unittest.TestCase):
    """Verify encoding fallback, delimiter detection, sparse OpenXML cells, and security guards."""

    def test_cp1254_csv_quoted_delimiters_and_tsv_detection(self) -> None:
        with tempfile.TemporaryDirectory() as tmp_str:
            tmp = Path(tmp_str)
            csv_file = tmp / "tr_veri.csv"
            lines = [
                "Tarih;Belge No;Cari Unvan;Açıklama;Genel Toplam",
                '01.02.2026;F-10;"Yıldız Şirketi; Merkez";İçerik ğüşiöçİĞÜŞÖÇ;"1.250,50 TL"',
            ]
            csv_file.write_bytes("\r\n".join(lines).encode("cp1254"))

            table = read_delimited_table(csv_file)
            self.assertEqual(table.encoding, "cp1254")
            self.assertEqual(table.delimiter, ";")
            self.assertEqual(table.rows[1][2], "Yıldız Şirketi; Merkez")
            self.assertIn("ğüşiöçİĞÜŞÖÇ", table.rows[1][3])
            self.assertEqual(table.rows[1][4], "1.250,50 TL")

    def test_xlsx_sparse_cells_rich_shared_strings_dates_and_formula_errors(self) -> None:
        with tempfile.TemporaryDirectory() as tmp_str:
            xlsx_path = Path(tmp_str) / "sparse_and_dates.xlsx"
            wb = xlsxwriter.Workbook(str(xlsx_path))
            ws = wb.add_worksheet("Veri")
            date_fmt = wb.add_format({"num_format": "dd.mm.yyyy"})
            bold_fmt = wb.add_format({"bold": True})

            ws.write(2, 0, "Tarih")
            ws.write(2, 1, "Belge No")
            ws.write(2, 4, "Taraf")
            ws.write(2, 5, "Tutar")

            ws.write_number(3, 0, 46342, date_fmt)  # 2026-11-16
            ws.write(3, 1, "FT-99")
            ws.write_rich_string(3, 4, bold_fmt, "Güney ", "Makina")
            ws.write_number(3, 5, 1999.95)
            wb.close()

            # Inject a formula error cell (#DIV/0!) cleanly into sheet1 XML to test detection
            patched = Path(tmp_str) / "with_err.xlsx"
            with zipfile.ZipFile(xlsx_path, "r") as zin, zipfile.ZipFile(patched, "w") as zout:
                for item in zin.infolist():
                    data = zin.read(item.filename)
                    if item.filename == "xl/worksheets/sheet1.xml":
                        text = data.decode("utf-8").replace(
                            "</sheetData>",
                            '<row r="5"><c r="A5" t="s"><v>0</v></c><c r="F5" t="e"><f>1/0</f><v>#DIV/0!</v></c></row></sheetData>',
                        )
                        data = text.encode("utf-8")
                    zout.writestr(item, data)

            tables = read_xlsx_tables(patched)
            self.assertEqual(len(tables), 1)
            t = tables[0]
            self.assertEqual(t.rows[3][0], "2026-11-16")
            self.assertEqual(t.rows[3][2], "")
            self.assertEqual(t.rows[3][4], "Güney Makina")
            self.assertEqual(t.rows[3][5], "1999.95")
            self.assertEqual(t.formula_error_cells, [(5, 6, "#DIV/0!")])

    def test_security_rejects_zip_slip_xxe_ole2_xls_and_temp_lockfiles(self) -> None:
        with tempfile.TemporaryDirectory() as tmp_str:
            tmp = Path(tmp_str)
            bad_dir = tmp / "bad"
            clean_dir = tmp / "clean"
            bad_dir.mkdir()
            clean_dir.mkdir()

            bad_zip = bad_dir / "slip.xlsx"
            with zipfile.ZipFile(bad_zip, "w") as zf:
                zf.writestr("../escape.txt", "evil")
            with self.assertRaises(ValueError):
                read_xlsx_tables(bad_zip)

            ole_file = bad_dir / "legacy.xls"
            ole_file.write_bytes(OLE2_MAGIC + b"\x00" * 64)
            with self.assertRaises(ValueError):
                load_input_tables([ole_file])

            # Lock file ~$*.xlsx and unsupported .txt must be ignored safely without failing CSV load
            (clean_dir / "~$temp.xlsx").write_bytes(b"locked")
            (clean_dir / "not.txt").write_text("ignore", encoding="utf-8")
            good_csv = clean_dir / "ok.csv"
            good_csv.write_text("Tarih,Belge No,Tutar\n2026-11-16,A1,100.00\n", encoding="utf-8")

            tables, ignored = load_input_tables([clean_dir])
            self.assertEqual(len(tables), 1)
            self.assertEqual(len(ignored), 2)

            # XXE entity inside workbook.xml must never resolve external filesystem content
            xxe_xlsx = bad_dir / "xxe.xlsx"
            with zipfile.ZipFile(xxe_xlsx, "w") as zf:
                zf.writestr(
                    "xl/workbook.xml",
                    '<?xml version="1.0"?><!DOCTYPE foo [<!ENTITY xxe SYSTEM "file:///etc/passwd">]><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheets><sheet name="&xxe;" sheetId="1" r:id="rId1" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/></sheets></workbook>',
                )
                zf.writestr(
                    "xl/_rels/workbook.xml.rels",
                    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Target="worksheets/sheet1.xml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"/></Relationships>',
                )
                zf.writestr(
                    "xl/worksheets/sheet1.xml",
                    '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/></worksheet>',
                )
            with self.assertRaises(ValueError):
                read_xlsx_tables(xxe_xlsx)
