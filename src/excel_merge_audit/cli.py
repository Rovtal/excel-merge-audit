"""Command-line interface for multi-Excel/CSV merge and audit."""

import argparse
from dataclasses import asdict
from decimal import Decimal
import json
from pathlib import Path
import sys
from typing import Optional

from .exporter import export_merge_workbook
from .merger import merge_and_audit
from .models import MergeReport
from .samples import create_dirty_demo_inputs


def _json_default(obj: object) -> object:
    if isinstance(obj, Decimal):
        return str(obj)
    raise TypeError(f"Serileştirilemeyen tip: {type(obj).__name__}")


def _write_json(report: MergeReport, path: Path) -> None:
    resolved = path.resolve()
    resolved.parent.mkdir(parents=True, exist_ok=True)
    payload = asdict(report)
    resolved.write_text(
        json.dumps(payload, ensure_ascii=False, indent=2, default=_json_default),
        encoding="utf-8",
    )


def _print_console_summary(report: MergeReport, xlsx_out: Path) -> None:
    s = report.summary
    print("=== Excel / CSV Birleştirme ve Denetim Özeti ===")
    print(f"Kaynak Tablo        : {s.source_sheets}")
    print(f"Ham Girdi Satırı    : {s.total_input_rows}")
    print(f"Temiz Alınan Satır  : {s.kept_rows}")
    print(f"Silinen Mükerrer    : {s.exact_duplicates_removed}")
    print(f"Çakışan Belge Grubu : {s.conflicting_documents}")
    print(f"Geçti / Uyarı / Hata: {s.passed_rows} / {s.warning_rows} / {s.error_rows}")
    print(f"Ham Brüt Toplam     : {s.raw_brut_total:.2f}")
    print(f"Silinen Mükerrer    : {s.duplicate_brut_removed:.2f}")
    print(f"Temiz Brüt Toplam   : {s.kept_brut_total:.2f}")
    print(f"Excel Rapor Çıktısı : {xlsx_out}")


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="excel-merge",
        description="Çevrimdışı çoklu .xlsx/.csv/.tsv birleştirici ve denetim raporlayıcı.",
    )
    sub = parser.add_subparsers(dest="command", required=True)

    merge_p = sub.add_parser("merge", help="Verilen dosya/klasörleri birleştir ve denetle.")
    merge_p.add_argument("inputs", nargs="+", help="Birleştirilecek .xlsx/.csv/.tsv dosya veya klasörleri.")
    merge_p.add_argument("-o", "--output", required=True, help="Çıktı .xlsx dosya yolu.")
    merge_p.add_argument("--json", help="İsteğe bağlı JSON özet çıktı yolu.")
    merge_p.add_argument("--sheet", help=".xlsx dosyalarında yalnızca belirtilen sayfa adını oku.")
    merge_p.add_argument("--keep-duplicates", action="store_true", help="Birebir mükerrerleri çıkarma.")
    merge_p.add_argument(
        "--fail-on-error",
        action="store_true",
        help="Kritik hata (durum=Hata) varsa çıkış kodunu 2 yap.",
    )

    demo_p = sub.add_parser("demo", help="Kirli örnek girdi dosyalarıyla tam demo çalıştır.")
    demo_p.add_argument("--out-dir", default="demo_output", help="Demo çıktı klasörü.")
    return parser


def main(argv: Optional[list[str]] = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)

    try:
        if args.command == "demo":
            out_dir = Path(args.out_dir).resolve()
            samples_dir = out_dir / "ornek_girdiler"
            sample_files = create_dirty_demo_inputs(samples_dir)
            report = merge_and_audit(sample_files)
            xlsx_out = out_dir / "birlestirilmis_denetim_raporu.xlsx"
            export_merge_workbook(report, xlsx_out)
            _write_json(report, out_dir / "ozet.json")
            _print_console_summary(report, xlsx_out)
            return 0

        input_paths = [Path(p) for p in args.inputs]
        report = merge_and_audit(
            input_paths,
            only_sheet=args.sheet,
            keep_duplicates=args.keep_duplicates,
        )
        if report.fatal_reason:
            print(f"HATA: {report.fatal_reason}", file=sys.stderr)
            return 1

        xlsx_out = export_merge_workbook(report, Path(args.output))
        if args.json:
            _write_json(report, Path(args.json))
        _print_console_summary(report, xlsx_out)

        if args.fail_on_error and report.summary.error_rows > 0:
            return 2
        return 0
    except (FileNotFoundError, ValueError) as exc:
        print(f"HATA: {exc}", file=sys.stderr)
        return 1
    except Exception as exc:
        print(f"BEKLENMEYEN HATA: {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
