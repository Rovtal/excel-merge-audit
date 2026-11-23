"""Command-line interface for multi-Excel/CSV merge and audit."""

import argparse
from dataclasses import asdict
from decimal import Decimal
import json
from pathlib import Path
import sys
from typing import Optional

from . import __version__
from .exporter import export_clean_csv, export_merge_workbook
from .i18n import DEMO_LAYOUT, SUPPORTED_LANGS, t, validate_lang
from .merger import merge_and_audit
from .models import MergeReport
from .samples import create_dirty_demo_inputs


def _detect_cli_lang(argv: Optional[list[str]]) -> str:
    args_list = sys.argv[1:] if argv is None else argv
    for idx, arg in enumerate(args_list):
        if arg == "--lang" and idx + 1 < len(args_list):
            candidate = args_list[idx + 1].strip().lower()
            return candidate if candidate in SUPPORTED_LANGS else "tr"
        if arg.startswith("--lang="):
            candidate = arg.split("=", 1)[1].strip().lower()
            return candidate if candidate in SUPPORTED_LANGS else "tr"
    return "tr"


def _json_default(obj: object) -> object:
    if isinstance(obj, Decimal):
        return str(obj)
    raise TypeError(t("en", "cli_unserializable_type", type_name=type(obj).__name__))


def _write_json(report: MergeReport, path: Path) -> None:
    resolved = path.resolve()
    resolved.parent.mkdir(parents=True, exist_ok=True)
    payload = asdict(report)
    resolved.write_text(
        json.dumps(payload, ensure_ascii=False, indent=2, default=_json_default),
        encoding="utf-8",
    )


def _print_console_summary(report: MergeReport, xlsx_out: Path, lang: str = "tr") -> None:
    s = report.summary
    print(t(lang, "cli_summary_title"))
    print(f"{t(lang, 'cli_lbl_tables')}: {s.source_sheets}")
    print(f"{t(lang, 'cli_lbl_raw_rows')}: {s.total_input_rows}")
    print(f"{t(lang, 'cli_lbl_kept_rows')}: {s.kept_rows}")
    print(f"{t(lang, 'cli_lbl_dups_removed')}: {s.exact_duplicates_removed}")
    print(f"{t(lang, 'cli_lbl_conflicts')}: {s.conflicting_documents}")
    print(f"{t(lang, 'cli_lbl_status')}: {s.passed_rows} / {s.warning_rows} / {s.error_rows}")
    print(f"{t(lang, 'cli_lbl_raw_gross')}: {s.raw_brut_total:.2f}")
    print(f"{t(lang, 'cli_lbl_dup_gross')}: {s.duplicate_brut_removed:.2f}")
    print(f"{t(lang, 'cli_lbl_kept_gross')}: {s.kept_brut_total:.2f}")
    print(f"{t(lang, 'cli_lbl_excel_out')}: {xlsx_out}")


def build_parser(lang: str = "tr") -> argparse.ArgumentParser:
    """Create the argument parser with localized descriptions for the selected language."""
    active_lang = validate_lang(lang)
    parser = argparse.ArgumentParser(
        prog="excel-merge",
        description=t(active_lang, "cli_desc"),
    )
    parser.add_argument(
        "-V",
        "--version",
        action="version",
        version=f"%(prog)s {__version__}",
    )
    parser.add_argument(
        "--lang",
        choices=SUPPORTED_LANGS,
        default="tr",
        help=t(active_lang, "cli_lang_help"),
    )
    sub = parser.add_subparsers(dest="command", required=True)

    merge_p = sub.add_parser("merge", help=t(active_lang, "cli_merge_help"))
    merge_p.add_argument("inputs", nargs="+", help=t(active_lang, "cli_inputs_help"))
    merge_p.add_argument("-o", "--output", required=True, help=t(active_lang, "cli_output_help"))
    merge_p.add_argument("--json", help=t(active_lang, "cli_json_help"))
    merge_p.add_argument("--csv-out", dest="csv_out", help=t(active_lang, "cli_csv_help"))
    merge_p.add_argument("--sheet", help=t(active_lang, "cli_sheet_help"))
    merge_p.add_argument("--keep-duplicates", action="store_true", help=t(active_lang, "cli_keep_dups_help"))
    merge_p.add_argument("--fail-on-error", action="store_true", help=t(active_lang, "cli_fail_err_help"))
    merge_p.add_argument(
        "--lang",
        choices=SUPPORTED_LANGS,
        default=argparse.SUPPRESS,
        help=t(active_lang, "cli_lang_help"),
    )

    demo_p = sub.add_parser("demo", help=t(active_lang, "cli_demo_help"))
    demo_p.add_argument("--out-dir", default="demo_output", help=t(active_lang, "cli_demo_outdir_help"))
    demo_p.add_argument(
        "--lang",
        choices=SUPPORTED_LANGS,
        default=argparse.SUPPRESS,
        help=t(active_lang, "cli_lang_help"),
    )
    return parser


def main(argv: Optional[list[str]] = None) -> int:
    """CLI entry point supporting merge and demo subcommands in TR and EN."""
    hinted_lang = _detect_cli_lang(argv)
    parser = build_parser(lang=hinted_lang)
    args = parser.parse_args(argv)
    active_lang = validate_lang(args.lang)

    try:
        if args.command == "demo":
            layout = DEMO_LAYOUT[active_lang]
            out_dir = Path(args.out_dir).resolve()
            samples_dir = out_dir / layout["samples_dir"]
            sample_files = create_dirty_demo_inputs(samples_dir, lang=active_lang)
            report = merge_and_audit(sample_files, lang=active_lang)
            xlsx_out = out_dir / layout["xlsx_name"]
            export_merge_workbook(report, xlsx_out, lang=active_lang)
            _write_json(report, out_dir / layout["json_name"])
            _print_console_summary(report, xlsx_out, lang=active_lang)
            return 0

        input_paths = [Path(p) for p in args.inputs]
        report = merge_and_audit(
            input_paths,
            only_sheet=args.sheet,
            keep_duplicates=args.keep_duplicates,
            lang=active_lang,
        )
        if report.fatal_reason:
            print(f"{t(active_lang, 'cli_err_prefix')}: {report.fatal_reason}", file=sys.stderr)
            return 1

        xlsx_out = export_merge_workbook(report, Path(args.output), lang=active_lang)
        if args.json:
            _write_json(report, Path(args.json))
        if args.csv_out:
            export_clean_csv(report, Path(args.csv_out), lang=active_lang)
        _print_console_summary(report, xlsx_out, lang=active_lang)

        if args.fail_on_error and report.summary.error_rows > 0:
            return 2
        return 0
    except (FileNotFoundError, ValueError) as exc:
        print(f"{t(active_lang, 'cli_err_prefix')}: {exc}", file=sys.stderr)
        return 1
    except Exception as exc:
        print(f"{t(active_lang, 'cli_unexpected_err_prefix')}: {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
