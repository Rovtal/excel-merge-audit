"""Offline secure readers for .xlsx (OpenXML via lxml) and .csv/.tsv files."""

import csv
from datetime import date, timedelta
import io
import posixpath
import re
from pathlib import Path
from typing import Optional
import zipfile

from lxml import etree

from .i18n import t, validate_lang
from .models import RawSheetTable


MAX_FILE_BYTES = 50 * 1024 * 1024
MAX_ZIP_ENTRIES = 2000
MAX_ZIP_UNCOMPRESSED_BYTES = 150 * 1024 * 1024
OLE2_MAGIC = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"

NS_SS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS_REL_ATTR = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_REL_PKG = "http://schemas.openxmlformats.org/package/2006/relationships"

BUILTIN_DATE_FMT_IDS = {14, 15, 16, 17, 18, 19, 20, 21, 22, 45, 46, 47}
CELL_REF_RE = re.compile(r"^([A-Z]+)(\d+)$", re.IGNORECASE)
SAFE_XML_PARSER = etree.XMLParser(
    resolve_entities=False,
    no_network=True,
    huge_tree=False,
)


def _parse_xml_bytes(data: bytes, source_label: str, lang: str = "tr") -> etree._Element:
    try:
        return etree.fromstring(data, parser=SAFE_XML_PARSER)
    except etree.XMLSyntaxError as exc:
        raise ValueError(t(lang, "reader_invalid_xml", source=source_label, exc=exc)) from exc


def _col_letters_to_index(letters: str) -> int:
    idx = 0
    for ch in letters.upper():
        idx = idx * 26 + (ord(ch) - ord("A") + 1)
    return idx - 1


def _is_date_format_code(code: str) -> bool:
    cleaned = re.sub(r'"[^"]*"', "", code)
    cleaned = re.sub(r"\[[^\]]*\]", "", cleaned)
    cleaned = re.sub(r"\\.", "", cleaned).lower()
    return ("yy" in cleaned or "dd" in cleaned or "mm" in cleaned) and (
        "0" not in cleaned and "#" not in cleaned
    )


def _serial_to_iso_date(serial_val: float, use_1904: bool = False) -> Optional[str]:
    if serial_val < 0 or serial_val > 2958465:
        return None
    days = int(serial_val)
    if use_1904:
        target = date(1904, 1, 1) + timedelta(days=days)
        return target.isoformat()
    if days == 0:
        return None
    if days < 60:
        target = date(1899, 12, 31) + timedelta(days=days)
    else:
        target = date(1899, 12, 30) + timedelta(days=days)
    return target.isoformat()


def _validate_zip_archive(zf: zipfile.ZipFile, file_name: str, lang: str = "tr") -> None:
    infos = zf.infolist()
    if len(infos) > MAX_ZIP_ENTRIES:
        raise ValueError(t(lang, "reader_zip_too_many", file_name=file_name, count=len(infos)))
    total_size = 0
    for info in infos:
        normalized = info.filename.replace("\\", "/")
        if normalized.startswith("/") or ".." in Path(normalized).parts:
            raise ValueError(t(lang, "reader_zip_unsafe_path", file_name=file_name, member=info.filename))
        total_size += info.file_size
        if total_size > MAX_ZIP_UNCOMPRESSED_BYTES:
            raise ValueError(t(lang, "reader_zip_too_large", file_name=file_name))


def _normalize_zip_target(target: str) -> str:
    clean = target.replace("\\", "/").lstrip("/")
    if not clean.startswith("xl/"):
        clean = posixpath.normpath(posixpath.join("xl", clean))
    return clean.lstrip("/")


def _read_shared_strings(
    zf: zipfile.ZipFile,
    file_name: str = "sharedStrings.xml",
    lang: str = "tr",
) -> list[str]:
    if "xl/sharedStrings.xml" not in zf.namelist():
        return []
    root = _parse_xml_bytes(zf.read("xl/sharedStrings.xml"), f"{file_name}:xl/sharedStrings.xml", lang=lang)
    values: list[str] = []
    for si in root.findall(f"{{{NS_SS}}}si"):
        parts = [t_el.text or "" for t_el in si.iter(f"{{{NS_SS}}}t")]
        values.append("".join(parts))
    return values


def _read_date_style_indexes(
    zf: zipfile.ZipFile,
    file_name: str = "styles.xml",
    lang: str = "tr",
) -> set[int]:
    if "xl/styles.xml" not in zf.namelist():
        return set()
    root = _parse_xml_bytes(zf.read("xl/styles.xml"), f"{file_name}:xl/styles.xml", lang=lang)
    custom_date_ids: set[int] = set()
    num_fmts = root.find(f"{{{NS_SS}}}numFmts")
    if num_fmts is not None:
        for fmt in num_fmts.findall(f"{{{NS_SS}}}numFmt"):
            fmt_id_raw = fmt.get("numFmtId")
            code = fmt.get("formatCode") or ""
            if fmt_id_raw and fmt_id_raw.isdigit() and _is_date_format_code(code):
                custom_date_ids.add(int(fmt_id_raw))

    date_xfs: set[int] = set()
    cell_xfs = root.find(f"{{{NS_SS}}}cellXfs")
    if cell_xfs is not None:
        for idx, xf in enumerate(cell_xfs.findall(f"{{{NS_SS}}}xf")):
            fmt_id_raw = xf.get("numFmtId")
            if fmt_id_raw and fmt_id_raw.isdigit():
                fmt_id = int(fmt_id_raw)
                if fmt_id in BUILTIN_DATE_FMT_IDS or fmt_id in custom_date_ids:
                    date_xfs.add(idx)
    return date_xfs


def _parse_xlsx_worksheet(
    xml_bytes: bytes,
    shared_strings: list[str],
    date_xfs: set[int],
    use_1904: bool,
    source_file: str,
    sheet_name: str,
    lang: str = "tr",
) -> RawSheetTable:
    root = _parse_xml_bytes(xml_bytes, f"{Path(source_file).name}:{sheet_name}", lang=lang)
    sheet_data = root.find(f"{{{NS_SS}}}sheetData")
    if sheet_data is None:
        return RawSheetTable(
            source_file=source_file,
            sheet_name=sheet_name,
            file_format="xlsx",
            encoding="utf-8",
            delimiter="",
            rows=[],
        )

    row_map: dict[int, dict[int, str]] = {}
    max_row = 0
    max_col = -1
    formula_errors: list[tuple[int, int, str]] = []
    seq_row = 0

    for row_el in sheet_data.findall(f"{{{NS_SS}}}row"):
        r_attr = row_el.get("r")
        seq_row = int(r_attr) if (r_attr and r_attr.isdigit()) else seq_row + 1
        max_row = max(max_row, seq_row)
        row_map.setdefault(seq_row, {})
        seq_col = -1

        for c_el in row_el.findall(f"{{{NS_SS}}}c"):
            ref = c_el.get("r")
            if ref:
                match = CELL_REF_RE.match(ref)
                if match:
                    col_idx = _col_letters_to_index(match.group(1))
                    row_num = int(match.group(2))
                else:
                    seq_col += 1
                    col_idx = seq_col
                    row_num = seq_row
            else:
                seq_col += 1
                col_idx = seq_col
                row_num = seq_row

            seq_col = col_idx
            max_row = max(max_row, row_num)
            max_col = max(max_col, col_idx)
            target_cells = row_map.setdefault(row_num, {})

            cell_type = c_el.get("t", "n")
            style_raw = c_el.get("s")
            style_idx = int(style_raw) if (style_raw and style_raw.isdigit()) else -1

            if cell_type == "inlineStr":
                is_el = c_el.find(f"{{{NS_SS}}}is")
                text_val = (
                    "".join(t_el.text or "" for t_el in is_el.iter(f"{{{NS_SS}}}t"))
                    if is_el is not None
                    else ""
                )
            else:
                v_el = c_el.find(f"{{{NS_SS}}}v")
                raw_val = v_el.text if (v_el is not None and v_el.text is not None) else ""
                if cell_type == "s":
                    if raw_val.isdigit():
                        s_idx = int(raw_val)
                        text_val = shared_strings[s_idx] if s_idx < len(shared_strings) else ""
                    else:
                        text_val = raw_val
                elif cell_type == "b":
                    text_val = t(lang, "reader_bool_true" if raw_val == "1" else "reader_bool_false")
                elif cell_type == "e":
                    text_val = raw_val or "#ERROR!"
                    formula_errors.append((row_num, col_idx + 1, text_val))
                elif cell_type == "str":
                    text_val = raw_val
                else:
                    if raw_val and style_idx in date_xfs:
                        try:
                            iso_date = _serial_to_iso_date(float(raw_val), use_1904=use_1904)
                            text_val = iso_date if iso_date is not None else raw_val
                        except ValueError:
                            text_val = raw_val
                    else:
                        text_val = raw_val

            target_cells[col_idx] = text_val.strip()

    if max_row == 0:
        return RawSheetTable(
            source_file=source_file,
            sheet_name=sheet_name,
            file_format="xlsx",
            encoding="utf-8",
            delimiter="",
            rows=[],
            formula_error_cells=formula_errors,
        )

    width = max(max_col + 1, 1)
    grid: list[list[str]] = []
    for r_num in range(1, max_row + 1):
        cells = row_map.get(r_num, {})
        grid.append([cells.get(c_idx, "") for c_idx in range(width)])

    return RawSheetTable(
        source_file=source_file,
        sheet_name=sheet_name,
        file_format="xlsx",
        encoding="utf-8",
        delimiter="",
        rows=grid,
        formula_error_cells=formula_errors,
    )


def read_xlsx_tables(
    file_path: Path,
    only_sheet: Optional[str] = None,
    lang: str = "tr",
) -> list[RawSheetTable]:
    """Read sheets from an OpenXML .xlsx workbook without openpyxl/pandas."""
    active_lang = validate_lang(lang)
    resolved = file_path.resolve()
    if resolved.stat().st_size > MAX_FILE_BYTES:
        raise ValueError(t(active_lang, "reader_file_too_large", file_name=resolved.name))

    with resolved.open("rb") as fh:
        header = fh.read(8)
    if header == OLE2_MAGIC:
        raise ValueError(t(active_lang, "reader_ole2_xlsx", file_name=resolved.name))

    if not zipfile.is_zipfile(resolved):
        raise ValueError(t(active_lang, "reader_invalid_zip", file_name=resolved.name))

    with zipfile.ZipFile(resolved, "r") as zf:
        _validate_zip_archive(zf, resolved.name, lang=active_lang)
        names = set(zf.namelist())
        if "xl/workbook.xml" not in names or "xl/_rels/workbook.xml.rels" not in names:
            raise ValueError(t(active_lang, "reader_corrupt_xlsx", file_name=resolved.name))

        wb_root = _parse_xml_bytes(zf.read("xl/workbook.xml"), f"{resolved.name}:xl/workbook.xml", lang=active_lang)
        wb_pr = wb_root.find(f"{{{NS_SS}}}workbookPr")
        use_1904 = bool(
            wb_pr is not None and wb_pr.get("date1904", "").lower() in ("1", "true")
        )

        rels_root = _parse_xml_bytes(
            zf.read("xl/_rels/workbook.xml.rels"),
            f"{resolved.name}:xl/_rels/workbook.xml.rels",
            lang=active_lang,
        )
        rel_map: dict[str, str] = {}
        for rel in rels_root.findall(f"{{{NS_REL_PKG}}}Relationship"):
            rid = rel.get("Id")
            target = rel.get("Target")
            if rid and target:
                rel_map[rid] = _normalize_zip_target(target)

        shared_strings = _read_shared_strings(zf, resolved.name, lang=active_lang)
        date_xfs = _read_date_style_indexes(zf, resolved.name, lang=active_lang)

        tables: list[RawSheetTable] = []
        sheets_el = wb_root.find(f"{{{NS_SS}}}sheets")
        if sheets_el is None:
            return []

        for sheet_el in sheets_el.findall(f"{{{NS_SS}}}sheet"):
            name = sheet_el.get("name") or t(active_lang, "reader_default_sheet")
            if only_sheet and name != only_sheet:
                continue
            rid = sheet_el.get(f"{{{NS_REL_ATTR}}}id")
            if not rid or rid not in rel_map:
                continue
            member = rel_map[rid]
            if member not in names:
                continue
            table = _parse_xlsx_worksheet(
                xml_bytes=zf.read(member),
                shared_strings=shared_strings,
                date_xfs=date_xfs,
                use_1904=use_1904,
                source_file=str(resolved),
                sheet_name=name,
                lang=active_lang,
            )
            tables.append(table)

    if only_sheet and not tables:
        raise ValueError(t(active_lang, "reader_missing_sheet", file_name=resolved.name, sheet=only_sheet))
    return tables


def _decode_csv_bytes(raw_bytes: bytes, file_name: str, lang: str = "tr") -> tuple[str, str]:
    if raw_bytes.startswith(OLE2_MAGIC):
        raise ValueError(t(lang, "reader_binary_csv", file_name=file_name))
    if raw_bytes.startswith(b"PK\x03\x04"):
        raise ValueError(t(lang, "reader_zip_in_csv", file_name=file_name))

    candidates = ("utf-8-sig", "utf-8", "cp1254", "iso-8859-9", "cp1252")
    for enc in candidates:
        try:
            text = raw_bytes.decode(enc)
            if enc in ("utf-8-sig", "utf-8") and "\ufffd" in text:
                continue
            return text, enc
        except UnicodeDecodeError:
            continue
    raise ValueError(t(lang, "reader_decode_failed", file_name=file_name))


def _detect_delimiter(text: str, ext: str) -> str:
    sample_lines = [ln for ln in text.splitlines() if ln.strip()][:15]
    if not sample_lines:
        return "\t" if ext == ".tsv" else ";"
    sample_block = "\n".join(sample_lines)

    candidates = ["\t", ";", ",", "|"] if ext == ".tsv" else [";", ",", "\t", "|"]
    best_delim = candidates[0]
    best_score = -1.0

    for delim in candidates:
        rows = list(csv.reader(io.StringIO(sample_block), delimiter=delim))
        widths = [len(r) for r in rows if any(cell.strip() for cell in r)]
        if not widths:
            continue
        max_w = max(widths)
        if max_w <= 1:
            continue
        consistent = sum(1 for w in widths if w == max_w)
        score = consistent * 10 + max_w
        if score > best_score:
            best_score = score
            best_delim = delim

    return best_delim


def read_delimited_table(file_path: Path, lang: str = "tr") -> RawSheetTable:
    """Read a CSV or TSV file with Turkish/UTF-8 encoding fallback and smart delimiter inference."""
    active_lang = validate_lang(lang)
    resolved = file_path.resolve()
    if resolved.stat().st_size > MAX_FILE_BYTES:
        raise ValueError(t(active_lang, "reader_file_too_large", file_name=resolved.name))

    raw_bytes = resolved.read_bytes()
    text, encoding = _decode_csv_bytes(raw_bytes, resolved.name, lang=active_lang)
    delimiter = _detect_delimiter(text, resolved.suffix.lower())
    rows = [
        [cell.strip() for cell in row]
        for row in csv.reader(io.StringIO(text), delimiter=delimiter)
    ]

    return RawSheetTable(
        source_file=str(resolved),
        sheet_name="CSV",
        file_format=resolved.suffix.lower().lstrip(".") or "csv",
        encoding=encoding,
        delimiter=repr(delimiter) if delimiter == "\t" else delimiter,
        rows=rows,
    )


def load_input_tables(
    inputs: list[Path],
    only_sheet: Optional[str] = None,
    lang: str = "tr",
) -> tuple[list[RawSheetTable], list[str]]:
    """Expand files/directories cleanly and load all supported .xlsx/.csv/.tsv tables."""
    active_lang = validate_lang(lang)
    tables: list[RawSheetTable] = []
    ignored: list[str] = []
    seen_files: set[Path] = set()

    expanded: list[Path] = []
    for raw_item in inputs:
        p = raw_item.resolve()
        if not p.exists():
            raise FileNotFoundError(t(active_lang, "reader_missing_input", raw_item=raw_item))
        if p.is_dir():
            for child in sorted(p.rglob("*")):
                if child.is_file():
                    expanded.append(child.resolve())
        elif p.is_file():
            expanded.append(p)

    for file_path in expanded:
        if file_path in seen_files:
            continue
        seen_files.add(file_path)

        if file_path.name.startswith("~$"):
            ignored.append(t(active_lang, "reader_ignored_lock", file_name=file_path.name))
            continue

        suffix = file_path.suffix.lower()
        if suffix == ".xlsx":
            tables.extend(read_xlsx_tables(file_path, only_sheet=only_sheet, lang=active_lang))
        elif suffix in (".csv", ".tsv"):
            tables.append(read_delimited_table(file_path, lang=active_lang))
        elif suffix == ".xls":
            raise ValueError(t(active_lang, "reader_legacy_xls", file_name=file_path.name))
        else:
            ignored.append(t(active_lang, "reader_ignored_ext", file_name=file_path.name))

    return tables, ignored
