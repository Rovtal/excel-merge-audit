"""Schema normalizer, smart header detector, Turkish/global Decimal and Date parsers."""

from datetime import datetime
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
import re
from typing import Optional

from .models import AuditIssue, NormalizedRecord, RawSheetTable


TR_MAP = str.maketrans("çğıöşüÇĞİÖŞÜ", "cgiosuCGIOSU")
WHITESPACE_RE = re.compile(r"\s+")
SCIENTIFIC_RE = re.compile(r"^\d+(?:[.,]\d+)?[eE][+-]?\d+$")
CURRENCY_RE = re.compile(r"(TL|TRY|USD|EUR|GBP|\$|€|₺)", re.IGNORECASE)

CANONICAL_ALIASES: dict[str, set[str]] = {
    "tarih": {
        "tarih",
        "fatura tarihi",
        "belge tarihi",
        "islem tarihi",
        "duzenleme tarihi",
        "date",
        "invoice date",
        "issue date",
    },
    "belge_no": {
        "belge no",
        "belge numarasi",
        "fatura no",
        "fatura numarasi",
        "fis no",
        "evrak no",
        "ettn",
        "invoice no",
        "doc no",
        "document no",
        "referans no",
    },
    "taraf": {
        "taraf",
        "cari",
        "cari unvan",
        "cari adi",
        "firma",
        "firma adi",
        "unvan",
        "musteri",
        "musteri adi",
        "tedarikci",
        "satici",
        "alici",
        "counterparty",
        "company",
        "customer",
        "vendor",
    },
    "vergi_no": {
        "vkn",
        "tckn",
        "vkn tckn",
        "vergi no",
        "vergi numarasi",
        "tc kimlik",
        "tax id",
        "tax no",
    },
    "aciklama": {
        "aciklama",
        "not",
        "detay",
        "hizmet",
        "urun",
        "kalem",
        "description",
        "notes",
        "memo",
    },
    "net_tutar": {
        "net tutar",
        "matrah",
        "ara toplam",
        "kdv haric",
        "net",
        "subtotal",
        "net amount",
    },
    "kdv_tutar": {
        "kdv",
        "kdv tutari",
        "vergi tutari",
        "vat",
        "tax amount",
    },
    "brut_tutar": {
        "genel toplam",
        "toplam tutar",
        "brut tutar",
        "tutar",
        "odenecek tutar",
        "toplam",
        "amount",
        "total",
        "gross total",
    },
    "para_birimi": {
        "para birimi",
        "doviz",
        "birim",
        "currency",
        "pb",
    },
}

ALIAS_LOOKUP: dict[str, str] = {}
for canon, aliases in CANONICAL_ALIASES.items():
    for alias in aliases:
        ALIAS_LOOKUP[alias] = canon


def normalize_header_token(raw: str) -> str:
    """Normalize arbitrary Turkish/English header labels to ASCII space-delimited keys."""
    ascii_str = raw.strip().translate(TR_MAP).lower()
    ascii_str = re.sub(r"[^a-z0-9]+", " ", ascii_str)
    return WHITESPACE_RE.sub(" ", ascii_str).strip()


def detect_header_row(
    rows: list[list[str]],
    scan_limit: int = 15,
) -> tuple[int, dict[int, str], dict[str, str]]:
    """Locate the real header row even when reports start with title/logo banner rows."""
    best_idx = -1
    best_score = 0
    best_col_map: dict[int, str] = {}
    best_label_map: dict[str, str] = {}

    for idx, row in enumerate(rows[:scan_limit]):
        if not any(cell.strip() for cell in row):
            continue
        col_map: dict[int, str] = {}
        label_map: dict[str, str] = {}
        seen_canon: set[str] = set()
        score = 0
        for col_idx, cell in enumerate(row):
            token = normalize_header_token(cell)
            if not token:
                continue
            canon = ALIAS_LOOKUP.get(token)
            if canon and canon not in seen_canon:
                seen_canon.add(canon)
                col_map[col_idx] = canon
                label_map[canon] = cell.strip()
                score += 2 if canon in ("belge_no", "brut_tutar", "tarih", "taraf") else 1

        if score > best_score:
            best_score = score
            best_idx = idx
            best_col_map = col_map
            best_label_map = label_map

    return best_idx, best_col_map, best_label_map


def parse_decimal_amount(raw: str) -> tuple[Optional[Decimal], Optional[str]]:
    """Parse TR/EN formatted currency/number strings strictly into quantized 2-decimal Decimal."""
    text = raw.strip()
    if not text:
        return Decimal("0.00"), None

    if text.startswith("#"):
        return None, f"Excel formül hatası içeriyor ({text})"

    detected_currency: Optional[str] = None
    cur_match = CURRENCY_RE.search(text)
    if cur_match:
        sym = cur_match.group(1).upper()
        detected_currency = "TRY" if sym in ("TL", "₺", "TRY") else (
            "USD" if sym == "$" else ("EUR" if sym == "€" else sym)
        )
        text = CURRENCY_RE.sub("", text).strip()

    negative = False
    if text.startswith("(") and text.endswith(")"):
        negative = True
        text = text[1:-1].strip()
    if text.startswith("-"):
        negative = not negative
        text = text[1:].strip()
    elif text.startswith("+"):
        text = text[1:].strip()

    text = text.replace("\u00a0", "").replace(" ", "")
    if not text:
        return None, f"Geçersiz sayısal tutar ({raw})"

    has_dot = "." in text
    has_comma = "," in text

    if has_dot and has_comma:
        last_dot = text.rfind(".")
        last_comma = text.rfind(",")
        if last_comma > last_dot:
            cleaned = text.replace(".", "").replace(",", ".")
        else:
            cleaned = text.replace(",", "")
    elif has_comma:
        parts = text.split(",")
        if len(parts) > 2:
            cleaned = text.replace(",", "")
        elif len(parts[1]) == 3 and len(parts[0]) >= 1 and parts[0].isdigit() and "." not in text:
            # Ambiguous e.g. 1,250 vs 1,2500 -> in CSV TR decimal comma usually has <=2 digits; 3 digits after comma in EN thousands
            cleaned = text.replace(",", "") if len(parts[0]) <= 3 and int(parts[0]) > 0 else text.replace(",", ".")
        else:
            cleaned = text.replace(",", ".")
    elif has_dot:
        parts = text.split(".")
        if len(parts) > 2:
            cleaned = text.replace(".", "")
        elif len(parts[1]) == 3 and len(parts[0]) >= 1 and parts[0].isdigit():
            # e.g. 1.250 in Turkish spreadsheets is 1250
            cleaned = text.replace(".", "")
        else:
            cleaned = text
    else:
        cleaned = text

    try:
        val = Decimal(cleaned).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
        if negative:
            val = -val
        return val, detected_currency
    except (InvalidOperation, ValueError):
        return None, f"Geçersiz tutar biçimi ({raw})"


def parse_date_value(raw: str) -> tuple[str, Optional[str]]:
    """Normalize TR/ISO/slash date strings cleanly to YYYY-MM-DD."""
    text = raw.strip()
    if not text:
        return "", "Eksik tarih"

    for fmt in (
        "%Y-%m-%d",
        "%d.%m.%Y",
        "%d/%m/%Y",
        "%d-%m-%Y",
        "%Y/%m/%d",
        "%Y-%m-%d %H:%M:%S",
        "%d.%m.%Y %H:%M:%S",
        "%d.%m.%Y %H:%M",
        "%d/%m/%Y %H:%M",
    ):
        try:
            return datetime.strptime(text, fmt).date().isoformat(), None
        except ValueError:
            continue

    return text, f"Tanınmayan tarih formatı ({text})"


def _normalize_id_token(val: str) -> str:
    return re.sub(r"[^A-Z0-9]", "", val.upper())


def _normalize_party_token(val: str) -> str:
    return normalize_header_token(val)


def normalize_table(
    table: RawSheetTable,
) -> tuple[list[NormalizedRecord], list[AuditIssue], int, dict[str, str]]:
    """Normalize a single RawSheetTable into canonical records and structured issues."""
    issues: list[AuditIssue] = []
    header_idx, col_map, label_map = detect_header_row(table.rows)
    if header_idx < 0 or not col_map:
        issues.append(
            AuditIssue(
                severity="Hata",
                code="BASLIK_BULUNAMADI",
                source_file=table.source_file,
                sheet_name=table.sheet_name,
                source_row=1,
                belge_no="-",
                message="Tabloda tanınabilir sütun başlığı (Belge No, Tutar, Tarih, Taraf) bulunamadı.",
            )
        )
        return [], issues, 0, {}

    formula_err_by_row: dict[int, list[str]] = {}
    for err_r, err_c, err_txt in table.formula_error_cells:
        formula_err_by_row.setdefault(err_r, []).append(f"Sütun {err_c}: {err_txt}")

    records: list[NormalizedRecord] = []
    for row_idx in range(header_idx + 1, len(table.rows)):
        row = table.rows[row_idx]
        excel_row_num = row_idx + 1
        if not any(cell.strip() for cell in row):
            continue

        def get_field(canon: str) -> str:
            for c_i, c_name in col_map.items():
                if c_name == canon and c_i < len(row):
                    return row[c_i].strip()
            return ""

        raw_belge = get_field("belge_no")
        raw_taraf = get_field("taraf")
        raw_tarih = get_field("tarih")
        raw_vergi = get_field("vergi_no")
        raw_aciklama = get_field("aciklama")
        raw_pb = get_field("para_birimi")

        footer_tokens = {
            normalize_header_token(v)
            for v in (raw_tarih, raw_belge, raw_taraf, raw_aciklama)
            if v
        }
        subtotal_labels = {"toplam", "genel toplam", "ara toplam", "subtotal", "grand total", "total"}
        _, tarih_probe_err = parse_date_value(raw_tarih)
        if (footer_tokens & subtotal_labels) and (not raw_tarih or tarih_probe_err is not None):
            continue

        row_status = "Geçti"
        belge_label = raw_belge or f"SATIR-{excel_row_num}"

        def flag(severity: str, code: str, msg: str) -> None:
            nonlocal row_status
            issues.append(
                AuditIssue(
                    severity=severity,
                    code=code,
                    source_file=table.source_file,
                    sheet_name=table.sheet_name,
                    source_row=excel_row_num,
                    belge_no=belge_label,
                    message=msg,
                )
            )
            if severity == "Hata":
                row_status = "Hata"
            elif severity == "Uyarı" and row_status == "Geçti":
                row_status = "Uyarı"

        if excel_row_num in formula_err_by_row:
            flag(
                "Hata",
                "FORMUL_HATASI",
                "Kaynak hücrede Excel formül hatası mevcut: " + ", ".join(formula_err_by_row[excel_row_num]),
            )

        if raw_vergi and SCIENTIFIC_RE.match(raw_vergi):
            flag(
                "Uyarı",
                "BILIMSEL_GOSTERIM_VKN",
                f"Vergi/Kimlik numarası bilimsel gösterimle bozulmuş ({raw_vergi}).",
            )
        if raw_belge and SCIENTIFIC_RE.match(raw_belge):
            flag(
                "Uyarı",
                "BILIMSEL_GOSTERIM_BELGE",
                f"Belge numarası bilimsel gösterimle bozulmuş ({raw_belge}).",
            )

        norm_date, date_err = parse_date_value(raw_tarih)
        if date_err and "tarih" in col_map.values():
            flag("Uyarı", "TARIH_FORMAT_HATASI", date_err)

        detected_currency: Optional[str] = None

        def parse_money_field(canon: str) -> tuple[Optional[Decimal], bool]:
            nonlocal detected_currency
            present = canon in col_map.values()
            if not present:
                return None, False
            val_str = get_field(canon)
            val, meta = parse_decimal_amount(val_str)
            if val is None:
                flag("Hata", "TUTAR_FORMAT_HATASI", f"{canon}: {meta}")
                return Decimal("0.00"), True
            if meta and not detected_currency:
                detected_currency = meta
            return val, True

        net_val, has_net = parse_money_field("net_tutar")
        kdv_val, has_kdv = parse_money_field("kdv_tutar")
        brut_val, has_brut = parse_money_field("brut_tutar")

        if not has_brut and (has_net or has_kdv):
            brut_val = (net_val or Decimal("0.00")) + (kdv_val or Decimal("0.00"))
        elif has_brut and has_net and has_kdv:
            expected_brut = (net_val or Decimal("0.00")) + (kdv_val or Decimal("0.00"))
            diff = abs(expected_brut - (brut_val or Decimal("0.00")))
            if diff > Decimal("0.02"):
                flag(
                    "Hata",
                    "TOPLAM_UYUMSUZLUK",
                    f"Net ({net_val}) + KDV ({kdv_val}) = {expected_brut}, ancak Brüt Tutar {brut_val} (Fark: {diff}).",
                )
        elif has_brut and not has_net and not has_kdv:
            net_val = brut_val
            kdv_val = Decimal("0.00")

        net_final = net_val if net_val is not None else Decimal("0.00")
        kdv_final = kdv_val if kdv_val is not None else Decimal("0.00")
        brut_final = brut_val if brut_val is not None else Decimal("0.00")

        if not raw_belge and "belge_no" in col_map.values():
            flag("Uyarı", "EKSIK_BELGE_NO", "Belge/Fatura numarası boş.")
        if not raw_taraf and "taraf" in col_map.values():
            flag("Uyarı", "EKSIK_TARAF", "Cari / Taraf unvanı boş.")

        pb_final = raw_pb.upper() if raw_pb else (detected_currency or "TRY")
        id_tok = _normalize_id_token(raw_belge)
        extra_id = _normalize_party_token(raw_aciklama) if not id_tok else ""
        dedupe_key = "|".join(
            [
                norm_date,
                id_tok,
                _normalize_party_token(raw_taraf),
                extra_id,
                f"{brut_final:.2f}",
                pb_final,
            ]
        )

        records.append(
            NormalizedRecord(
                source_file=table.source_file,
                sheet_name=table.sheet_name,
                source_row=excel_row_num,
                tarih=norm_date,
                belge_no=raw_belge,
                taraf=raw_taraf,
                vergi_no=raw_vergi,
                aciklama=raw_aciklama,
                net_tutar=net_final,
                kdv_tutar=kdv_final,
                brut_tutar=brut_final,
                para_birimi=pb_final,
                durum=row_status,
                dedupe_key=dedupe_key,
            )
        )

    return records, issues, header_idx + 1, label_map
