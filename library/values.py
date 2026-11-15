"""Werte lesen und formatieren: Beträge, Datumsangaben, Zeiträume, Währungen, IBAN."""

import re
from datetime import date, datetime, timedelta
from decimal import Decimal, ROUND_HALF_UP

from .constants import CURRENCIES, MONTH_NAMES


CENT = Decimal("0.01")


def to_decimal(v) -> Decimal:
    return Decimal(str(v))


def fmt_amount(v: Decimal) -> str:
    return format(v.quantize(CENT, rounding=ROUND_HALF_UP), "f")


def fmt_num(v: Decimal) -> str:
    """Preis/Menge: mindestens 2, höchstens 8 Nachkommastellen (ohne überflüssige Nullen)."""
    t = format(v.quantize(Decimal("0.00000001"), rounding=ROUND_HALF_UP), "f").rstrip("0")
    whole, _, frac = t.partition(".")
    return f"{whole}.{frac.ljust(2, '0')}"


def to_yyyymmdd(v) -> str:
    """date-Objekt, JJJJ-MM-TT, TT.MM.JJJJ oder JJJJMMTT -> JJJJMMTT (Format 102)."""
    if isinstance(v, (date, datetime)):
        return v.strftime("%Y%m%d")
    s = str(v).strip()
    for fmt in ("%Y-%m-%d", "%d.%m.%Y", "%Y%m%d"):
        try:
            return datetime.strptime(s, fmt).strftime("%Y%m%d")
        except ValueError:
            pass
    raise ValueError(f"Ungültiges Datum: {v!r}")


def parse_amount(s: str) -> Decimal:
    """'EUR 1.234,56' / '3.000,-' / 'GBP 2,345.67' -> Decimal"""
    t = re.sub(r"[^\d,.\-]", "", s)
    t = re.sub(r"[,.]-+$", "", t)                      # 3.000,-  -> 3.000
    neg = t.startswith("-")
    t = t.strip("-")
    if "," in t and "." in t:                          # letztes Trennzeichen = Dezimaltrenner
        dec = "," if t.rfind(",") > t.rfind(".") else "."
        t = t.replace("." if dec == "," else ",", "").replace(",", ".")
    elif "," in t:
        t = t.replace(",", ".")
    elif "." in t and re.fullmatch(r"\d{1,3}(?:\.\d{3})+", t):
        t = t.replace(".", "")                         # 3.000 = dreitausend
    if not t:
        raise ValueError(f"Kein Betrag in {s!r}")
    return -Decimal(t) if neg else Decimal(t)


def detect_currency(s: str):
    m = re.search(r"\b([A-Z]{3})\b|([€£$])", s)
    if not m or not re.search(r"\d", s):
        return None
    code = m.group(1) or m.group(2)
    return CURRENCIES.get(code, code if m.group(1) else None)


def parse_date(text: str):
    """'1. November 2026', '1 November 2026', 'November 1, 2026', '1.11.2026', '2026-11-1'."""
    if not text:
        return None
    names = "|".join(sorted(MONTH_NAMES, key=len, reverse=True))
    tries = [(r"(\d{1,2})\.?\s*(" + names + r")\.?\s+(\d{4})", (1, 2, 3)),
             (r"\b(" + names + r")\.?\s+(\d{1,2})(?:st|nd|rd|th)?,?\s+(\d{4})", (2, 1, 3))]
    for pat, (di, mi, yi) in tries:
        m = re.search(pat, text, re.I)
        if m:
            try:
                return date(int(m.group(yi)), MONTH_NAMES[m.group(mi).lower()], int(m.group(di)))
            except ValueError:
                pass
    if (m := re.search(r"(\d{1,2})\.(\d{1,2})\.(\d{4})", text)):
        return date(int(m.group(3)), int(m.group(2)), int(m.group(1)))
    if (m := re.search(r"(\d{4})-(\d{2})-(\d{2})", text)):
        return date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
    return None


def iban_ok(iban: str) -> bool:
    s = iban.replace(" ", "").upper()
    if not re.fullmatch(r"[A-Z]{2}\d{2}[A-Z0-9]{11,30}", s):
        return False
    return int("".join(str(int(c, 36)) for c in s[4:] + s[:4])) % 97 == 1


def month_period(mon: str, yr: str):
    mo, y = MONTH_NAMES[mon.lower()], int(yr)
    return date(y, mo, 1), date(y + (mo == 12), mo % 12 + 1, 1) - timedelta(days=1)


def parse_period(text: str, fallback_year: int = None):
    """Zeitraum aus Text: '20.8.2026-19.9.2026', '1.-10.11.2026', '15.-31.10', 'Oktober 2026', 'October 2026'."""
    if not text:
        return None
    sep = r"\s*(?:-|–|—|bis|to|until)\s*"
    m = re.search(r"(\d{1,2}\.\d{1,2}\.\d{4})" + sep + r"(\d{1,2}\.\d{1,2}\.\d{4})", text)
    if m:
        a, b = parse_date(m.group(1)), parse_date(m.group(2))
        return (a, b) if a and b and a <= b else None
    months = re.findall(r"\b(" + "|".join(MONTH_NAMES) + r")\s+(\d{4})\b", text, re.I)
    m = re.search(r"\b(\d{1,2})\.?" + sep + r"(\d{1,2})\.(\d{1,2})\.(\d{4})?", text)
    if m:
        year = int(m.group(4)) if m.group(4) else (int(months[0][1]) if months else fallback_year)
        if year:
            try:
                a, b = date(year, int(m.group(3)), int(m.group(1))), date(year, int(m.group(3)), int(m.group(2)))
                if a <= b:
                    return a, b
            except ValueError:
                pass
    if len({(x.lower(), y) for x, y in months}) == 1:
        return month_period(*months[0])
    return None
