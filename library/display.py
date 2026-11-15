"""Übersicht der erkannten Rechnungsbestandteile im Terminal (sortiert nach BT-/BG-Nummer)."""

import re
from datetime import date, datetime
from decimal import Decimal

from .values import fmt_amount, fmt_num


BT_NAMES = {
    "BT-1": "Rechnungsnummer", "BT-2": "Rechnungsdatum", "BT-3": "Rechnungsart", "BT-5": "Währung",
    "BT-9": "Fälligkeitsdatum", "BT-10": "Käuferreferenz", "BT-11": "Projekt", "BT-12": "Vertrag",
    "BT-13": "Bestellnummer", "BT-20": "Zahlungsbedingungen", "BT-22": "Bemerkung",
    "BT-27": "Name", "BT-28": "Handelsname", "BT-29": "Verkäuferkennung", "BT-30": "Registernummer",
    "BT-31": "USt-IdNr.", "BT-32": "Steuernummer", "BT-34": "Elektronische Adresse", "BT-35": "Straße",
    "BT-36": "Adresszusatz", "BT-37": "Ort", "BT-38": "PLZ", "BT-39": "Region", "BT-40": "Land",
    "BT-41": "Ansprechpartner", "BT-42": "Telefon", "BT-43": "E-Mail",
    "BT-44": "Name", "BT-46": "Käuferkennung", "BT-47": "Registernummer", "BT-48": "USt-IdNr.",
    "BT-49": "Elektronische Adresse", "BT-50": "Straße", "BT-51": "Adresszusatz", "BT-52": "Ort",
    "BT-53": "PLZ", "BT-54": "Region", "BT-55": "Land", "BT-56": "Ansprechpartner", "BT-57": "Telefon",
    "BT-58": "E-Mail", "BT-72": "Leistungsdatum", "BT-73": "Beginn", "BT-74": "Ende",
    "BT-81": "Zahlungsart", "BT-82": "Zahlungsart (Text)", "BT-83": "Verwendungszweck", "BT-84": "IBAN",
    "BT-85": "Kontoinhaber", "BT-86": "BIC",
    "BT-106": "Summe Positionen netto", "BT-109": "Gesamtbetrag netto", "BT-110": "Umsatzsteuer gesamt",
    "BT-112": "Gesamtbetrag brutto", "BT-113": "Bereits gezahlt", "BT-115": "Zahlbetrag",
    "BT-116": "Bemessungsgrundlage", "BT-117": "Steuerbetrag", "BT-118": "Steuerkategorie",
    "BT-119": "Steuersatz %", "BT-120": "Befreiungsgrund", "BT-121": "Befreiungsgrund-Code",
    "BT-126": "Positionsnummer", "BT-127": "Positionshinweis", "BT-129": "Menge", "BT-130": "Einheit",
    "BT-131": "Nettobetrag", "BT-134": "Leistung von", "BT-135": "Leistung bis", "BT-146": "Einzelpreis netto",
    "BT-151": "Steuerkategorie", "BT-152": "Steuersatz %", "BT-153": "Bezeichnung", "BT-154": "Beschreibung",
    "BT-155": "Artikelnummer",
}


BG_NAMES = {"BG-4": "Verkäufer", "BG-7": "Käufer", "BG-14": "Rechnungszeitraum", "BG-16": "Zahlungsanweisung",
            "BG-22": "Gesamtbeträge (berechnet)", "BG-23": "Umsatzsteueraufschlüsselung", "BG-25": "Rechnungsposition"}


CODE_TEXT = {
    "BT-3": {380: "Rechnung", 381: "Gutschrift", 384: "Korrekturrechnung", 389: "Selbstfakturierung"},
    "BT-81": {58: "SEPA-Überweisung", 30: "Überweisung", 59: "SEPA-Lastschrift", 10: "bar"},
    "BT-130": {"C62": "Stück", "HUR": "Stunde", "DAY": "Tag", "MON": "Monat", "LS": "Pauschale", "KMT": "km"},
    "BT-118": {"S": "Normalsatz", "Z": "Nullsatz", "E": "steuerbefreit", "AE": "Reverse Charge",
               "K": "innergemeinschaftliche Lieferung", "G": "Ausfuhr", "O": "nicht steuerbar"},
}


CODE_TEXT["BT-151"] = CODE_TEXT["BT-118"]


def num_key(k: str):
    m = re.match(r"(BT|BG)-(\d+)", k)
    return (int(m.group(2)), 0 if m.group(1) == "BG" else 1) if m else (9999, 0)


def show_invoice(data: dict) -> None:
    """Gibt alle erkannten bzw. berechneten Rechnungsbestandteile sortiert nach BT-/BG-Nummer aus."""
    o_mode = data.get("_O")

    def de(x: str) -> str:                       # 10000.00 -> 10.000,00
        whole, _, frac = x.partition(".")
        sign = "-" if whole.startswith("-") else ""
        whole = f"{int(whole.lstrip('-') or 0):,}".replace(",", ".")
        return sign + whole + ("," + frac if frac else "")

    def fmt(k, v):
        if k in ("BT-119", "BT-152") and o_mode:
            return "–  (bei Kategorie O kein Steuersatz)"
        if isinstance(v, (date, datetime)):
            s_ = v.strftime("%d.%m.%Y")
        elif isinstance(v, bool):
            s_ = str(v)
        elif isinstance(v, (Decimal, float, int)) and k not in ("BT-1", "BT-3", "BT-81", "BT-126"):
            d_ = Decimal(str(v))
            s_ = de(fmt_num(d_) if k in ("BT-129", "BT-146") else fmt_amount(d_))
            if k in ("BT-119", "BT-152"):
                s_ = format(d_.normalize(), "f").replace(".", ",")
            if k == "BT-129" and d_ == d_.to_integral():
                s_ = str(int(d_))
        else:
            s_ = str(v)
        text = CODE_TEXT.get(k, {}).get(v) or CODE_TEXT.get(k, {}).get(str(v))
        if text:
            s_ += f"  ({text})"
        if o_mode and k in ("BT-31", "BT-48"):
            s_ += "  [nicht im XML: Kategorie O]"
        return s_

    def line(k, v, indent):
        name = BT_NAMES.get(k, "")
        print(f"{'  ' * indent}{k:<7} {name:<26} {fmt(k, v)}")

    groups = dict(data)
    if groups.get("_BG-23"):
        groups["BG-23"] = groups.pop("_BG-23")
    if o_mode and groups.get("BG-4") and not groups["BG-4"].get("BT-29") and groups["BG-4"].get("BT-32"):
        groups["BG-4"] = {**groups["BG-4"], "BT-29": f"{groups['BG-4']['BT-32']}  [= Steuernummer, Kategorie O]"}
    print("\nErkannte Rechnungsbestandteile (EN 16931):")
    print("-" * 78)
    for k in sorted((k for k in groups if not k.startswith("_")), key=num_key):
        v = groups[k]
        if k.startswith("BG-"):
            items = v if isinstance(v, list) else [v]
            for i, item in enumerate(items, 1):
                title = BG_NAMES.get(k, "")
                print(f"{k:<7} {title}{f' {i}' if len(items) > 1 else ''}")
                for kk in sorted(item, key=num_key):
                    line(kk, item[kk], 1)
        else:
            line(k, v, 0)
    print("-" * 78)
