"""Rechnungsdaten aus einer Word-Rechnung (.docx) auslesen (deutsch/englisch)."""

import re
import zipfile
from datetime import timedelta
from decimal import Decimal, ROUND_HALF_UP

from lxml import etree

from .constants import (COUNTRIES, EU_COUNTRIES, NOT_TAXABLE, PLZ_RE, TAX_EXEMPTIONS, UNITS, VAT_ID_RE, W_NS)
from .errors import InvoiceError
from .values import CENT, detect_currency, iban_ok, parse_amount, parse_date, parse_period


def docx_paragraphs(path: str):
    """Absatztexte in Dokumentreihenfolge: (Hauptteil, Kopfzeilen). Berücksichtigt Tabellen und Textfelder."""
    ns = {"w": W_NS}

    def paras(xml: bytes):
        root = etree.fromstring(xml)
        out = []
        for p in root.iter("{%s}p" % W_NS):
            if p.xpath("ancestor::*[local-name()='Fallback']"):
                continue  # Textfelder stehen in Word doppelt (Choice/Fallback)
            txt = ""
            for n in p.iter():
                if n.xpath("ancestor::w:p[1]", namespaces=ns)[:1] != [p] and n is not p:
                    continue  # Text verschachtelter Absätze (Textfelder) nicht doppelt zählen
                if n.tag == "{%s}t" % W_NS:
                    txt += n.text or ""
                elif n.tag == "{%s}tab" % W_NS:
                    txt += "\t"
                elif n.tag in ("{%s}br" % W_NS, "{%s}cr" % W_NS):
                    txt += "\n"
            out.append(txt.strip())
        return out

    with zipfile.ZipFile(path) as z:
        body = paras(z.read("word/document.xml"))
        head = []
        for name in sorted(n for n in z.namelist() if re.match(r"word/header\d*\.xml", n)):
            head += [t for t in paras(z.read(name)) if t and t not in head]
    return body, head


def find_labels(lines: list) -> dict:
    """'Label: Wert' (gleiche Zeile oder Wert in der Folgezeile, z. B. zweispaltige Tabellen) -> {label: wert}."""
    out = {}
    for i, t in enumerate(lines):
        m = re.match(r"^([^:\t\d][^:\t]{1,40}):\s*(.*)$", t)
        if not m:
            continue
        label, val = m.group(1).strip().lower(), m.group(2).strip()
        if not val and i + 1 < len(lines) and not re.match(r"^[^:\t\d][^:\t]{1,40}:\s*$", lines[i + 1]):
            val = lines[i + 1].strip()
        if val and label not in out:
            out[label] = val
    return out


def get_label(labels: dict, *pats):
    for k, v in labels.items():
        if any(re.fullmatch(p, k) for p in pats):
            return v
    return None


def parse_address(lines: list) -> dict:
    """Empfängeradresse: Name, [Zusatz], Straße, 'PLZ Ort' bzw. 'Ort, …, POSTCODE' bzw. 'Ort', [Land]."""
    lines = [re.sub(r"\s+", " ", t).strip() for t in lines if t and t.strip()]
    if len(lines) < 2:
        raise InvoiceError("Empfängeradresse nicht erkennbar (mindestens Name und Ort erwartet).")
    a = {"BT-44": lines[0]}
    rest = lines[1:]
    if rest and rest[-1].lower().strip(" .") in COUNTRIES:
        a["BT-55"] = COUNTRIES[rest.pop().lower().strip(" .")]
    if not rest:
        raise InvoiceError("Empfängeradresse: Ortszeile fehlt.")
    city_line, street = rest[-1], rest[:-1]
    if m := re.match(PLZ_RE, city_line):
        a["BT-53"], a["BT-52"] = m.group(1), m.group(2)
    elif m := re.match(r"^(.*?)[,\s]+([A-Z]{1,2}\d[A-Z\d]?\s*\d[A-Z]{2})$", city_line):
        a["BT-53"] = m.group(2)
        parts = [p.strip() for p in m.group(1).split(",") if p.strip()]
        a["BT-52"] = parts[0] if parts else city_line
        if len(parts) > 1:
            a["BT-54"] = ", ".join(parts[1:])
    elif m := re.match(r"^(.*?\D)\s+(\d{4,6})$", city_line):
        a["BT-52"], a["BT-53"] = m.group(1).strip(" ,"), m.group(2)
    else:
        a["BT-52"] = city_line
    if street:
        a["BT-50"] = street[-1]
        if len(street) > 1:
            a["BT-51"] = ", ".join(street[:-1])
    return a


def extract_from_docx(path: str) -> dict:
    """Liest die Rechnungsdaten aus einer Word-Rechnung (deutsch oder englisch)."""
    from docx import Document

    body, head = docx_paragraphs(path)
    body = [t for t in body if t]
    all_lines = body + [t for t in head if t]
    full = "\n".join(all_lines)
    headtxt = "\n".join(head)
    labels = find_labels(body)
    doc = Document(path)
    d: dict = {}
    warn: list = []

    def cells(row):
        return [c.text.strip() for c in row.cells]

    # --- Kopf: Nummer, Art, Datum -----------------------------------------------------------
    m = re.search(r"(Rechnung|Gutschrift|Korrekturrechnung|Invoice|Credit\s*Note)\s*"
                  r"(?:Nr\.?|-?Nummer|-Nr\.?|No\.?|Number|#)\s*[:.]?\s*([A-Za-z0-9][A-Za-z0-9\-/_.]*)", full, re.I)
    if not m:
        raise InvoiceError("Rechnungsnummer nicht gefunden (erwartet z. B. 'Rechnung Nr. …' oder 'Invoice No. …').")
    d["BT-1"] = m.group(2).rstrip(".")
    d["BT-3"] = 381 if re.match(r"gutschrift|credit", m.group(1), re.I) else 380
    inv_date = parse_date(get_label(labels, r"(rechnungsdatum|datum|invoice date|date)") or "")
    if not inv_date:
        for line in all_lines:
            if re.match(r"^[A-ZÄÖÜ][^,\d]{1,40},\s*(?:\d|[A-Za-z]+\s+\d)", line) and (inv_date := parse_date(line)):
                break
    if not inv_date:
        raise InvoiceError("Rechnungsdatum nicht gefunden (erwartet z. B. 'Hamburg, 10. November 2026' oder 'Date: 10.11.2026').")
    d["BT-2"] = inv_date

    # --- Verkäufer: Absenderzeile 'Name · Straße · PLZ Ort' + Briefkopf --------------------------
    seller: dict = {}
    sender = next((t for t in body if re.search(r"[·•|]", t) and re.search(r"\b\d{4,5}\b", t)), None)
    if not sender:
        raise InvoiceError("Absenderzeile ('Name · Straße · PLZ Ort') nicht gefunden.")
    parts = [p.strip() for p in re.split(r"\s*[·•|]\s*", sender) if p.strip()]
    name = re.sub(r"^(RA|RAin|Rechtsanwalt|Rechtsanwältin)\s+", "", parts[0])
    seller["BT-27"] = name
    if not re.search(r"\b(GmbH|AG|KG|OHG|UG|e\.\s?K\.|SE|mbB|PartG|GbR|Ltd|Limited|LLP|LLC|Inc|S\.A\.|B\.V\.)\b", name):
        seller["BT-41"] = name                     # Einzelperson: zugleich Ansprechpartner
    if len(parts) >= 3:
        seller["BT-35"] = parts[1]
        if mm := re.match(PLZ_RE, parts[2]):
            seller["BT-38"], seller["BT-37"] = mm.group(1), mm.group(2)
    seller["BT-40"] = "DE"
    if mm := re.search(r"(?:Telefon|Tel\.?|Phone|Fon)\s*:?\s*(\+?[\d ()/\-]{6,})", headtxt or full):
        seller["BT-42"] = re.sub(r"[^\d+]", "", mm.group(1))
    if mm := re.search(r"[\w.+\-]+@[\w\-]+(?:\.[\w\-]+)+", headtxt or full):
        seller["BT-34"] = seller["BT-43"] = mm.group(0)
    if mm := re.search(r"(?:Steuer-?\s*Nr\.?|Steuernummer|Tax\s*(?:ID|No\.?|Number))\s*:?\s*(\d[\d/ ]{5,19}\d)", full, re.I):
        seller["BT-32"] = mm.group(1).strip()
    for src in (headtxt, full):   # eigene USt-IdNr. zuerst im Briefkopf (im Text kann die des Kunden stehen)
        if (mm := re.search(r"(?:USt\.?-?\s*Id\.?-?\s*Nr\.?|VAT\s*(?:ID|No\.?|Reg\.?\s*No\.?)|UID)\s*:?\s*(DE\s?\d{9})", src, re.I)
                or re.search(r"\b(DE\d{9})\b", src)):
            seller["BT-31"] = mm.group(1).replace(" ", "")
            break
    d["BG-4"] = seller

    # --- Zahlung ---------------------------------------------------------------------------------
    pay: dict = {}
    if mm := re.search(r"IBAN\s*:?\s*([A-Z]{2}\d{2}[A-Z0-9 ]{10,32})", full):
        iban = re.sub(r"\s+", "", mm.group(1))
        iban = next((iban[:n] for n in range(len(iban), 14, -1) if iban_ok(iban[:n])), None)
        if not iban:
            raise InvoiceError("IBAN gefunden, aber die Prüfsumme ist ungültig.")
        pay["BT-81"], pay["BT-84"] = 58, iban
        if mm := re.search(r"(?:BIC|SWIFT)[^:\n]{0,12}:?\s*([A-Z]{6}[A-Z0-9]{2}(?:[A-Z0-9]{3})?)\b", full):
            pay["BT-86"] = mm.group(1)
        if owner := get_label(labels, r"(account name|kontoinhaber(in)?|account holder|empfänger)"):
            pay["BT-85"] = owner
    else:
        warn.append("keine IBAN gefunden -> in der Word-Vorlage ergänzen ('IBAN: …')")
    d["BG-16"] = pay

    # --- Käufer -----------------------------------------------------------------------------------
    addr_lines = None
    for t in doc.tables:                       # Adressfeld = Zelle unter der Absenderzeile
        rows_ = list(t.rows)
        for i, r in enumerate(rows_[:-1]):
            if any(re.sub(r"\s+", " ", c.text).strip() == sender for c in r.cells):
                addr_lines = rows_[i + 1].cells[0].text.split("\n")
                break
        if addr_lines:
            break
    if not addr_lines:                         # sonst: Absätze nach der Absenderzeile
        addr_lines = []
        for t in body[body.index(sender) + 1:][:10]:
            if re.match(r"^(Rechnung|Invoice|Gutschrift|Credit|Mein Zeichen|Our ref|[A-ZÄÖÜ][^,\d]{1,40},\s*\d)", t):
                break
            addr_lines.append(t)
    buyer = parse_address(addr_lines)
    # USt-IdNr. des Leistungsempfängers: beschriftete Nummer im Text, die nicht die eigene ist
    bodytxt = "\n".join(body)
    for mm in re.finditer(VAT_ID_RE, bodytxt):
        vid = re.sub(r"\s+", "", mm.group(1)).upper()
        if vid != seller.get("BT-31") and re.fullmatch(r"[A-Z]{2}[0-9A-Z]{8,12}", vid) and re.search(r"\d", vid):
            buyer["BT-48"] = vid
            break
    if "BT-48" not in buyer:   # z. B. 'Recipient VAT ID: 697 2532 99' (ohne Länderkürzel)
        mm = re.search(r"(?i:(?:recipient|customer|buyer|client|empfänger|leistungsempfänger|kunde)\w*[^:\n]{0,25}"
                       r"(?:VAT|USt|UID)[^:\n]{0,20}|(?:VAT|USt-?Id\.?-?Nr\.?|UID)[^:\n]{0,15}"
                       r"(?:recipient|customer|buyer|client|leistungsempfänger\w*|kunden?))\s*:\s*(\d[\d ]{6,14}\d)\b", bodytxt)
        if mm and buyer.get("BT-55"):
            prefix = {"GR": "EL"}.get(buyer["BT-55"], buyer["BT-55"])
            buyer["BT-48"] = prefix + re.sub(r"\s+", "", mm.group(1))
        elif mm:
            warn.append(f"USt-IdNr. des Kunden ohne Länderkürzel ('{mm.group(1)}') und ohne Landzeile -> ignoriert")
    if "BT-55" not in buyer:
        prefix = buyer.get("BT-48", "DE")[:2]
        buyer["BT-55"] = {"EL": "GR", "XI": "GB"}.get(prefix, prefix)
        if buyer["BT-55"] == "DE" and len(buyer.get("BT-53", "00000")) != 5:
            warn.append(f"Empfänger-PLZ '{buyer.get('BT-53')}' ist keine deutsche PLZ, aber keine Landzeile "
                        "gefunden -> Land unter die Ortszeile schreiben")
    d["BG-7"] = buyer

    # --- Projekt / Vertrag ---------------------------------------------------------------------------
    if v := get_label(labels, r"(projekt|project)"):
        d["BT-11"] = v
    if v := get_label(labels, r"(grundlage|basis|vertrag|contract|agreement|vereinbarung)"):
        d["BT-12"] = v

    # --- Positionstabelle --------------------------------------------------------------------------
    DESC = r"(position|bezeichnung|beschreibung|leistung|description|item|service)"
    AMT = r"(betrag|gesamt|summe|netto|amount|total|net)"
    table = hdr = None
    for t in doc.tables:
        for r in t.rows:
            c = [x.lower() for x in cells(r)]
            if any(re.match(DESC, x) for x in c) and any(re.match(AMT, x) for x in c):
                table, hdr = t, c
                break
        if table:
            break
    if not table:
        raise InvoiceError("Positionstabelle (Kopfzeile mit 'Position'/'Description' und 'Betrag'/'Amount') nicht gefunden.")

    def col(*pats):
        for i, x in enumerate(hdr):
            if any(re.match(p, x) for p in pats):
                return i
        return None

    c_desc = col(DESC)
    c_qty = col(r"menge|anzahl|stunden|std|tage|quantity|qty|hours|days|units")
    c_unit = col(r"einheit|unit$")
    c_price = col(r"einzelpreis|e-preis|preis|stundensatz|tagessatz|satz|tarif|honorar|unit price|price|rate|hourly")
    c_amt = next((i for i, x in enumerate(hdr) if re.match(AMT, x) and i not in (c_desc, c_price, c_qty)), None)
    rows = list(table.rows)
    start = next(i for i, r in enumerate(rows) if [x.lower() for x in cells(r)] == hdr) + 1
    TOTAL = (r"^\+?\s*(zwischensumme|summe|netto|gesamt|umsatzsteuer|ust|mwst|zzgl|rechnungsbetrag|endbetrag|"
             r"sub-?total|total|net\b|vat|value added tax|tax|grand total|amount due|balance due|plus)")
    positions, totals, rate, currencies = [], {}, None, set()
    in_totals = False
    for r in rows[start:]:
        c = cells(r)
        nonempty = [x for x in c if x]
        if not nonempty:
            continue
        currencies |= {detect_currency(x) for x in nonempty if detect_currency(x)}
        label = re.sub(r"\s+", " ", nonempty[0])
        if in_totals or re.match(TOTAL, label, re.I):
            in_totals = True
            amount = parse_amount(nonempty[-1]) if len(nonempty) > 1 and re.search(r"\d", nonempty[-1]) else Decimal("0")
            if re.search(r"umsatzsteuer|\bust\b|mwst|\bvat\b|value added tax|\btax\b", label, re.I):
                totals["vat"] = amount
                if mm := re.search(r"(\d{1,2}(?:[,.]\d+)?)\s*%", label):
                    rate = parse_amount(mm.group(1))
            elif re.search(r"zwischensumme|netto|sub-?total|\bnet\b|^summe$", label, re.I):
                totals["net"] = amount
            else:
                totals["gross"] = amount
            continue
        amount = parse_amount(c[c_amt]) if c_amt is not None and c[c_amt] else parse_amount(nonempty[-1])
        qty = parse_amount(c[c_qty]) if c_qty is not None and re.search(r"\d", c[c_qty]) else Decimal("1")
        unit = "C62"
        if c_qty is not None:
            for word in re.findall(r"[A-Za-zäöüÄÖÜß.]+", hdr[c_qty] + " " + c[c_qty]):
                if word.lower() in UNITS:
                    unit = UNITS[word.lower()]
        if c_unit is not None and c[c_unit]:
            unit = UNITS.get(c[c_unit].lower(), unit)
        if c_price is not None and c[c_price]:
            price = parse_amount(c[c_price])
        else:
            price = (amount / qty).quantize(Decimal("0.0001"), rounding=ROUND_HALF_UP)
        if (qty * price).quantize(CENT, rounding=ROUND_HALF_UP) != amount:
            # Menge im Dokument gerundet (z. B. 0,87 Std. statt 52 Min.): Menge aus Betrag/Preis herleiten
            exact = amount / price if price else None
            half = Decimal(5) / Decimal(10) ** (len(c[c_qty].split(",")[1].split()[0]) + 1) \
                if c_qty is not None and "," in c[c_qty] else Decimal("0.005")
            fixed = None
            if exact is not None and abs(exact - qty) <= half:
                for places in range(2, 9):
                    cand = exact.quantize(Decimal(1).scaleb(-places), rounding=ROUND_HALF_UP)
                    if (cand * price).quantize(CENT, rounding=ROUND_HALF_UP) == amount:
                        fixed = cand
                        break
            if fixed is None:
                raise InvoiceError(f"Position '{label}': Menge x Einzelpreis ({qty} x {price}) ergibt nicht {amount}.")
            warn.append(f"Position '{label}': Menge im Dokument gerundet ({qty}); im XML wird {fixed} verwendet, "
                        f"damit Menge x Preis = {amount} ergibt.")
            qty = fixed
        desc = re.sub(r"\s+", " ", c[c_desc] if c_desc is not None else label).strip()
        positions.append({"desc": desc, "qty": qty, "unit": unit, "price": price, "net": amount,
                          "period": parse_period(desc, inv_date.year)})
    if not positions:
        raise InvoiceError("Keine Positionszeilen in der Tabelle gefunden.")
    if len(currencies) > 1:
        raise InvoiceError(f"Mehrere Währungen in der Positionstabelle: {', '.join(sorted(currencies))}.")
    d["BT-5"] = currencies.pop() if currencies else "EUR"

    # --- Steuerkategorie: Befreiungs-/Reverse-Charge-Vermerk im Text (ohne Briefkopf) -----------------
    body_lines = [re.sub(r"\s+", " ", t) for t in body] + [re.sub(r"\s+", " ", " ".join(cells(r))) for r in rows]
    category, reason, reason_code = "S", None, None
    for cat, pat, code in TAX_EXEMPTIONS:
        hit = next((t for t in body_lines if re.search(pat, t, re.I)), None)
        if hit:
            sentence = next((x for x in re.split(r"(?<=[.!?])\s+", hit) if re.search(pat, x, re.I)), hit)
            category, reason, reason_code = cat, sentence.strip(), code
            break
    # Nicht steuerbar (Kategorie O): ausdrücklicher Vermerk oder Reverse Charge an Kunden außerhalb der EU
    # (Leistungsort im Drittland, § 3a Abs. 2 UStG).
    if category == "S" and (hit := next((t for t in body_lines if re.search(NOT_TAXABLE, t, re.I)), None)):
        sentence = next((x for x in re.split(r"(?<=[.!?])\s+", hit) if re.search(NOT_TAXABLE, x, re.I)), hit)
        category, reason = "O", sentence.strip()
    if category == "AE" and buyer["BT-55"] not in EU_COUNTRIES:
        category = "O"
        warn.append(f"Reverse-Charge-Vermerk bei Kunde außerhalb der EU ({buyer['BT-55']}) -> als 'nicht steuerbar' "
                    "(Kategorie O) ausgegeben; USt-IdNr. von Verkäufer und Kunde entfallen im XML, "
                    "die Steuernummer dient als Verkäuferkennung.")
    if category == "O":
        reason_code = "VATEX-EU-O"
    if category == "S":
        if rate is None or rate == 0:
            raise InvoiceError("Umsatzsteuersatz nicht erkennbar bzw. 0 %. Erwartet wird eine Summenzeile wie "
                     "'+ Umsatzsteuer (19%)' bzw. 'VAT (19%)'. Bei steuerfreien Rechnungen bzw. Reverse Charge "
                     "muss der Grund im Rechnungstext stehen (z. B. 'Steuerschuldnerschaft des Leistungsempfängers' "
                     "oder 'Reverse charge').")
    else:
        if (rate or 0) != 0 or totals.get("vat", Decimal("0")) != 0:
            raise InvoiceError(f"Widerspruch: Rechnungstext nennt eine Steuerbefreiung/Reverse Charge ('{reason}'), "
                     f"die Tabelle weist aber Umsatzsteuer aus.")
        rate = Decimal("0")
        d["BG-23"] = [{"BT-118": category, "BT-119": 0, "BT-120": reason[:300],
                       **({"BT-121": reason_code} if reason_code else {})}]
    d["BG-25"] = []
    for i, p in enumerate(positions, 1):
        line = {"BT-126": i, "BT-153": p["desc"], "BT-131": p["net"], "BT-129": float(p["qty"]) if p["qty"] % 1 else int(p["qty"]),
                "BT-130": p["unit"], "BT-146": float(p["price"].quantize(Decimal("0.0001"))),
                "BT-151": category, "BT-152": float(rate) if rate % 1 else int(rate)}
        if p["period"]:
            line["BT-134"], line["BT-135"] = p["period"]
        d["BG-25"].append(line)

    # Summenabgleich: Dokument gegen eigene Berechnung
    net = sum(p["net"] for p in positions)
    vat = (net * rate / 100).quantize(CENT, rounding=ROUND_HALF_UP)
    for key, calc in (("net", net), ("vat", vat), ("gross", net + vat)):
        if key in totals and totals[key] != calc:
            raise InvoiceError(f"Summenprüfung fehlgeschlagen: Dokument {key} = {totals[key]}, berechnet = {calc}.")

    # --- Leistungszeitraum: Angabe 'Leistungszeitraum:'/'Service Period:', sonst aus den Positionen ------
    period = parse_period(get_label(labels, r"(leistungszeitraum|abrechnungszeitraum|zeitraum|service period|"
                                          r"period of service|billing period|period)") or "", inv_date.year)
    if not period:
        lp = [p["period"] for p in positions if p["period"]]
        if lp:
            period = (min(a for a, _ in lp), max(b for _, b in lp))
    if period:
        d["BG-14"] = {"BT-73": period[0], "BT-74": period[1]}

    # --- Zahlungsbedingungen -------------------------------------------------------------------------
    term = next((t for t in body if re.search(r"zahlbar|fällig|zahlungsziel|payable|payment due|due date|"
                                              r"payment terms|due upon|due on", t, re.I)), None)
    if term:
        sentence = next((x for x in re.split(r"(?<=[.!?])\s+", term)
                         if re.search(r"zahlbar|fällig|zahlungsziel|payable|payment|due", x, re.I)), term)
        txt = re.sub(r"\s+(?:auf das|to)\b.*$", "", sentence).strip().rstrip(".: ")
        d["BT-20"] = txt
        if (mm := re.search(r"(?:innerhalb von|binnen|within)\s*(\d+)\s*(?:Tagen|days)|(\d+)\s*(?:Tage|days)\s*(?:netto|net)",
                            term, re.I)):
            d["BT-9"] = inv_date + timedelta(days=int(mm.group(1) or mm.group(2)))
        elif re.search(r"mit rechnungserhalt|bei erhalt|nach erhalt|sofort|upon receipt|on receipt|immediately",
                       term, re.I):
            d["BT-9"] = inv_date
        elif dd := parse_date(term):
            d["BT-9"] = dd
    else:
        warn.append("keine Zahlungsbedingung gefunden -> in der Word-Vorlage ergänzen "
                    "(z. B. 'Zahlbar innerhalb von 30 Tagen' / 'Payable within 30 days')")
    for w in warn:
        print("Hinweis:", w)
    return d


def docx_fonts(docx_path: str) -> set:
    """Schriften, die das Dokument tatsächlich verwendet (direkte Formatierung + verwendete Formatvorlagen)."""
    with zipfile.ZipFile(docx_path) as z:
        parts = "".join(z.read(n).decode("utf-8", "ignore") for n in z.namelist()
                        if re.match(r"word/(document|header\d*|footer\d*)\.xml$", n))
        styles = z.read("word/styles.xml").decode("utf-8", "ignore") if "word/styles.xml" in z.namelist() else ""
        theme = "".join(z.read(n).decode("utf-8", "ignore") for n in z.namelist() if n.startswith("word/theme/"))
    fonts = set(re.findall(r'w:ascii="([^"]+)"', parts))
    used = set(re.findall(r'<w:(?:pStyle|rStyle|tblStyle) w:val="([^"]+)"', parts))
    style_map = {m.group(1): m.group(0) for m in
                 re.finditer(r'<w:style [^>]*w:styleId="([^"]+)".*?</w:style>', styles, re.S)}
    default = re.search(r'<w:style [^>]*w:type="paragraph"[^>]*w:default="1"[^>]*w:styleId="([^"]+)"', styles)
    if default:
        used.add(default.group(1))
    seen = set()
    while used - seen:  # Formatvorlagen inkl. basedOn-Kette
        sid = (used - seen).pop()
        seen.add(sid)
        body = style_map.get(sid, "")
        fonts |= set(re.findall(r'w:ascii="([^"]+)"', body))
        used |= set(re.findall(r'<w:basedOn w:val="([^"]+)"', body))
    if not fonts:  # keine explizite Schrift -> Design-Schriftart (Theme) gilt
        m = re.search(r'<a:minorFont><a:latin typeface="([^"]+)"', theme)
        if m:
            fonts.add(m.group(1))
    return {f for f in fonts if f}
