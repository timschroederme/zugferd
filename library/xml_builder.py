"""ZUGFeRD-/Factur-X-XML (CII, Profil EN 16931) aus den Rechnungsdaten erzeugen und prüfen."""

import re
from decimal import Decimal, ROUND_HALF_UP

from lxml import etree

from .errors import InvoiceError
from .values import CENT, fmt_amount, fmt_num, to_decimal, to_yyyymmdd


NS = {
    "rsm": "urn:un:unece:uncefact:data:standard:CrossIndustryInvoice:100",
    "ram": "urn:un:unece:uncefact:data:standard:ReusableAggregateBusinessInformationEntity:100",
    "udt": "urn:un:unece:uncefact:data:standard:UnqualifiedDataType:100",
}


def add_el(parent, tag, text=None, **attrs):
    """Kindelement 'prefix:Name' anlegen (Reihenfolge = Aufrufreihenfolge -> XSD-Reihenfolge beachten!)."""
    prefix, name = tag.split(":")
    e = etree.SubElement(parent, "{%s}%s" % (NS[prefix], name))
    if text is not None:
        e.text = str(text)
    for k, v in attrs.items():
        e.set(k, str(v))
    return e


def add_date_el(parent, tag, v):
    outer = add_el(parent, tag)
    add_el(outer, "udt:DateTimeString", to_yyyymmdd(v), format="102")


def split_scheme(v, default_scheme=None):
    """Skalar oder {value:, schemeID:} -> (wert, schema)."""
    if isinstance(v, dict):
        return str(v.get("value", "")), v.get("schemeID", default_scheme)
    return str(v), default_scheme


def add_party(parent, tag, p: dict, role: str, err: list):
    """Verkäufer (role='seller', BT-27ff.) oder Käufer (role='buyer', BT-44ff.)."""
    s = role == "seller"
    n = (lambda a, b: a if s else b)  # BT-Nummer je nach Rolle
    NAME, ID, LEGAL, VAT = n("BT-27", "BT-44"), n("BT-29", "BT-46"), n("BT-30", "BT-47"), n("BT-31", "BT-48")
    ADDR = [n(f"BT-{i}", f"BT-{i + 15}") for i in (35, 36, 37, 38, 39, 40)]  # Straße, Zusatz, Ort, PLZ, Land..
    # Seller: BT-35 Straße, BT-36 Zusatz, BT-37 Ort, BT-38 PLZ, BT-39 Bundesland, BT-40 Land
    # Buyer:  BT-50 Straße, BT-51 Zusatz, BT-52 Ort, BT-53 PLZ, BT-54 Bundesland, BT-55 Land
    street, street2, city, plz, subdiv, country = ADDR
    ELEC = n("BT-34", "BT-49")
    CNAME, CPHONE, CMAIL = n(("BT-41", "BT-42", "BT-43"), ("BT-56", "BT-57", "BT-58"))

    t = add_el(parent, tag)
    if p.get(ID) is not None:
        v, sch = split_scheme(p[ID])
        if sch:
            add_el(t, "ram:GlobalID", v, schemeID=sch)
        else:
            add_el(t, "ram:ID", v)
    if not p.get(NAME):
        err.append(f"{NAME} (Name) fehlt")
    add_el(t, "ram:Name", p.get(NAME, ""))
    trading = p.get("BT-28") if s else None
    if p.get(LEGAL) is not None or trading:
        lo = add_el(t, "ram:SpecifiedLegalOrganization")
        if p.get(LEGAL) is not None:
            v, sch = split_scheme(p[LEGAL])
            add_el(lo, "ram:ID", v, **({"schemeID": sch} if sch else {}))
        if trading:
            add_el(lo, "ram:TradingBusinessName", trading)
    if any(p.get(k) for k in (CNAME, CPHONE, CMAIL)):
        c = add_el(t, "ram:DefinedTradeContact")
        if p.get(CNAME):
            add_el(c, "ram:PersonName", p[CNAME])
        if p.get(CPHONE):
            add_el(add_el(c, "ram:TelephoneUniversalCommunication"), "ram:CompleteNumber", p[CPHONE])
        if p.get(CMAIL):
            add_el(add_el(c, "ram:EmailURIUniversalCommunication"), "ram:URIID", p[CMAIL])
    for k in (city, country):
        if not p.get(k):
            err.append(f"{k} (Adresse: Ort/Land) fehlt")
    a = add_el(t, "ram:PostalTradeAddress")
    if p.get(plz):
        add_el(a, "ram:PostcodeCode", p[plz])
    if p.get(street):
        add_el(a, "ram:LineOne", p[street])
    if p.get(street2):
        add_el(a, "ram:LineTwo", p[street2])
    add_el(a, "ram:CityName", p.get(city, ""))
    add_el(a, "ram:CountryID", p.get(country, ""))
    if p.get(subdiv):
        add_el(a, "ram:CountrySubDivisionName", p[subdiv])
    if p.get(ELEC) is not None:
        v, sch = split_scheme(p[ELEC], "EM")
        add_el(add_el(t, "ram:URIUniversalCommunication"), "ram:URIID", v, schemeID=sch)
    if p.get(VAT):
        add_el(add_el(t, "ram:SpecifiedTaxRegistration"), "ram:ID", p[VAT], schemeID="VA")
    if s and p.get("BT-32"):
        add_el(add_el(t, "ram:SpecifiedTaxRegistration"), "ram:ID", p["BT-32"], schemeID="FC")


def build_cii_xml(data: dict) -> bytes:
    """Erzeugt aus den Rechnungsdaten (BT/BG-Dictionary) das CII-XML (EN 16931). Summen/MwSt. werden berechnet."""
    err: list = []
    seller, buyer = data.get("BG-4") or {}, data.get("BG-7") or {}
    pay, lines = data.get("BG-16") or {}, data.get("BG-25") or []
    period = data.get("BG-14") or {}
    reasons = {(r.get("BT-118"), to_decimal(r.get("BT-119", 0))): r for r in (data.get("BG-23") or [])}

    for k in ("BT-1", "BT-2"):
        if not data.get(k):
            err.append(f"{k} fehlt")
    if not lines:
        err.append("BG-25 (Rechnungspositionen) fehlt")
    if not pay.get("BT-81"):
        err.append("BG-16/BT-81 (Zahlungsart-Code, z. B. 58 = SEPA-Überweisung) fehlt")
    if str(pay.get("BT-81")) in ("30", "58") and not pay.get("BT-84"):
        err.append("BG-16/BT-84 (IBAN) fehlt, obwohl Überweisung gewählt")
    if not (data.get("BT-9") or data.get("BT-20")):
        err.append("BT-9 (Fälligkeit) oder BT-20 (Zahlungsbedingungen) erforderlich")

    # --- Positionen rechnen -------------------------------------------------
    groups: dict = {}   # (Kategorie, Satz) -> Bemessungsgrundlage
    line_total = Decimal("0")
    calc = []
    for i, l in enumerate(lines, 1):
        for k in ("BT-129", "BT-130", "BT-146", "BT-153", "BT-151"):
            if l.get(k) is None:
                err.append(f"Position {i}: {k} fehlt")
        if err:
            continue
        qty, price = to_decimal(l["BT-129"]), to_decimal(l["BT-146"])
        net = to_decimal(l["BT-131"]) if l.get("BT-131") is not None else qty * price
        net = net.quantize(CENT, rounding=ROUND_HALF_UP)
        cat = str(l["BT-151"])
        rate = to_decimal(l.get("BT-152", 0)) if cat != "O" else Decimal("0")
        groups[(cat, rate)] = groups.get((cat, rate), Decimal("0")) + net
        line_total += net
        calc.append((i, l, qty, price, net, cat, rate))
    if err:
        raise InvoiceError("Fehlende/ungültige Rechnungsangaben (keine PDF erzeugt):\n  - " + "\n  - ".join(err))

    tax_total = Decimal("0")
    for (cat, rate), basis in groups.items():
        tax_total += (basis * rate / 100).quantize(CENT, rounding=ROUND_HALF_UP)
        if cat in ("E", "AE", "K", "G", "O"):
            r = reasons.get((cat, rate), {})
            if not (r.get("BT-120") or r.get("BT-121")):
                err.append(f"BG-23: für Kategorie {cat} ({rate} %) fehlt BT-120 (Text) oder BT-121 (Code)")
    if any(c in ("AE", "K") for c, _ in groups) and not (buyer.get("BT-48") or buyer.get("BT-47")):
        err.append("Bei Reverse Charge bzw. innergemeinschaftlicher Lieferung muss die USt-IdNr. des "
                   "Leistungsempfängers (BG-7/BT-48) in der Rechnung stehen, z. B. "
                   "'USt-IdNr. des Leistungsempfängers: ATU12345678' (Regeln BR-AE-02 / BR-IC-02)")
    # USt-IdNr. des Rechnungsstellers ist Pflicht (Regel BR-CO-26: Verkäufer muss maschinell
    # identifizierbar sein; die Steuernummer BT-32 genügt dafür nicht).
    vat_id = str(seller.get("BT-31") or "").replace(" ", "").upper()
    if not vat_id:
        err.append("USt-IdNr. des Rechnungsstellers (BG-4/BT-31) fehlt. Bitte im Briefkopf der Word-Vorlage "
                   "angeben ('USt-IdNr.: DE…') "
                   "(Regel BR-CO-26)")
    elif not re.fullmatch(r"[A-Z]{2}[0-9A-Z+*.]{2,13}", vat_id) or \
            (vat_id.startswith("DE") and not re.fullmatch(r"DE\d{9}", vat_id)):
        err.append(f"USt-IdNr. des Rechnungsstellers (BG-4/BT-31) hat ein ungültiges Format: {seller['BT-31']!r} "
                   "(erwartet z. B. DE123456789)")
    else:
        seller["BT-31"] = vat_id
    if any(c == "O" for c, _ in groups):
        # BR-O-02: bei 'nicht steuerbar' keine USt-IdNr. von Verkäufer/Käufer im XML.
        # BR-CO-26 verlangt dann eine andere Verkäuferkennung (BT-29): die Steuernummer.
        seller, buyer = dict(seller), dict(buyer)
        seller.pop("BT-31", None)
        buyer.pop("BT-48", None)
        if not seller.get("BT-29"):
            if seller.get("BT-32"):
                seller["BT-29"] = seller["BT-32"]
            else:
                err.append("Bei 'nicht steuerbar' (Kategorie O) wird die Steuernummer des Rechnungsstellers "
                           "als Verkäuferkennung benötigt ('Steuer-Nr.: …' / 'Tax ID: …' im Briefkopf)")

    root = etree.Element("{%s}CrossIndustryInvoice" % NS["rsm"], nsmap=NS)
    ctx = add_el(root, "rsm:ExchangedDocumentContext")
    add_el(add_el(ctx, "ram:GuidelineSpecifiedDocumentContextParameter"), "ram:ID", "urn:cen.eu:en16931:2017")

    doc = add_el(root, "rsm:ExchangedDocument")
    add_el(doc, "ram:ID", data.get("BT-1", ""))
    add_el(doc, "ram:TypeCode", data.get("BT-3", 380))
    add_date_el(doc, "ram:IssueDateTime", data.get("BT-2"))
    notes = data.get("BT-22")
    for note in ([notes] if isinstance(notes, str) else (notes or [])):
        add_el(add_el(doc, "ram:IncludedNote"), "ram:Content", note)

    trade = add_el(root, "rsm:SupplyChainTradeTransaction")
    for i, l, qty, price, net, cat, rate in calc:
        li = add_el(trade, "ram:IncludedSupplyChainTradeLineItem")
        ald = add_el(li, "ram:AssociatedDocumentLineDocument")
        add_el(ald, "ram:LineID", l.get("BT-126", i))
        if l.get("BT-127"):
            add_el(add_el(ald, "ram:IncludedNote"), "ram:Content", l["BT-127"])
        prod = add_el(li, "ram:SpecifiedTradeProduct")
        if l.get("BT-155"):
            add_el(prod, "ram:SellerAssignedID", l["BT-155"])
        add_el(prod, "ram:Name", l["BT-153"])
        if l.get("BT-154"):
            add_el(prod, "ram:Description", l["BT-154"])
        add_el(add_el(add_el(li, "ram:SpecifiedLineTradeAgreement"), "ram:NetPriceProductTradePrice"),
            "ram:ChargeAmount", fmt_num(price))
        add_el(add_el(li, "ram:SpecifiedLineTradeDelivery"), "ram:BilledQuantity", fmt_num(qty), unitCode=l["BT-130"])
        st = add_el(li, "ram:SpecifiedLineTradeSettlement")
        tx = add_el(st, "ram:ApplicableTradeTax")
        add_el(tx, "ram:TypeCode", "VAT")
        add_el(tx, "ram:CategoryCode", cat)
        if cat != "O":
            add_el(tx, "ram:RateApplicablePercent", fmt_amount(rate))
        if l.get("BT-134") or l.get("BT-135"):
            bp = add_el(st, "ram:BillingSpecifiedPeriod")
            if l.get("BT-134"):
                add_date_el(bp, "ram:StartDateTime", l["BT-134"])
            if l.get("BT-135"):
                add_date_el(bp, "ram:EndDateTime", l["BT-135"])
        add_el(add_el(st, "ram:SpecifiedTradeSettlementLineMonetarySummation"), "ram:LineTotalAmount", fmt_amount(net))

    agr = add_el(trade, "ram:ApplicableHeaderTradeAgreement")
    if data.get("BT-10"):
        add_el(agr, "ram:BuyerReference", data["BT-10"])
    add_party(agr, "ram:SellerTradeParty", seller, "seller", err)
    add_party(agr, "ram:BuyerTradeParty", buyer, "buyer", err)
    if data.get("BT-13"):
        add_el(add_el(agr, "ram:BuyerOrderReferencedDocument"), "ram:IssuerAssignedID", data["BT-13"])
    if data.get("BT-12"):
        add_el(add_el(agr, "ram:ContractReferencedDocument"), "ram:IssuerAssignedID", data["BT-12"])
    if data.get("BT-11"):
        pj = add_el(agr, "ram:SpecifiedProcuringProject")
        add_el(pj, "ram:ID", data["BT-11"])
        add_el(pj, "ram:Name", data.get("BT-11_Name", data["BT-11"]))
    if err:
        raise InvoiceError("Fehlende/ungültige Rechnungsangaben (keine PDF erzeugt):\n  - " + "\n  - ".join(err))

    dlv = add_el(trade, "ram:ApplicableHeaderTradeDelivery")
    # Leistungsdatum: BT-72, sonst Ende des Leistungszeitraums (BT-74). Das Element darf nicht leer
    # bleiben (Peppol-Regel R008), deshalb nur anlegen, wenn ein Datum vorhanden ist.
    delivery = data.get("BT-72")
    if not delivery and period.get("BT-74") and to_yyyymmdd(period["BT-74"]) <= to_yyyymmdd(data.get("BT-2")):
        delivery = period["BT-74"]   # nur abgelaufene Zeiträume; zukünftiges Leistungsdatum wäre irreführend
    if delivery:
        add_date_el(add_el(dlv, "ram:ActualDeliverySupplyChainEvent"), "ram:OccurrenceDateTime", delivery)

    st = add_el(trade, "ram:ApplicableHeaderTradeSettlement")
    if data.get("BT-83"):
        add_el(st, "ram:PaymentReference", data["BT-83"])
    add_el(st, "ram:InvoiceCurrencyCode", data.get("BT-5", "EUR"))
    pm = add_el(st, "ram:SpecifiedTradeSettlementPaymentMeans")
    add_el(pm, "ram:TypeCode", pay["BT-81"])
    if pay.get("BT-82"):
        add_el(pm, "ram:Information", pay["BT-82"])
    if pay.get("BT-84"):
        acc = add_el(pm, "ram:PayeePartyCreditorFinancialAccount")
        add_el(acc, "ram:IBANID", str(pay["BT-84"]).replace(" ", ""))
        if pay.get("BT-85"):
            add_el(acc, "ram:AccountName", pay["BT-85"])
    if pay.get("BT-86"):
        add_el(add_el(pm, "ram:PayeeSpecifiedCreditorFinancialInstitution"), "ram:BICID", pay["BT-86"])
    for (cat, rate), basis in groups.items():
        tx = add_el(st, "ram:ApplicableTradeTax")
        add_el(tx, "ram:CalculatedAmount", fmt_amount(basis * rate / 100))
        add_el(tx, "ram:TypeCode", "VAT")
        r = reasons.get((cat, rate), {})
        if r.get("BT-120"):
            add_el(tx, "ram:ExemptionReason", r["BT-120"])
        add_el(tx, "ram:BasisAmount", fmt_amount(basis))
        add_el(tx, "ram:CategoryCode", cat)
        if r.get("BT-121"):
            add_el(tx, "ram:ExemptionReasonCode", r["BT-121"])
        if cat != "O":
            add_el(tx, "ram:RateApplicablePercent", fmt_amount(rate))
    if period.get("BT-73") or period.get("BT-74"):
        bp = add_el(st, "ram:BillingSpecifiedPeriod")
        if period.get("BT-73"):
            add_date_el(bp, "ram:StartDateTime", period["BT-73"])
        if period.get("BT-74"):
            add_date_el(bp, "ram:EndDateTime", period["BT-74"])
    if data.get("BT-9") or data.get("BT-20"):
        pt = add_el(st, "ram:SpecifiedTradePaymentTerms")
        if data.get("BT-20"):
            add_el(pt, "ram:Description", data["BT-20"])
        if data.get("BT-9"):
            add_date_el(pt, "ram:DueDateDateTime", data["BT-9"])
    grand = line_total + tax_total
    prepaid = to_decimal(data.get("BT-113", 0))
    sm = add_el(st, "ram:SpecifiedTradeSettlementHeaderMonetarySummation")
    add_el(sm, "ram:LineTotalAmount", fmt_amount(line_total))
    add_el(sm, "ram:TaxBasisTotalAmount", fmt_amount(line_total))
    add_el(sm, "ram:TaxTotalAmount", fmt_amount(tax_total), currencyID=data.get("BT-5", "EUR"))
    add_el(sm, "ram:GrandTotalAmount", fmt_amount(grand))
    if prepaid:
        add_el(sm, "ram:TotalPrepaidAmount", fmt_amount(prepaid))
    add_el(sm, "ram:DuePayableAmount", fmt_amount(grand - prepaid))

    # berechnete Werte für die Übersicht (show_invoice) festhalten
    data["BG-22"] = {"BT-106": line_total, "BT-109": line_total, "BT-110": tax_total, "BT-112": grand,
                     **({"BT-113": prepaid} if prepaid else {}), "BT-115": grand - prepaid}
    data["_BG-23"] = [{"BT-118": cat, "BT-119": rate, "BT-116": basis,
                       "BT-117": (basis * rate / 100).quantize(CENT, rounding=ROUND_HALF_UP),
                       **{k: v for k, v in reasons.get((cat, rate), {}).items() if k in ("BT-120", "BT-121")}}
                      for (cat, rate), basis in groups.items()]
    data["_O"] = any(c == "O" for c, _ in groups)
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", pretty_print=True)


def check_xsd(xml_bytes: bytes) -> None:
    """XSD-Prüfung, falls die Bibliothek `factur-x` installiert ist (pip install factur-x)."""
    try:
        from facturx import xml_check_xsd
    except ImportError:
        print("Hinweis: `pip install factur-x` aktiviert zusätzlich die XSD-Prüfung des XML.")
        return
    xml_check_xsd(xml_bytes, flavor="factur-x", level="en16931")
    print("XML besteht die XSD-Prüfung (EN 16931).")
