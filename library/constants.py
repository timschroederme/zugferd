"""Konstanten: Monatsnamen, Länder, Einheiten, Währungen, Muster für USt-IdNr. und Steuervermerke."""


MONTHS = {m: i for i, m in enumerate(
    ["januar", "februar", "märz", "april", "mai", "juni", "juli", "august",
     "september", "oktober", "november", "dezember"], 1)}


MONTHS_EN = {m: i for i, m in enumerate(
    ["january", "february", "march", "april", "may", "june", "july", "august",
     "september", "october", "november", "december"], 1)}


MONTH_NAMES = {**MONTHS, **MONTHS_EN, "jan": 1, "feb": 2, "mar": 3, "apr": 4, "jun": 6, "jul": 7, "aug": 8,
               "sep": 9, "sept": 9, "oct": 10, "okt": 10, "nov": 11, "dec": 12, "dez": 12, "maerz": 3}


CURRENCIES = {"€": "EUR", "£": "GBP", "$": "USD", "CHF": "CHF", "EUR": "EUR", "GBP": "GBP", "USD": "USD",
              "SEK": "SEK", "DKK": "DKK", "NOK": "NOK", "PLN": "PLN", "CZK": "CZK", "JPY": "JPY", "TRY": "TRY"}


COUNTRIES = {"deutschland": "DE", "germany": "DE", "türkiye": "TR", "turkey": "TR", "türkei": "TR",
             "united kingdom": "GB", "uk": "GB", "great britain": "GB", "england": "GB", "scotland": "GB",
             "wales": "GB", "northern ireland": "GB", "netherlands": "NL", "the netherlands": "NL",
             "france": "FR", "italy": "IT", "spain": "ES", "poland": "PL", "belgium": "BE", "denmark": "DK",
             "sweden": "SE", "norway": "NO", "norwegen": "NO", "ireland": "IE", "united states": "US",
             "united states of america": "US", "canada": "CA", "kanada": "CA", "japan": "JP", "china": "CN", "österreich": "AT", "austria": "AT", "schweiz": "CH",
             "switzerland": "CH", "liechtenstein": "LI", "tschechien": "CZ", "schweden": "SE",
             "finnland": "FI", "irland": "IE", "portugal": "PT", "griechenland": "GR", "ungarn": "HU",
             "slowakei": "SK", "slowenien": "SI", "kroatien": "HR", "rumänien": "RO", "bulgarien": "BG",
             "estland": "EE", "lettland": "LV", "litauen": "LT", "malta": "MT", "zypern": "CY",
             "vereinigtes königreich": "GB", "großbritannien": "GB", "usa": "US", "vereinigte staaten": "US",
             "niederlande": "NL", "frankreich": "FR", "italien": "IT", "spanien": "ES",
             "polen": "PL", "belgien": "BE", "luxemburg": "LU", "dänemark": "DK"}


UNITS = {"hour": "HUR", "hours": "HUR", "hrs": "HUR", "hr": "HUR", "day": "DAY", "days": "DAY",
         "piece": "C62", "pieces": "C62", "pcs": "C62", "month": "MON", "months": "MON", "flat": "LS",
         "std": "HUR", "std.": "HUR", "stunde": "HUR", "stunden": "HUR", "h": "HUR",
         "stk": "C62", "stk.": "C62", "stück": "C62", "tag": "DAY", "tage": "DAY", "tagen": "DAY",
         "monat": "MON", "monate": "MON", "pauschal": "LS", "psch.": "LS", "km": "KMT"}


W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


PLZ_RE = r"^(?:[A-Z]{1,2}-)?(\d{4,5}(?:\s?[A-Z]{2})?)\s+(\S.*)$"   # 22609 Hamburg, 1010 Wien, 1012 AB Amsterdam


VAT_ID_RE = (r"(?i:USt\.?-?\s*Id(?:ent)?\.?-?\s*Nr\.?|UID(?:-Nr\.?)?|VAT\s*(?:ID|No\.?|Reg(?:istration)?\.?\s*No\.?|number))"
             r"[^:\n]{0,40}?:?\s*([A-Z]{2}\s?[0-9A-Z](?:\s?[0-9A-Z]){7,12})(?![0-9A-Za-z])")


# Steuerbefreiungen / Reverse Charge, erkannt am Rechnungstext.  Reihenfolge = Priorität.
# (Kategorie, Muster, VATEX-Code für BT-121)
TAX_EXEMPTIONS = [
    ("AE", r"steuerschuldnerschaft des leistungsempfängers|reverse[\s-]*charge|§\s*13\s*b\s*ustg|"
           r"umkehr(?:ung)? der steuerschuld|steuerschuld(?:ner)? (?:ist|geht auf) de[nr] leistungsempfänger|"
           r"customer is (?:required|liable) to account for the vat|vat to be accounted for by the (?:customer|recipient)|"
           r"article 196|art\.?\s*196", "VATEX-EU-AE"),
    ("K",  r"intra-community supply|intra-community delivery|innergemeinschaftliche (?:steuerfreie )?lieferung|§\s*6\s*a\s*ustg|§\s*4\s*nr\.?\s*1\s*b\s*ustg", "VATEX-EU-IC"),
    ("G",  r"export delivery|exempt export|ausfuhrlieferung|steuerfreie ausfuhr|§\s*6\s+ustg|§\s*4\s*nr\.?\s*1\s*a\s*ustg", "VATEX-EU-G"),
    ("E",  r"vat exempt|exempt from vat|small business exemption|kleinunternehmer|§\s*19\s*(?:abs\.?\s*1\s*)?ustg|umsatzsteuerbefreit|von der umsatzsteuer befreit|"
           r"steuerfrei(?:e)? (?:leistung|umsatz)|steuerfrei (?:gemäß|nach|gem\.)|§\s*4\s*nr\.?\s*\d+\s*\w?\s*ustg", None),
]


EU_COUNTRIES = {"AT", "BE", "BG", "CY", "CZ", "DE", "DK", "EE", "ES", "FI", "FR", "GR", "HR", "HU", "IE", "IT",
                "LT", "LU", "LV", "MT", "NL", "PL", "PT", "RO", "SE", "SI", "SK"}


NOT_TAXABLE = r"not subject to (?:german )?vat|outside the scope of (?:german )?vat|nicht (?:im inland )?steuerbar|nicht der (?:deutschen )?umsatzsteuer unterliegend"
