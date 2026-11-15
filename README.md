# zugferd – ZUGFeRD-Hybridrechnungen aus Word-Rechnungen

Das Skript macht aus einer Rechnung im Word-Format (`.docx`) eine **ZUGFeRD-/Factur-X-Hybridrechnung**: eine PDF/A-3-Datei, die wie die Word-Rechnung aussieht und zusätzlich die Rechnungsdaten als  maschinenlesbares XML (`factur-x.xml`, Profil **EN 16931**) enthält. Solche Rechnungen erfüllen die Anforderungen an eine E-Rechnung im Sinne von § 14 UStG und lassen sich von der Buchhaltungssoftware  des Empfängers automatisch einlesen.

```
zugferd word-rechnung        →   word-rechnung.pdf   (PDF/A-3 mit eingebettetem ZUGFeRD-XML)
```

## Inhalt

- [Funktionsweise](#funktionsweise)
- [Systemvoraussetzungen](#systemvoraussetzungen)
- [Installation](#installation)
- [Benutzung](#benutzung)
- [Anforderungen an die Word-Rechnung](#anforderungen-an-die-word-rechnung)
- [Umsatzsteuer: Normalsatz, Befreiungen, Reverse Charge](#umsatzsteuer-normalsatz-befreiungen-reverse-charge)
- [Prüfungen und Meldungen](#prüfungen-und-meldungen)
- [Layout: Abweichungen zwischen Word und LibreOffice](#layout-abweichungen-zwischen-word-und-libreoffice)
- [Grenzen](#grenzen)
- [Ergebnis prüfen](#ergebnis-prüfen)
- [Aufbau des Codes](#aufbau-des-codes)

## Funktionsweise

1. **Auslesen** – Die Rechnungsdaten werden aus der Word-Datei gelesen: Rechnungsnummer und -datum, Absender (aus der Absenderzeile und dem Briefkopf), Empfänger (aus dem Adressfeld), Positionen und Summen (aus der Positionstabelle), Bankverbindung, Steuernummer/USt-IdNr., Zahlungsbedingungen, Leistungszeitraum, Projekt und Vertragsgrundlage. Deutsche und englische Rechnungen werden erkannt.
2. **Berechnen und abgleichen** – Nettobeträge, Umsatzsteuer und Bruttobetrag werden aus den Positionen neu berechnet und mit den Summen im Dokument verglichen. Bei Abweichungen bricht das Skript ab.
3. **XML erzeugen** – Aus den Daten entsteht das ZUGFeRD-XML (UN/CEFACT CII, Profil EN 16931). Es wird gegen das offizielle XML-Schema geprüft (sofern das Paket `factur-x` installiert ist).
4. **PDF erzeugen** – LibreOffice wandelt die Word-Datei in ein PDF/A-3 um.
5. **Einbetten** – Das XML wird als Anhang `factur-x.xml` in die PDF eingebettet, die PDF/A- und Factur-X-Metadaten (XMP) werden ergänzt.

Die Word-Datei selbst wird nie verändert; alle Anpassungen erfolgen an temporären Kopien.

## Systemvoraussetzungen

| Komponente | Pflicht | Hinweis |
|---|---|---|
| Python 3.9 oder neuer | ja | |
| LibreOffice | ja | wird automatisch gefunden (PATH, Standardordner, unter macOS auch per Spotlight, unter Windows über die Registry) |
| Python-Pakete aus `requirements.txt` | ja | `pikepdf`, `lxml`, `python-docx`, `pypdfium2`; optional `factur-x` |
| Ghostscript | nein | nur für `--pdf` mit einer PDF, die noch kein PDF/A-3 ist |
| Schriften der Word-Vorlage | empfohlen | sonst ersetzt LibreOffice sie (siehe [Layout](#layout-abweichungen-zwischen-word-und-libreoffice)) |

Die Python-Pakete bringen ihre C-Bibliotheken (qpdf, libxml2, PDFium) mit; weitere Systembibliotheken sind nicht nötig. Entwickelt und getestet unter Linux (LibreOffice 24.2) sowie im Einsatz unter macOS (LibreOffice 26.2).

## Installation

### macOS

```bash
# LibreOffice (falls noch nicht vorhanden)
brew install --cask libreoffice
# optional: Ghostscript
brew install ghostscript

# Projektordner mit virtueller Python-Umgebung
cd zugferd
python3 -m venv .venv
.venv/bin/pip install -r requirements.txt
```

Damit der Befehl `zugferd` von überall funktioniert, z. B. in `~/.zshrc`:

```bash
alias zugferd='/Pfad/zu/zugferd/.venv/bin/python /Pfad/zu/zugferd/zugferd_embed.py'
```

### Windows / Linux

Wie unter macOS; LibreOffice über de.libreoffice.org bzw. den Paketmanager installieren
(`sudo apt install libreoffice`), virtuelle Umgebung mit `python -m venv .venv` anlegen und
`.venv\Scripts\pip install -r requirements.txt` (Windows) bzw. `.venv/bin/pip …` (Linux) ausführen.

`zugferd_embed.py` und der Ordner `library/` müssen nebeneinander liegen.

## Benutzung

```bash
zugferd word-rechnung                                # word-rechnung.docx -> word-rechnung.pdf (im selben Ordner)
zugferd word-rechnung.docx                           # gleichwertig
zugferd word-rechnung -o ~/Rechnungen/               # PDF in einem anderen Ordner ablegen
zugferd word-rechnung -o rechnung.pdf                # PDF unter anderem Namen speichern
zugferd word-rechnung --xml-out word-rechnung.xml    # das erzeugte XML zusätzlich speichern
zugferd word-rechnung -q                             # ohne Übersicht der Rechnungsbestandteile
```

| Option | Bedeutung |
|---|---|
| `datei` bzw. `--docx datei` | Word-Rechnung; die Endung `.docx` kann entfallen |
| `-o datei.pdf` / `-o ordner/` | anderer Name bzw. Zielordner (Standard: Name der Word-Datei mit `.pdf`) |
| `--xml-out datei.xml` | erzeugtes XML zusätzlich speichern |
| `--pdf datei.pdf` | statt der LibreOffice-Umwandlung ein vorhandenes PDF verwenden |
| `--soffice pfad` | bestimmte LibreOffice-Installation verwenden |
| `-q` | Übersicht der erkannten Rechnungsbestandteile nicht ausgeben |

Eine bereits vorhandene PDF gleichen Namens wird ohne Rückfrage überschrieben.

### Ausgabe

Bei jedem Lauf zeigt das Skript alle erkannten bzw. berechneten Rechnungsbestandteile, sortiert nach den Feldnummern der EN 16931 (BT = einzelnes Feld, BG = Gruppe), mit deutscher Bezeichnung und Erklärung der Codes:

```
BT-1    Rechnungsnummer            202603002
BT-2    Rechnungsdatum             1.11.2026
BT-3    Rechnungsart               380  (Rechnung)
BG-4    Verkäufer
  BT-27   Name                       Tiny Timmy Rechtsanwalt
  BT-31   USt-IdNr.                  DE202192389
  …
BG-22   Gesamtbeträge (berechnet)
  BT-112  Gesamtbetrag brutto        247,52
…
```

Bricht das Skript wegen fehlender Angaben ab, wird die Übersicht trotzdem angezeigt – so ist sofort erkennbar, was gelesen wurde und was fehlt.

## Anforderungen an die Word-Rechnung

Das Auslesen beruht auf typischen Formulierungen und Layouts deutscher und englischer Rechnungen. Erwartet werden:

| Angabe | Beispiele, die erkannt werden |
|---|---|
| Rechnungsnummer | `Rechnung Nr. 202603002`, `Invoice No. 202603002`, `Gutschrift Nr. …`, `Credit Note No. …` |
| Rechnungsdatum | `Berlin, 1. November 2026`, `Stuttgart, 11.11.2026`, `Datum: …`, `Date: November 10, 2026` |
| Absenderzeile | `RA Tiny Timmy · Sesamstr. 21 · 10095 Berlin` (Trenner `·`, `•` oder `\|`) |
| Empfänger | Adressfeld unter der Absenderzeile: Name, ggf. Zusatz, Straße, Ort mit PLZ, ggf. Land – auch ausländische Formate wie `1010 Wien`, `Falmouth, England, TR11 5TR` oder nur `Madrid` |
| Briefkopf | `Telefon:`/`Phone:`, E-Mail, `Steuer-Nr.:`/`Tax ID:`, **`USt-IdNr.:`/`VAT ID:`**, `IBAN:`, `BIC:` |
| Positionstabelle | Kopfzeile mit `Position`/`Description` und `Betrag`/`Amount`; optional Spalten für Menge (`Anzahl Stunden`, `Quantity`, `Hours` …), Einheit und Einzelpreis (`Stundensatz`, `Unit price`, `Rate` …) |
| Summenzeilen | `Zwischensumme Netto`/`Subtotal`, `+ Umsatzsteuer (19%)`/`VAT (19%)`, `Rechnungsbetrag Brutto`/`Total` |
| Beträge | `EUR 1.234,56`, `GBP 2.345,67`, `£ 2,345.67`; die Währung wird aus den Beträgen übernommen |
| Leistungszeitraum | in der Position (`Januar 2026`, `(20.8.2026-19.9.2026)`, `(1.-10.11.2026)`) oder als `Leistungszeitraum:`/`Service Period:` |
| Zahlungsbedingung | `Zahlbar innerhalb von 30 Tagen`, `Zahlbar mit Rechnungserhalt`, `Payable upon receipt`, `Payable within 14 days` – daraus wird das Fälligkeitsdatum berechnet |
| Projekt / Vertrag | `Projekt:`/`Project:` und `Grundlage:`/`Contract:`/`Agreement:` |

**Pflicht ist die USt-IdNr. des Rechnungsstellers** im Briefkopf (Regel BR-CO-26). Fehlt sie, wird keine PDF erzeugt.

Gerundete Mengen (z. B. `0,87` Stunden bei 52 Minuten) werden erkannt: Das Skript leitet die genaue Menge aus Betrag und Einzelpreis her, damit Menge × Preis im XML exakt den Positionsbetrag ergibt, und gibt dazu einen Hinweis aus.

## Umsatzsteuer: Normalsatz, Befreiungen, Reverse Charge

Die Steuerkategorie wird aus dem Rechnungstext erkannt (der Briefkopf wird dabei nicht berücksichtigt):

| Vermerk in der Rechnung (Beispiele) | Kategorie | Code |
|---|---|---|
| kein Vermerk, Umsatzsteuerzeile mit Satz | S – Normalsatz | – |
| „Steuerschuldnerschaft des Leistungsempfängers“, „Reverse charge“, „§ 13b UStG“ | AE – Reverse Charge | VATEX-EU-AE |
| „innergemeinschaftliche Lieferung“, „§ 6a UStG“, „intra-community supply“ | K | VATEX-EU-IC |
| „Ausfuhrlieferung“, „§ 6 UStG“, „export delivery“ | G | VATEX-EU-G |
| „§ 19 UStG“, „Kleinunternehmer“, „steuerfrei gemäß …“, „VAT exempt“ | E – steuerbefreit | – |
| „nicht steuerbar“, „not subject to German VAT“ | O – nicht steuerbar | VATEX-EU-O |

Der Satz mit dem Vermerk wird als Befreiungsgrund ins XML übernommen.

Besonderheiten:

- **Reverse Charge (AE)** verlangt die USt-IdNr. des Kunden in der Rechnung, z. B. `USt-IdNr. des Leistungsempfängers: ATU12345678` oder `Recipient VAT ID: 123 3456 99`  (fehlt das Länderkürzel, wird es aus dem Land des Kunden ergänzt).
- **Kunden außerhalb der EU** (z. B. Türkei, Großbritannien): Ein Reverse-Charge-Vermerk wird als **nicht steuerbar (O)** ausgegeben. Nach EN 16931 enthält das XML dann keine USt-IdNr. (weder die eigene noch die des Kunden, Regel BR-O-02); als Verkäuferkennung dient die Steuernummer. Für solche Rechnungen ist ein Vermerk wie *„Not subject to German VAT (place of supply outside Germany).“* treffender als „Reverse charge“.
- Nennt der Text eine Befreiung, die Tabelle weist aber Umsatzsteuer aus, bricht das Skript ab.

## Prüfungen und Meldungen

Das Skript erzeugt **keine PDF**, wenn

- Rechnungsnummer, Rechnungsdatum, Absenderzeile, Empfänger oder Positionstabelle nicht erkannt werden,
- die USt-IdNr. des Rechnungsstellers fehlt oder ein ungültiges Format hat,
- die Summen im Dokument nicht zur Berechnung passen,
- Menge × Einzelpreis nicht den Positionsbetrag ergibt (außer bei erkennbarer Rundung),
- die IBAN-Prüfsumme falsch ist,
- bei Reverse Charge die USt-IdNr. des Kunden fehlt,
- Steuervermerk und Umsatzsteuerzeile sich widersprechen,
- mehrere Währungen in der Positionstabelle vorkommen.

**Hinweise** (die PDF wird trotzdem erzeugt) erscheinen z. B. bei gerundeten Mengen, bei der Umstellung auf „nicht steuerbar“ für Kunden außerhalb der EU oder wenn Schriften ersetzt wurden.

## Grenzen

- Nur **ein Steuersatz** je Rechnung (keine gemischten Rechnungen mit 19 % und 7 % oder steuerpflichtigen und steuerfreien Positionen).
- Keine Nachlässe/Zuschläge auf Rechnungsebene, keine Anzahlungs- oder Schlussrechnungen mit Verrechnung, keine Anhänge.
- Profil **EN 16931 (ZUGFeRD COMFORT)**, nicht **XRechnung**. Öffentliche Auftraggeber verlangen in der Regel eine XRechnung mit Leitweg-ID (BT-10); die XRechnung-spezifischen Hinweise eines Validators (BR-DE-…) sind für ZUGFeRD unerheblich.
- Rechnungen in Fremdwährung mit ausgewiesener Umsatzsteuer: Der Steuerbetrag wird im XML nicht zusätzlich in Euro angegeben (Felder BT-6/BT-111). Bei Reverse Charge bzw. 0 % ohne Bedeutung.
- Die Gültigkeit der USt-IdNr. (Bestand beim BZSt) wird nicht geprüft, nur ihr Format.
- Das Auslesen ist heuristisch. Bei neuen Vorlagen die Übersicht beim ersten Lauf kontrollieren.

## Ergebnis prüfen

Vor dem ersten Versand neuer Rechnungsvorlagen empfiehlt sich eine Prüfung mit einem Validator, z. B. dem von RechneX (https://rechnex.de/zugferd-validator). Bitte dabei bedenken, dass die Rechnung dafür in der Regel an einen fremden Server übertragen wird. Das Skript selbst prüft das XML nur gegen das Schema, nicht gegen alle Geschäftsregeln (Schematron).

## Aufbau des Codes

```
zugferd/
├── zugferd_embed.py       Hauptskript: Befehlszeile und Fehlerausgabe
├── requirements.txt       Python-Pakete
├── README.md
└── library/
    ├── pipeline.py        Gesamtablauf (create_hybrid_invoice) und Ausgabedateiname
    ├── docx_reader.py     Rechnungsdaten aus der Word-Datei lesen
    ├── xml_builder.py     ZUGFeRD-XML erzeugen und gegen das Schema prüfen
    ├── pdf_render.py      Word → PDF/A-3 mit LibreOffice, Adressfeld-Ausrichtung, Schriftprüfung
    ├── pdfa.py            PDF/A-3 per Ghostscript, XML-Anhang und Metadaten einbetten
    ├── display.py         Übersicht nach BT-/BG-Nummern
    ├── values.py          Beträge, Datumsangaben, Zeiträume, Währungen, IBAN
    ├── constants.py       Länder, Monate, Einheiten, Erkennungsmuster für Steuervermerke
    └── errors.py          Fehlerklasse InvoiceError
```

Die Bibliothek lässt sich auch aus anderen Python-Skripten verwenden:

```python
from library.pipeline import create_hybrid_invoice
from library.errors import InvoiceError

try:
    daten = create_hybrid_invoice("RE 4711.docx", "RE 4711.pdf", show=False)
except InvoiceError as e:
    print("Rechnung nicht erzeugt:", e)
```

Fehler werden in der Bibliothek als `InvoiceError` gemeldet (mit anwenderfreundlichem Text); das aufrufende Programm wird dabei nicht beendet.
