"""Bibliothek für ZUGFeRD-/Factur-X-Hybridrechnungen aus Word-Rechnungen.

Module:
  pipeline     Gesamtablauf (create_hybrid_invoice) – vom Hauptskript aufgerufen
  docx_reader  Rechnungsdaten aus der Word-Datei lesen
  xml_builder  ZUGFeRD-XML (EN 16931) erzeugen und prüfen
  pdf_render   Word -> PDF/A-3 mit LibreOffice
  pdfa         PDF/A-3 herstellen, XML einbetten
  display      Übersicht der Rechnungsbestandteile im Terminal
  values       Beträge, Datumsangaben, Zeiträume lesen/formatieren
  constants    Länder, Monate, Einheiten, Muster
  errors       InvoiceError
"""
