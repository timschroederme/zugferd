"""Fehlerklasse der Bibliothek."""


class InvoiceError(Exception):
    """Rechnungsdaten unvollständig/widersprüchlich oder Verarbeitung fehlgeschlagen.
    Die Meldung ist für Anwender formuliert und kann direkt ausgegeben werden."""
