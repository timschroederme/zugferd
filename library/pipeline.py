"""Gesamtablauf: Word-Rechnung -> Rechnungsdaten -> ZUGFeRD-XML -> PDF/A-3 -> Hybridrechnung."""

import os
import tempfile

from .display import show_invoice
from .docx_reader import extract_from_docx
from .errors import InvoiceError
from .pdf_render import docx_to_pdfa3, warn_font_substitution
from .pdfa import embed
from .xml_builder import build_cii_xml, check_xsd


def default_output_path(docx_path: str, out: str = None) -> str:
    """Ausgabepfad: Name der Word-Datei mit .pdf (im selben Ordner bzw. im Ordner `out`)."""
    name = os.path.splitext(os.path.basename(docx_path))[0] + ".pdf"
    if not out:
        return os.path.join(os.path.dirname(os.path.abspath(docx_path)), name)
    if os.path.isdir(out):
        return os.path.join(out, name)
    return out


def create_hybrid_invoice(docx_path: str, pdf_out: str, pdf_in: str = None, soffice: str = None,
                          xml_out: str = None, show: bool = True) -> dict:
    """Erzeugt die ZUGFeRD-Hybridrechnung `pdf_out` aus der Word-Rechnung `docx_path`.

    pdf_in   vorhandenes PDF statt der Umwandlung mit LibreOffice
    soffice  LibreOffice-Programm (sonst automatische Suche)
    xml_out  erzeugtes XML zusätzlich speichern
    show     Übersicht der erkannten Rechnungsbestandteile ausgeben (auch im Fehlerfall)
    Gibt die Rechnungsdaten zurück; wirft InvoiceError bei unvollständigen/ungültigen Angaben."""
    data = extract_from_docx(docx_path)
    try:
        xml = build_cii_xml(data)
    except InvoiceError:
        if show:
            show_invoice(data)  # auch bei Fehlern zeigen, was erkannt wurde
        raise
    if show:
        show_invoice(data)
    check_xsd(xml)
    if xml_out:
        with open(xml_out, "wb") as f:
            f.write(xml)

    tmp_pdf = None
    if not pdf_in:
        fd, tmp_pdf = tempfile.mkstemp(suffix=".pdf")
        os.close(fd)
        print("Erzeuge PDF/A-3 mit LibreOffice ...")
        docx_to_pdfa3(docx_path, tmp_pdf, soffice)
        warn_font_substitution(docx_path, tmp_pdf)
        pdf_in = tmp_pdf
    try:
        embed(pdf_in, xml, pdf_out)
    finally:
        if tmp_pdf and os.path.exists(tmp_pdf):
            os.remove(tmp_pdf)
    return data
