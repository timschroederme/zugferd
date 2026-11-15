"""PDF/A-3 prüfen bzw. mit Ghostscript herstellen und das XML als ZUGFeRD-Anhang einbetten."""

import glob
import os
import shutil
import subprocess
import tempfile
from datetime import datetime, timezone

import pikepdf
from pikepdf import Array, Dictionary, Name, Stream, String

from .errors import InvoiceError


XML_NAME = "factur-x.xml"  # ZUGFeRD 2.x / Factur-X; (ZUGFeRD 1.0 hieß ZUGFeRD-invoice.xml)


FX_NS = "urn:factur-x:pdfa:CrossIndustryDocument:invoice:1p0#"


# PDF/A-Extension-Schema: erklärt PDF/A-Readern die fx:-Properties.
# Ohne diesen Block ist die XMP-Datei nicht PDF/A-konform.
EXT_SCHEMA = f"""<rdf:Description rdf:about=""
    xmlns:pdfaExtension="http://www.aiim.org/pdfa/ns/extension/"
    xmlns:pdfaSchema="http://www.aiim.org/pdfa/ns/schema#"
    xmlns:pdfaProperty="http://www.aiim.org/pdfa/ns/property#">
  <pdfaExtension:schemas><rdf:Bag><rdf:li rdf:parseType="Resource">
    <pdfaSchema:schema>Factur-X PDFA Extension Schema</pdfaSchema:schema>
    <pdfaSchema:namespaceURI>{FX_NS}</pdfaSchema:namespaceURI>
    <pdfaSchema:prefix>fx</pdfaSchema:prefix>
    <pdfaSchema:property><rdf:Seq>
      <rdf:li rdf:parseType="Resource"><pdfaProperty:name>DocumentFileName</pdfaProperty:name><pdfaProperty:valueType>Text</pdfaProperty:valueType><pdfaProperty:category>external</pdfaProperty:category><pdfaProperty:description>Name des eingebetteten XML</pdfaProperty:description></rdf:li>
      <rdf:li rdf:parseType="Resource"><pdfaProperty:name>DocumentType</pdfaProperty:name><pdfaProperty:valueType>Text</pdfaProperty:valueType><pdfaProperty:category>external</pdfaProperty:category><pdfaProperty:description>INVOICE</pdfaProperty:description></rdf:li>
      <rdf:li rdf:parseType="Resource"><pdfaProperty:name>Version</pdfaProperty:name><pdfaProperty:valueType>Text</pdfaProperty:valueType><pdfaProperty:category>external</pdfaProperty:category><pdfaProperty:description>Version des XMP-Schemas</pdfaProperty:description></rdf:li>
      <rdf:li rdf:parseType="Resource"><pdfaProperty:name>ConformanceLevel</pdfaProperty:name><pdfaProperty:valueType>Text</pdfaProperty:valueType><pdfaProperty:category>external</pdfaProperty:category><pdfaProperty:description>Profil der Rechnung</pdfaProperty:description></rdf:li>
    </rdf:Seq></pdfaSchema:property>
  </rdf:li></rdf:Bag></pdfaExtension:schemas>
</rdf:Description>"""


def is_pdfa3(path: str) -> bool:
    """Prüft im XMP-Block, ob pdfaid:part == 3 gesetzt ist."""
    with pikepdf.open(path) as pdf:
        if "/Metadata" not in pdf.Root:
            return False
        return pdf.open_metadata().get("pdfaid:part") == "3"


def find_ghostscript() -> str:
    for name in ("gs", "gswin64c", "gswin32c"):
        exe = shutil.which(name)
        if exe:
            return exe
    for exe in ("/opt/homebrew/bin/gs", "/usr/local/bin/gs", "/opt/local/bin/gs", "/usr/bin/gs",
                *sorted(glob.glob(r"C:\Program Files\gs\gs*\bin\gswin64c.exe"), reverse=True)):
        if os.path.isfile(exe):
            return exe
    raise InvoiceError("Ghostscript nicht gefunden. Es wird gebraucht, um die PDF in PDF/A-3 umzuwandeln.\n"
             "  macOS:   brew install ghostscript   (Homebrew: https://brew.sh)\n"
             "  Windows: https://ghostscript.com/releases/gsdnld.html\n"
             "  Linux:   sudo apt install ghostscript\n"
             "Nur nötig für --pdf mit einem PDF, das noch kein PDF/A-3 ist; Word-Dateien (--docx) brauchen es nicht.")


def find_icc(gs_exe: str) -> str:
    """Sucht das mit Ghostscript gelieferte sRGB-Profil (default_rgb.icc)."""
    base = os.path.dirname(os.path.dirname(os.path.realpath(gs_exe)))
    patterns = [
        "/usr/share/ghostscript/*/iccprofiles/default_rgb.icc",
        "/usr/local/share/ghostscript/*/iccprofiles/default_rgb.icc",
        "/opt/homebrew/share/ghostscript/*/iccprofiles/default_rgb.icc",
        "/opt/homebrew/share/ghostscript/iccprofiles/default_rgb.icc",
        "/opt/homebrew/Cellar/ghostscript/*/share/ghostscript/*/iccprofiles/default_rgb.icc",
        "/opt/homebrew/Cellar/ghostscript/*/share/ghostscript/iccprofiles/default_rgb.icc",
        "/usr/local/Cellar/ghostscript/*/share/ghostscript/*/iccprofiles/default_rgb.icc",
        "/usr/local/share/ghostscript/iccprofiles/default_rgb.icc",
        "/opt/local/share/ghostscript/*/iccprofiles/default_rgb.icc",
        os.path.join(base, "share", "ghostscript", "iccprofiles", "default_rgb.icc"),
        os.path.join(base, "share", "ghostscript", "*", "iccprofiles", "default_rgb.icc"),
        os.path.join(base, "lib", "iccprofiles", "default_rgb.icc"),  # Windows
        os.path.join(base, "..", "lib", "iccprofiles", "default_rgb.icc"),
    ]
    for pat in patterns:
        hits = sorted(glob.glob(pat))
        if hits:
            return hits[-1]
    raise InvoiceError("sRGB-ICC-Profil (default_rgb.icc) von Ghostscript nicht gefunden.")


def convert_to_pdfa3(pdf_in: str, pdf_out: str) -> None:
    """Konvertiert beliebiges PDF mit Ghostscript nach PDF/A-3b (sRGB-OutputIntent)."""
    gs = find_ghostscript()
    icc = find_icc(gs).replace("\\", "/")
    # PS-Datei, die den OutputIntent (Pflicht bei PDF/A) mit dem ICC-Profil anlegt.
    icc_ps = icc.replace("\\", "\\\\").replace("(", "\\(").replace(")", "\\)")
    pdfa_def = f"""%!
/ICCProfile ({icc_ps}) def
[/_objdef {{icc_PDFA}} /type /stream /OBJ pdfmark
[{{icc_PDFA}} <</N 3>> /PUT pdfmark
[{{icc_PDFA}} ICCProfile (r) file /PUT pdfmark
[/_objdef {{OutputIntent_PDFA}} /type /dict /OBJ pdfmark
[{{OutputIntent_PDFA}} <<
  /Type /OutputIntent /S /GTS_PDFA1
  /DestOutputProfile {{icc_PDFA}}
  /OutputConditionIdentifier (sRGB) /Info (sRGB)
  /RegistryName (http://www.color.org)
>> /PUT pdfmark
[{{Catalog}} <</OutputIntents [ {{OutputIntent_PDFA}} ]>> /PUT pdfmark
"""
    with tempfile.TemporaryDirectory() as tmp:
        def_ps = os.path.join(tmp, "PDFA_def.ps")
        with open(def_ps, "w", encoding="utf-8") as f:
            f.write(pdfa_def)
        cmd = [
            gs, "-dBATCH", "-dNOPAUSE", "-dQUIET", "-dNOSAFER",
            "-sDEVICE=pdfwrite",
            "-dPDFA=3",                       # PDF/A-3
            "-dPDFACompatibilityPolicy=1",    # nicht konforme Elemente verwerfen statt abbrechen
            "-sColorConversionStrategy=RGB",
            "-sProcessColorModel=DeviceRGB",
            f"-sOutputFile={pdf_out}",
            def_ps, pdf_in,
        ]
        res = subprocess.run(cmd, capture_output=True, text=True)
        if res.returncode != 0 or not os.path.exists(pdf_out):
            raise InvoiceError(f"Ghostscript-Konvertierung fehlgeschlagen:\n{res.stdout}\n{res.stderr}")


def embed(pdf_in: str, xml_bytes: bytes, pdf_out: str, level: str = "EN 16931") -> None:

    # 0) Falls nötig zuerst nach PDF/A-3b konvertieren.
    tmp_pdfa = None
    if not is_pdfa3(pdf_in):
        print("Eingabe ist kein PDF/A-3 -> konvertiere mit Ghostscript ...")
        fd, tmp_pdfa = tempfile.mkstemp(suffix=".pdf")
        os.close(fd)
        convert_to_pdfa3(pdf_in, tmp_pdfa)
        pdf_in = tmp_pdfa

    now = datetime.now(timezone.utc)
    pdf_date = now.strftime("D:%Y%m%d%H%M%SZ")

    pdf = pikepdf.open(pdf_in)

    # 1) Das XML als "Embedded File Stream" ablegen.
    #    Subtype "text/xml" (Slash als #2F kodiert, macht pikepdf automatisch über Name).
    ef_stream = Stream(pdf, xml_bytes)
    ef_stream["/Type"] = Name.EmbeddedFile
    ef_stream["/Subtype"] = Name("/text/xml")  # pikepdf kodiert "/" selbst als #2F (nicht vorab kodieren!)
    ef_stream["/Params"] = Dictionary(
        Size=len(xml_bytes), ModDate=String(pdf_date)
    )
    # Kompression ist erlaubt; hier bewusst unkomprimiert -> leichter lesbar/debuggbar.

    # 2) Filespec-Objekt: beschreibt die Anlage.
    #    AFRelationship: "Data" (reine Daten) | "Source" | "Alternative" (XML ist
    #    alternative Darstellung desselben Inhalts) | "Supplement".
    #    ZUGFeRD/Factur-X verlangt "Alternative" bzw. in der Praxis oft "Data".
    filespec = pdf.make_indirect(
        Dictionary(
            Type=Name.Filespec,
            F=String(XML_NAME),
            UF=String(XML_NAME),  # Unicode-Dateiname
            Desc=String("Factur-X/ZUGFeRD Rechnungsdaten"),
            AFRelationship=Name.Alternative,
            EF=Dictionary(F=ef_stream, UF=ef_stream),
        )
    )

    # 3) Anlage im Katalog registrieren: zwei Stellen.
    #    a) /Names/EmbeddedFiles (Name-Tree) -> sichtbar in Acrobat als Dateianhang
    #    b) /AF (Associated Files) -> von PDF/A-3 verlangt, verknüpft Anlage mit Dokument
    root = pdf.Root
    names = root.get("/Names")
    if names is None:
        names = root["/Names"] = Dictionary()
    names["/EmbeddedFiles"] = Dictionary(Names=Array([String(XML_NAME), filespec]))
    root["/AF"] = Array([filespec])

    # 4) XMP-Metadaten ergänzen: fx-Namespace + PDF/A-Extension-Schema.
    #    pikepdf's open_metadata() erhält vorhandene Einträge (u. a. pdfaid:part=3).
    with pdf.open_metadata(set_pikepdf_as_editor=False, update_docinfo=True) as meta:
        meta["pdfaid:part"] = "3"
        meta["pdfaid:conformance"] = "B"
        meta["{%s}DocumentFileName" % FX_NS] = XML_NAME
        meta["{%s}DocumentType" % FX_NS] = "INVOICE"
        meta["{%s}Version" % FX_NS] = "1.0"
        meta["{%s}ConformanceLevel" % FX_NS] = level
    inject_extension_schema(pdf)

    pdf.save(pdf_out)
    pdf.close()
    if tmp_pdfa:
        os.remove(tmp_pdfa)


def inject_extension_schema(pdf: pikepdf.Pdf) -> None:
    """Fügt den pdfaExtension-Block in den rdf:RDF-Teil des XMP-Streams ein."""
    xmp = pdf.Root.Metadata
    data = xmp.read_bytes().decode("utf-8")
    if "pdfaExtension:schemas" in data:
        return
    data = data.replace("</rdf:RDF>", EXT_SCHEMA + "</rdf:RDF>", 1)
    xmp.write(data.encode("utf-8"))
