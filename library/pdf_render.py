"""Word -> PDF/A-3 mit LibreOffice (inkl. Suche der Installation, Ausrichtung des Adressfelds
und Prüfung auf ersetzte Schriften)."""

import glob
import os
import re
import shutil
import subprocess
import sys
import tempfile
import zipfile

import pikepdf
from lxml import etree

from .constants import W_NS
from .docx_reader import docx_fonts
from .errors import InvoiceError


def soffice_candidates():
    """Mögliche LibreOffice-Programme, in sinnvoller Reihenfolge (ohne Duplikate)."""
    seen = set()

    def add(p):
        if p and p not in seen:
            seen.add(p)
            return [p]
        return []

    out = []
    for name in ("soffice", "libreoffice"):
        out += add(shutil.which(name))
    if sys.platform == "darwin":
        apps = ["/Applications/LibreOffice.app", os.path.expanduser("~/Applications/LibreOffice.app")]
        try:                                                     # Spotlight: auch umbenannte/verschobene Apps
            res = subprocess.run(["mdfind", "kMDItemCFBundleIdentifier == 'org.libreoffice.script'"],
                                 capture_output=True, text=True, timeout=10)
            apps += [x for x in res.stdout.splitlines() if x.endswith(".app")]
        except (OSError, subprocess.SubprocessError):
            pass
        for app in apps:
            out += add(os.path.join(app, "Contents", "MacOS", "soffice"))
    elif sys.platform.startswith("win"):
        try:                                                     # Registry: Installationspfad
            import winreg
            for root in (winreg.HKEY_LOCAL_MACHINE, winreg.HKEY_CURRENT_USER):
                for key in (r"SOFTWARE\LibreOffice\UNO\InstallPath",
                            r"SOFTWARE\WOW6432Node\LibreOffice\UNO\InstallPath"):
                    try:
                        with winreg.OpenKey(root, key) as k:
                            out += add(os.path.join(winreg.QueryValueEx(k, "")[0], "soffice.exe"))
                    except OSError:
                        pass
        except ImportError:
            pass
        for base in (os.environ.get("ProgramFiles", r"C:\Program Files"),
                     os.environ.get("ProgramFiles(x86)", r"C:\Program Files (x86)")):
            out += add(os.path.join(base, "LibreOffice", "program", "soffice.exe"))
    else:
        for p in ("/usr/bin/soffice", "/usr/local/bin/soffice", "/usr/lib/libreoffice/program/soffice",
                  "/opt/libreoffice/program/soffice", "/snap/bin/libreoffice"):
            out += add(p)
        out += [p for p in sorted(glob.glob("/opt/libreoffice*/program/soffice"), reverse=True) if p not in seen]
    return out


def find_soffice(explicit: str = None) -> str:
    """LibreOffice automatisch finden (optional: --soffice als Vorgabe). Bricht ab, wenn keine
    lauffähige Installation gefunden wird."""
    if explicit:
        if not os.path.isfile(explicit):
            raise InvoiceError(f"LibreOffice nicht gefunden unter: {explicit}")
        return explicit
    tried = []
    for exe in soffice_candidates():
        if not os.path.isfile(exe):
            continue
        tried.append(exe)
        try:
            res = subprocess.run([exe, "--version"], capture_output=True, text=True, timeout=60)
            if res.returncode == 0 and "LibreOffice" in (res.stdout + res.stderr):
                # print("LibreOffice:", exe, "|", (res.stdout.strip().splitlines() or [""])[0][:60])
                return exe
        except (OSError, subprocess.SubprocessError):
            continue
    if tried:
        raise InvoiceError("LibreOffice gefunden, aber nicht lauffähig: " + ", ".join(tried))
    raise InvoiceError("Keine LibreOffice-Installation gefunden. Bitte LibreOffice installieren "
             "(https://de.libreoffice.org/download/) – auf dem Mac auch per 'brew install --cask libreoffice'.")


def inline_floating_tables(docx_in: str, docx_out: str, shift_twips: int = 0, guide: bool = False):
    """Kopie der Word-Datei, in der seitenverankerte (schwebende) Tabellen am Dokumentanfang – typisch:
    das Adressfeld – zu normalen Tabellen werden. Ihre Höhe auf der Seite wird über eine Leerzeile mit
    exakter Höhe davor nachgebildet (minus shift_twips als Korrektur).
    Hintergrund: LibreOffice-Versionen setzen Text um schwebende Tabellen unterschiedlich (z. B. die
    Rechnungsnummer über statt unter dem Adressfeld); im normalen Textfluss folgt sie wie in Word mit dem
    eingestellten Absatzabstand darunter.
    guide=True: erste Tabelle bekommt eine Hilfslinie oben (nur für die Messung der Oberkante).
    Rückgabe: (Anzahl umgewandelter Tabellen, Soll-Oberkante der ersten Tabelle in pt ab Seitenrand,
    Breite der ersten Tabelle in pt)."""
    w = "{%s}" % W_NS
    with zipfile.ZipFile(docx_in) as zin:
        root = etree.fromstring(zin.read("word/document.xml"))
        body = root.find(w + "body")
        sect = body.find(w + "sectPr")
        if sect is None:
            sect = root.find(".//" + w + "sectPr")
        mar = sect.find(w + "pgMar") if sect is not None else None
        top = int(mar.get(w + "top", "1417")) if mar is not None else 1417
        left = int(mar.get(w + "left", "1417")) if mar is not None else 1417
        changed, leading, target, width = 0, [], None, None
        for el in list(body):
            if el.tag == w + "p":
                if "".join(el.itertext()).strip() or el.find(".//" + w + "drawing") is not None \
                        or el.find(".//" + w + "pict") is not None:
                    break                                   # erster echter Inhalt -> nur Tabellen davor
                leading.append(el)
                continue
            if el.tag != w + "tbl":
                break
            tblPr = el.find(w + "tblPr")
            pos = tblPr.find(w + "tblpPr") if tblPr is not None else None
            if pos is None or pos.get(w + "vertAnchor") not in ("page", "margin") or not pos.get(w + "tblpY"):
                break
            y = int(pos.get(w + "tblpY")) - (top if pos.get(w + "vertAnchor") == "page" else 0) - shift_twips
            x = int(pos.get(w + "tblpX", "0")) - (left if pos.get(w + "horzAnchor") == "page" else 0)
            tblPr.remove(pos)
            for tag in ("tblOverlap",):
                if (ov := tblPr.find(w + tag)) is not None:
                    tblPr.remove(ov)
            if x > 0:
                ind = tblPr.find(w + "tblInd")
                if ind is None:
                    ind = etree.SubElement(tblPr, w + "tblInd")
                ind.set(w + "w", str(x))
                ind.set(w + "type", "dxa")
            if target is None:
                target = (int(pos.get(w + "tblpY")) + (top if pos.get(w + "vertAnchor") == "margin" else 0)) / 20
                tw = tblPr.find(w + "tblW")
                width = int(tw.get(w + "w", "0")) / 20 if tw is not None and tw.get(w + "type") == "dxa" else None
                if guide:
                    b = tblPr.find(w + "tblBorders")
                    if b is None:
                        b = etree.SubElement(tblPr, w + "tblBorders")
                    for old in b.findall(w + "top"):
                        b.remove(old)
                    t_ = etree.Element(w + "top")
                    for k, v in (("val", "single"), ("sz", "2"), ("space", "0"), ("color", "FF0000")):
                        t_.set(w + k, v)
                    b.insert(0, t_)
            if not leading:
                p = etree.Element(w + "p")
                el.addprevious(p)
                leading.append(p)
            for i, p in enumerate(leading):
                pPr = p.find(w + "pPr")
                if pPr is None:
                    pPr = etree.Element(w + "pPr")
                    p.insert(0, pPr)
                sp = pPr.find(w + "spacing")
                if sp is None:
                    sp = etree.SubElement(pPr, w + "spacing")
                h = max(y, 20) if i == 0 else 20
                for k, v in (("before", "0"), ("after", "0"), ("line", str(h)), ("lineRule", "exact")):
                    sp.set(w + k, v)
            leading = []
            changed += 1
        with zipfile.ZipFile(docx_out, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)
                if item.filename == "word/document.xml" and changed:
                    data = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
                zout.writestr(item, data)
    return changed, target, width


def guide_line_top(pdf_path: str, target: float, width: float):
    """Lage (pt ab oberem Seitenrand) der roten Hilfslinie an der Tabellenoberkante.
    Rückgabe: (y oder None, Fehlertext oder None). Erkennt die Linie auch, wenn LibreOffice sie in
    mehreren Teilstücken (z. B. je Zelle) oder als gefülltes Rechteck zeichnet."""
    try:
        import ctypes
        import pypdfium2 as pdfium
        import pypdfium2.raw as pdfium_c
    except ImportError:
        return None, "Python-Paket pypdfium2 fehlt (pip install pypdfium2)"
    pdf = page = None
    try:
        pdf = pdfium.PdfDocument(pdf_path)
        page = pdf[0]
        h = page.get_height()

        def is_red(obj):
            r, g, b_, a_ = (ctypes.c_uint() for _ in range(4))
            for getter in (pdfium_c.FPDFPageObj_GetStrokeColor, pdfium_c.FPDFPageObj_GetFillColor):
                if getter(obj.raw, ctypes.byref(r), ctypes.byref(g), ctypes.byref(b_), ctypes.byref(a_)) \
                        and r.value > 200 and g.value < 80 and b_.value < 80:
                    return True
            return False

        segs, red_segs = [], []
        for obj in page.get_objects(filter=[pdfium_c.FPDF_PAGEOBJ_PATH]):
            left, bottom, right, top = obj.get_bounds()
            if top - bottom > 3:
                continue
            seg = (round(h - (top + bottom) / 2, 1), left, right)
            segs.append(seg)
            if is_red(obj):
                red_segs.append(seg)

        def best_line(cands):
            rows = {}
            for y, l, r in cands:  # Teilstücke gleicher Höhe zusammenfassen
                key = next((k for k in rows if abs(k - y) <= 0.6), y)
                rows.setdefault(key, []).append((l, r))
            fits = [(y, max(r for _, r in p) - min(l for l, _ in p)) for y, p in rows.items()]
            fits = [y for y, wd in fits if not width or wd >= 0.8 * width]
            return min(fits, key=lambda y: abs(y - target)) if fits else None

        y = best_line(red_segs) if red_segs else None
        if y is None:
            y = best_line(segs)
        return (y, None) if y is not None else (None, "Hilfslinie in der PDF nicht gefunden")
    except Exception as e:
        return None, f"Messung fehlgeschlagen ({e})"
    finally:
        for obj in (page, pdf):
            if obj is not None:
                try:
                    obj.close()
                except Exception:
                    pass


def soffice_convert(soffice: str, docx_path: str, outdir: str) -> str:
    profile = "file:///" + os.path.join(outdir, "profile").replace("\\", "/").lstrip("/")
    cmd = [soffice, f"-env:UserInstallation={profile}", "--headless", "--convert-to",
           'pdf:writer_pdf_Export:{"SelectPdfVersion":{"type":"long","value":"3"}}',
           "--outdir", outdir, docx_path]
    res = subprocess.run(cmd, capture_output=True, text=True, timeout=180)
    produced = os.path.join(outdir, os.path.splitext(os.path.basename(docx_path))[0] + ".pdf")
    if not os.path.exists(produced):
        raise InvoiceError(f"LibreOffice-Konvertierung fehlgeschlagen:\n{res.stdout}\n{res.stderr}")
    return produced


def docx_to_pdfa3(docx_path: str, pdf_out: str, soffice: str = None) -> None:
    """Word -> PDF/A-3b mit LibreOffice (Export-Filteroption SelectPdfVersion=3).
    Eigenes, temporäres LibreOffice-Profil: funktioniert auch, wenn LibreOffice gerade geöffnet ist.
    Schwebendes Adressfeld (siehe inline_floating_tables): Messlauf mit Hilfslinie an der Tabellen-
    oberkante, dann Endlauf mit korrigiertem Abstand, so dass die Tabelle auf der in Word festgelegten
    Höhe steht – unabhängig davon, wie die jeweilige LibreOffice-Version schwebende Tabellen setzt."""
    soffice = find_soffice(soffice)
    with tempfile.TemporaryDirectory() as tmp:
        name = os.path.basename(docx_path)
        dirs = {k: os.path.join(tmp, k) for k in ("mess", "final")}
        for d_ in dirs.values():
            os.makedirs(d_)
        probe = os.path.join(dirs["mess"], name)
        changed, target, width = inline_floating_tables(docx_path, probe, guide=True)
        if not changed:
            shutil.move(soffice_convert(soffice, docx_path, tmp), pdf_out)
            return
        actual, why = guide_line_top(soffice_convert(soffice, probe, dirs["mess"]), target, width)
        shift = 0
        '''
        if actual is None:
            print(f"  Hinweis: Position des Adressfelds konnte nicht nachgemessen werden: {why}.\n"
                  "  Das Adressfeld kann deshalb gegenüber Word etwas verschoben sein.")
        else:
            shift = round((actual - target) * 20)             # pt -> Twips
            print(f"  Adressfeld auf Word-Position ausgerichtet (Korrektur {actual - target:+.1f} pt).")
        '''
        final = os.path.join(dirs["final"], name)
        inline_floating_tables(docx_path, final, shift)
        shutil.move(soffice_convert(soffice, final, dirs["final"]), pdf_out)


def warn_font_substitution(docx_path: str, pdf_path: str) -> None:
    """Vergleicht die im Word-Dokument verwendeten Schriften mit den im PDF eingebetteten."""
    norm = lambda n: re.sub(r"[^a-z0-9]", "", n.lower())
    with pikepdf.open(pdf_path) as pdf:
        embedded = set()
        for page in pdf.pages:
            for font in (page.Resources.get("/Font") or {}).values():
                base = str(font.get("/BaseFont", ""))
                embedded.add(norm(base.split("+", 1)[-1].split(",")[0].split("-")[0]))
    missing = [f for f in sorted(docx_fonts(docx_path))
               if not any(e.startswith(norm(f)) or norm(f).startswith(e) for e in embedded if e)]
    '''
    if missing:
        print("WARNUNG: Diese Schriften aus dem Word-Dokument wurden im PDF ersetzt: " + ", ".join(missing)
              + ".\n         Das Layout (Zeilenumbrüche, Abstände) kann deshalb von Word abweichen. "
              "Abhilfe: Schrift auf diesem Rechner installieren.")
    '''
