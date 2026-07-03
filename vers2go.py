"""
Vers2Go — transformă versurile unei melodii într-o prezentare PowerPoint.

Fiecare strofă (separată prin linii goale) devine un slide cu fundal negru
și text alb, bold, centrat. Mărimea fontului este calculată explicit în
funcție de lungimea strofei, astfel încât textul să încapă pe slide fără
a depinde de auto-fit-ul din PowerPoint (rezultatul arată corect și în
LibreOffice Impress sau Google Slides).

Aplicația NU are dependențe externe: fișierul .pptx (care este de fapt o
arhivă ZIP cu documente XML în formatul Office Open XML) este generat
direct cu bibliotecile standard `zipfile` și `xml.sax.saxutils`. Astfel
funcționează și pe sisteme unde politicile de securitate blochează
DLL-urile native din pachetele pip (de ex. lxml).

Rulare în development:  python vers2go.py
Build executabil:       vezi README.md
"""

import os
import sys
import zipfile
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
from xml.sax.saxutils import escape


def resource_path(filename: str) -> str:
    """Calea către un fișier-resursă (de ex. iconița aplicației).

    În development, resursele stau lângă acest script. În executabilul
    PyInstaller --onefile, ele sunt despachetate la rulare într-un folder
    temporar a cărui cale e pusă în sys._MEIPASS.
    """
    base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, filename)

# ---------------------------------------------------------------------------
# Constante de design
# ---------------------------------------------------------------------------

EMU_PER_INCH = 914400  # unitatea de măsură OOXML (English Metric Units)

# Dimensiuni slide 16:9 (standardul actual pentru proiectoare / ecrane)
SLIDE_WIDTH_IN = 13.333
SLIDE_HEIGHT_IN = 7.5

# Marginea lăsată liberă pe fiecare latură a casetei de text
MARGIN_IN = 0.45

BACKGROUND_COLOR = "000000"  # negru (hex RGB)
TEXT_COLOR = "FFFFFF"        # alb  (hex RGB)

FONT_NAME = "Arial"

# Limitele mărimii fontului (în puncte)
MAX_FONT_PT = 66
MIN_FONT_PT = 14

# Factori empirici pentru estimarea spațiului ocupat de text:
# - un caracter de Arial Bold ocupă în medie ~0.62 din mărimea fontului în lățime
# - o linie de text ocupă ~1.25 din mărimea fontului în înălțime (cu spațiere)
CHAR_WIDTH_FACTOR = 0.62
LINE_HEIGHT_FACTOR = 1.25


# ---------------------------------------------------------------------------
# Logica de împărțire a textului și calcul al fontului
# ---------------------------------------------------------------------------

def split_into_stanzas(lyrics: str) -> list[str]:
    """Împarte textul în strofe folosind liniile goale ca separator.

    Elimină spațiile de la începutul/finalul fiecărei linii și strofe și
    ignoră strofele complet goale (de ex. mai multe linii goale la rând).
    """
    stanzas = []
    current_lines = []

    for raw_line in lyrics.splitlines():
        line = raw_line.strip()
        if line:
            current_lines.append(line)
        elif current_lines:
            # linie goală => se închide strofa curentă
            stanzas.append("\n".join(current_lines))
            current_lines = []

    # ultima strofă, dacă textul nu se termină cu linie goală
    if current_lines:
        stanzas.append("\n".join(current_lines))

    return stanzas


def _strip_chorus_marker(stanza: str) -> str | None:
    """Dacă strofa e marcată ca refren (prima linie începe cu „R:"),
    returnează textul refrenului fără marcaj; altfel returnează None."""
    lines = stanza.splitlines()
    if not lines or not lines[0].upper().startswith("R:"):
        return None

    # eliminăm marcajul „R:" de pe prima linie
    first = lines[0][2:].strip()
    rest = lines[1:]
    chorus_lines = ([first] if first else []) + rest
    return "\n".join(chorus_lines) if chorus_lines else None


def build_slide_texts(stanzas: list[str]) -> list[str]:
    """Construiește lista finală de texte pentru slide-uri, inserând refrenul.

    O strofă care începe cu „R:" este refrenul melodiei: nu apare ca slide
    de sine stătător la poziția ei, ci este pusă automat după fiecare
    strofă obișnuită (inclusiv după strofele dinaintea definiției, cazul
    tipic fiind: strofa 1, apoi refrenul marcat cu R:).

    Dacă melodia are mai multe refrene marcate (de ex. se schimbă refrenul
    spre final), fiecare strofă primește refrenul cel mai recent definit
    înaintea ei (sau primul refren, pentru strofele de dinainte).
    """
    # Separăm refrenele de strofele obișnuite, păstrând ordinea
    items = []  # liste de perechi (este_refren, text)
    for stanza in stanzas:
        chorus = _strip_chorus_marker(stanza)
        if chorus is not None:
            items.append((True, chorus))
        else:
            items.append((False, stanza))

    choruses = [text for is_chorus, text in items if is_chorus]
    if not choruses:
        # nicio strofă marcată cu R: => comportamentul obișnuit
        return [text for _, text in items]

    slide_texts = []
    current_chorus = choruses[0]
    for is_chorus, text in items:
        if is_chorus:
            current_chorus = text  # de aici înainte se aplică acest refren
        else:
            slide_texts.append(text)
            slide_texts.append(current_chorus)

    # Caz limită: textul conține DOAR refren(e), fără strofe obișnuite —
    # le afișăm ca slide-uri normale ca să nu producem o prezentare goală.
    if not slide_texts:
        slide_texts = choruses

    return slide_texts


def compute_font_size(stanza: str) -> int:
    """Calculează cea mai mare mărime de font (pt) la care strofa încape pe slide.

    Se aplică două constrângeri independente și se ia minimul dintre ele:
      1. Lățime:  cea mai lungă linie trebuie să încapă pe lățimea utilă.
      2. Înălțime: toate liniile trebuie să încapă pe înălțimea utilă.

    Estimarea folosește factori medii pentru Arial Bold, deci nu depinde de
    auto-fit-ul aplicației care deschide fișierul.
    """
    lines = stanza.splitlines()
    num_lines = max(len(lines), 1)
    longest_line = max((len(line) for line in lines), default=1)
    longest_line = max(longest_line, 1)

    # Spațiul util în puncte (1 inch = 72 pt)
    usable_width_pt = (SLIDE_WIDTH_IN - 2 * MARGIN_IN) * 72
    usable_height_pt = (SLIDE_HEIGHT_IN - 2 * MARGIN_IN) * 72

    # Constrângerea pe lățime: longest_line * CHAR_WIDTH_FACTOR * font <= usable_width
    font_by_width = usable_width_pt / (longest_line * CHAR_WIDTH_FACTOR)

    # Constrângerea pe înălțime: num_lines * LINE_HEIGHT_FACTOR * font <= usable_height
    font_by_height = usable_height_pt / (num_lines * LINE_HEIGHT_FACTOR)

    font_size = int(min(font_by_width, font_by_height))

    # Limităm la un interval rezonabil: strofele foarte scurte nu devin
    # absurd de mari, iar cele foarte lungi rămân lizibile.
    return max(MIN_FONT_PT, min(MAX_FONT_PT, font_size))


# ---------------------------------------------------------------------------
# Generarea fișierului .pptx (Office Open XML, fără biblioteci externe)
#
# Un .pptx este o arhivă ZIP cu o structură fixă de fișiere XML. Mai jos
# sunt părțile statice (schelet identic pentru orice prezentare) și
# funcțiile care construiesc părțile dinamice (slide-urile).
# ---------------------------------------------------------------------------

_XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'

# Namespace-urile OOXML folosite peste tot
_NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
_NS_P = "http://schemas.openxmlformats.org/presentationml/2006/main"
_NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_NS_REL = "http://schemas.openxmlformats.org/package/2006/relationships"

# Relațiile de la rădăcina pachetului: documentul principal + proprietăți
_ROOT_RELS = _XML_DECL + f"""<Relationships xmlns="{_NS_REL}">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>"""

_APP_XML = _XML_DECL + """<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
<Application>Vers2Go</Application>
</Properties>"""

# Un grup de forme gol — scheletul minim obligatoriu pentru orice cSld
_EMPTY_SPTREE = """<p:spTree>
<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>
</p:spTree>"""

# Fundalul negru, folosit și pe master și pe fiecare slide
_BLACK_BG = f"""<p:bg><p:bgPr><a:solidFill><a:srgbClr val="{BACKGROUND_COLOR}"/></a:solidFill><a:effectLst/></p:bgPr></p:bg>"""

# Master-ul: fundal negru, fără placeholder-e, cu layout-ul "blank" atașat
_SLIDE_MASTER = _XML_DECL + f"""<p:sldMaster xmlns:a="{_NS_A}" xmlns:r="{_NS_R}" xmlns:p="{_NS_P}">
<p:cSld>{_BLACK_BG}{_EMPTY_SPTREE}</p:cSld>
<p:clrMap bg1="dk1" tx1="lt1" bg2="dk2" tx2="lt2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>
<p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/></p:sldLayoutIdLst>
<p:txStyles><p:titleStyle/><p:bodyStyle/><p:otherStyle/></p:txStyles>
</p:sldMaster>"""

_SLIDE_MASTER_RELS = _XML_DECL + f"""<Relationships xmlns="{_NS_REL}">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="../theme/theme1.xml"/>
</Relationships>"""

_SLIDE_LAYOUT = _XML_DECL + f"""<p:sldLayout xmlns:a="{_NS_A}" xmlns:r="{_NS_R}" xmlns:p="{_NS_P}" type="blank" preserve="1">
<p:cSld name="Blank">{_EMPTY_SPTREE}</p:cSld>
<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
</p:sldLayout>"""

_SLIDE_LAYOUT_RELS = _XML_DECL + f"""<Relationships xmlns="{_NS_REL}">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/>
</Relationships>"""

# Tema: obligatorie în pachet; conține schemele minime de culori/fonturi/formate
_THEME = _XML_DECL + f"""<a:theme xmlns:a="{_NS_A}" name="Vers2Go">
<a:themeElements>
<a:clrScheme name="Vers2Go">
<a:dk1><a:srgbClr val="000000"/></a:dk1>
<a:lt1><a:srgbClr val="FFFFFF"/></a:lt1>
<a:dk2><a:srgbClr val="1F1F1F"/></a:dk2>
<a:lt2><a:srgbClr val="EEEEEE"/></a:lt2>
<a:accent1><a:srgbClr val="4472C4"/></a:accent1>
<a:accent2><a:srgbClr val="ED7D31"/></a:accent2>
<a:accent3><a:srgbClr val="A5A5A5"/></a:accent3>
<a:accent4><a:srgbClr val="FFC000"/></a:accent4>
<a:accent5><a:srgbClr val="5B9BD5"/></a:accent5>
<a:accent6><a:srgbClr val="70AD47"/></a:accent6>
<a:hlink><a:srgbClr val="0563C1"/></a:hlink>
<a:folHlink><a:srgbClr val="954F72"/></a:folHlink>
</a:clrScheme>
<a:fontScheme name="Vers2Go">
<a:majorFont><a:latin typeface="{FONT_NAME}"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont>
<a:minorFont><a:latin typeface="{FONT_NAME}"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont>
</a:fontScheme>
<a:fmtScheme name="Vers2Go">
<a:fillStyleLst>
<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
</a:fillStyleLst>
<a:lnStyleLst>
<a:ln w="9525"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>
<a:ln w="19050"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>
<a:ln w="28575"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>
</a:lnStyleLst>
<a:effectStyleLst>
<a:effectStyle><a:effectLst/></a:effectStyle>
<a:effectStyle><a:effectLst/></a:effectStyle>
<a:effectStyle><a:effectLst/></a:effectStyle>
</a:effectStyleLst>
<a:bgFillStyleLst>
<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
</a:bgFillStyleLst>
</a:fmtScheme>
</a:themeElements>
</a:theme>"""


def _content_types_xml(num_slides: int) -> str:
    """[Content_Types].xml — declară tipul MIME al fiecărei părți din pachet."""
    slide_overrides = "\n".join(
        f'<Override PartName="/ppt/slides/slide{i}.xml" '
        f'ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>'
        for i in range(1, num_slides + 1)
    )
    return _XML_DECL + f"""<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
<Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>
<Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>
<Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>
{slide_overrides}
<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
</Types>"""


def _presentation_xml(num_slides: int) -> str:
    """ppt/presentation.xml — lista slide-urilor și dimensiunea lor."""
    # rId1 = slide master; slide-urile încep de la rId2
    slide_ids = "\n".join(
        f'<p:sldId id="{255 + i}" r:id="rId{i + 1}"/>'
        for i in range(1, num_slides + 1)
    )
    cx = int(SLIDE_WIDTH_IN * EMU_PER_INCH)
    cy = int(SLIDE_HEIGHT_IN * EMU_PER_INCH)
    return _XML_DECL + f"""<p:presentation xmlns:a="{_NS_A}" xmlns:r="{_NS_R}" xmlns:p="{_NS_P}">
<p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst>
<p:sldIdLst>
{slide_ids}
</p:sldIdLst>
<p:sldSz cx="{cx}" cy="{cy}"/>
<p:notesSz cx="6858000" cy="9144000"/>
</p:presentation>"""


def _presentation_rels_xml(num_slides: int) -> str:
    """Relațiile prezentării: master-ul + fiecare slide."""
    slide_rels = "\n".join(
        f'<Relationship Id="rId{i + 1}" '
        f'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" '
        f'Target="slides/slide{i}.xml"/>'
        for i in range(1, num_slides + 1)
    )
    return _XML_DECL + f"""<Relationships xmlns="{_NS_REL}">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/>
{slide_rels}
</Relationships>"""


_SLIDE_RELS = _XML_DECL + f"""<Relationships xmlns="{_NS_REL}">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
</Relationships>"""


def _slide_xml(text: str, font_size_pt: int) -> str:
    """XML-ul unui slide: fundal negru + casetă de text centrată.

    Fiecare linie a strofei devine un paragraf centrat orizontal; caseta
    de text acoperă zona utilă a slide-ului și e ancorată la mijloc
    (centrare verticală). Mărimea fontului e scrisă explicit pe fiecare
    run (sz este în sutimi de punct), fără auto-fit (<a:noAutofit/>).
    """
    # Poziția și dimensiunea casetei de text, în EMU
    off = int(MARGIN_IN * EMU_PER_INCH)
    ext_cx = int((SLIDE_WIDTH_IN - 2 * MARGIN_IN) * EMU_PER_INCH)
    ext_cy = int((SLIDE_HEIGHT_IN - 2 * MARGIN_IN) * EMU_PER_INCH)
    sz = font_size_pt * 100  # OOXML exprimă mărimea fontului în 1/100 pt

    paragraphs = "\n".join(
        f'<a:p><a:pPr algn="ctr"/>'
        f'<a:r><a:rPr lang="ro-RO" sz="{sz}" b="1" dirty="0">'
        f'<a:solidFill><a:srgbClr val="{TEXT_COLOR}"/></a:solidFill>'
        f'<a:latin typeface="{FONT_NAME}"/><a:cs typeface="{FONT_NAME}"/></a:rPr>'
        f'<a:t>{escape(line)}</a:t></a:r></a:p>'
        for line in text.splitlines()
    )

    return _XML_DECL + f"""<p:sld xmlns:a="{_NS_A}" xmlns:r="{_NS_R}" xmlns:p="{_NS_P}">
<p:cSld>{_BLACK_BG}
<p:spTree>
<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>
<p:sp>
<p:nvSpPr><p:cNvPr id="2" name="Versuri"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>
<p:spPr>
<a:xfrm><a:off x="{off}" y="{off}"/><a:ext cx="{ext_cx}" cy="{ext_cy}"/></a:xfrm>
<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
<a:noFill/>
</p:spPr>
<p:txBody>
<a:bodyPr wrap="square" lIns="0" tIns="0" rIns="0" bIns="0" anchor="ctr"><a:noAutofit/></a:bodyPr>
<a:lstStyle/>
{paragraphs}
</p:txBody>
</p:sp>
</p:spTree>
</p:cSld>
<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
</p:sld>"""


def _core_xml(title: str) -> str:
    """docProps/core.xml — metadatele documentului (titlul, dacă există)."""
    title_el = f"<dc:title>{escape(title)}</dc:title>" if title else ""
    return _XML_DECL + f"""<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/">
{title_el}<dc:creator>Vers2Go</dc:creator>
</cp:coreProperties>"""


def generate_presentation(lyrics: str, title: str, output_path: str) -> int:
    """Generează fișierul .pptx și returnează numărul de slide-uri create.

    Aruncă ValueError dacă textul nu conține nicio strofă.
    """
    stanzas = split_into_stanzas(lyrics)
    if not stanzas:
        raise ValueError("Textul nu conține versuri.")

    # Slide de titlu (doar dacă utilizatorul a completat titlul) + strofele,
    # cu refrenul (strofa marcată „R:") inserat după fiecare strofă
    title = title.strip()
    slide_texts = ([title] if title else []) + build_slide_texts(stanzas)
    num_slides = len(slide_texts)

    # Scriem arhiva ZIP cu toate părțile pachetului OOXML
    with zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", _content_types_xml(num_slides))
        zf.writestr("_rels/.rels", _ROOT_RELS)
        zf.writestr("docProps/core.xml", _core_xml(title))
        zf.writestr("docProps/app.xml", _APP_XML)
        zf.writestr("ppt/presentation.xml", _presentation_xml(num_slides))
        zf.writestr("ppt/_rels/presentation.xml.rels", _presentation_rels_xml(num_slides))
        zf.writestr("ppt/slideMasters/slideMaster1.xml", _SLIDE_MASTER)
        zf.writestr("ppt/slideMasters/_rels/slideMaster1.xml.rels", _SLIDE_MASTER_RELS)
        zf.writestr("ppt/slideLayouts/slideLayout1.xml", _SLIDE_LAYOUT)
        zf.writestr("ppt/slideLayouts/_rels/slideLayout1.xml.rels", _SLIDE_LAYOUT_RELS)
        zf.writestr("ppt/theme/theme1.xml", _THEME)

        for i, text in enumerate(slide_texts, start=1):
            font_size = compute_font_size(text)
            zf.writestr(f"ppt/slides/slide{i}.xml", _slide_xml(text, font_size))
            zf.writestr(f"ppt/slides/_rels/slide{i}.xml.rels", _SLIDE_RELS)

    return num_slides


# ---------------------------------------------------------------------------
# Interfața grafică (tkinter)
# ---------------------------------------------------------------------------

class Vers2GoApp:
    """Fereastra principală a aplicației."""

    def __init__(self, root: tk.Tk):
        self.root = root
        root.title("Vers2Go — Versuri în prezentare")
        root.geometry("720x600")
        root.minsize(520, 400)

        # Iconița ferestrei: opțională — dacă lipsește sau formatul nu e
        # suportat (de ex. .ico pe Linux), aplicația pornește fără ea.
        try:
            root.iconbitmap(resource_path("icon.ico"))
        except tk.TclError:
            pass

        # --- Câmpul de titlu (opțional) ---
        title_frame = tk.Frame(root)
        title_frame.pack(fill="x", padx=12, pady=(12, 4))

        tk.Label(title_frame, text="Titlul melodiei (opțional):").pack(side="left")
        self.title_entry = tk.Entry(title_frame, font=("Segoe UI", 11))
        self.title_entry.pack(side="left", fill="x", expand=True, padx=(8, 0))

        # --- Zona de versuri ---
        tk.Label(
            root,
            text="Lipește versurile mai jos (strofe separate prin linii goale; "
                 "refrenul începe cu „R:” și se repetă după fiecare strofă):",
            anchor="w",
        ).pack(fill="x", padx=12, pady=(8, 2))

        self.lyrics_text = scrolledtext.ScrolledText(
            root, font=("Segoe UI", 11), wrap="word", undo=True
        )
        self.lyrics_text.pack(fill="both", expand=True, padx=12, pady=(0, 8))

        # --- Butonul de generare ---
        self.generate_button = tk.Button(
            root,
            text="Generează PowerPoint",
            font=("Segoe UI", 11, "bold"),
            command=self.on_generate,
            height=2,
        )
        self.generate_button.pack(fill="x", padx=12, pady=(0, 12))

    def on_generate(self):
        """Handler-ul butonului: validează, cere destinația și generează."""
        lyrics = self.lyrics_text.get("1.0", "end")
        title = self.title_entry.get()

        # Validare: textarea goală => mesaj de eroare, nu crash
        if not split_into_stanzas(lyrics):
            messagebox.showerror(
                "Text lipsă",
                "Lipește mai întâi versurile melodiei în caseta de text.",
            )
            return

        # Numele implicit al fișierului, derivat din titlu dacă există
        default_name = title.strip() or "prezentare_versuri"
        # eliminăm caracterele interzise în numele de fișiere pe Windows
        for ch in '\\/:*?"<>|':
            default_name = default_name.replace(ch, "")

        output_path = filedialog.asksaveasfilename(
            title="Salvează prezentarea",
            defaultextension=".pptx",
            initialfile=default_name + ".pptx",
            filetypes=[("Prezentare PowerPoint", "*.pptx")],
        )
        if not output_path:
            return  # utilizatorul a anulat dialogul

        try:
            num_slides = generate_presentation(lyrics, title, output_path)
        except PermissionError:
            messagebox.showerror(
                "Eroare la salvare",
                "Fișierul nu a putut fi salvat.\n"
                "Verifică dacă nu este deja deschis în PowerPoint.",
            )
            return
        except Exception as exc:  # orice altă eroare neprevăzută
            messagebox.showerror("Eroare", f"A apărut o eroare:\n{exc}")
            return

        messagebox.showinfo(
            "Gata!",
            f"Prezentarea a fost salvată ({num_slides} slide-uri):\n{output_path}",
        )


def main():
    root = tk.Tk()
    Vers2GoApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
