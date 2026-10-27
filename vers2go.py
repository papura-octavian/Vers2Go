"""
Vers2Go — transformă versurile unei cântări într-o prezentare PowerPoint.

Fiecare secțiune a cântării (strofă, refren, punte, final) devine un slide
cu fundal negru și text alb, bold, centrat. Secțiunile se marchează cu o
etichetă pe un rând separat (S1, R, B, C — vezi `parse_song`), iar ordinea
de cântare se poate scrie explicit (de ex. „S1 R S2 R B R”) sau e calculată
automat, cu refrenul repetat după fiecare strofă.

Mărimea fontului este calculată explicit în funcție de lungimea textului,
astfel încât acesta să încapă pe slide fără a depinde de auto-fit-ul din
PowerPoint (rezultatul arată corect și în LibreOffice Impress sau Google
Slides).

Aplicația NU are dependențe externe: fișierul .pptx (care este de fapt o
arhivă ZIP cu documente XML în formatul Office Open XML) este generat
direct cu bibliotecile standard `zipfile` și `xml.sax.saxutils`. Astfel
funcționează și pe sisteme unde politicile de securitate blochează
DLL-urile native din pachetele pip (de ex. lxml).

Rulare în development:  python vers2go.py
Build executabil:       vezi README.md
"""

import os
import re
import sys
import zipfile
import tkinter as tk
import tkinter.font as tkfont
import unicodedata
from dataclasses import dataclass, field
from tkinter import filedialog, messagebox, ttk
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

# Eticheta opțională a secțiunii („Refren”, „Strofa 2”), mică și gri, în
# colțul din stânga-sus. Când e afișată, rezervăm o bandă sus și jos (egală,
# ca textul să rămână centrat) ca versurile să nu se suprapună cu ea.
LABEL_COLOR = "8C8C8C"
LABEL_FONT_PT = 18
LABEL_SPACE_IN = 0.5

# Estimarea spațiului ocupat de text:
# - lățimea fiecărui caracter, în fracțiuni din mărimea fontului, după
#   metricile Arial Bold (literele cu diacritice au lățimea literei de bază);
#   majusculele care nu sunt în tabel au ~0.72, celelalte caractere ~0.56
# - o linie de text ocupă ~1.25 din mărimea fontului în înălțime (cu spațiere)
# - o marjă de siguranță de 6%, ca rândul cel mai lung să nu fie rupt în două
_CHAR_WIDTHS = {
    **dict.fromkeys("abcdeghknopqsuvxy", 0.58),
    **dict.fromkeys("ijl'", 0.28), **dict.fromkeys("ftr", 0.36),
    "m": 0.89, "w": 0.78, "z": 0.50,
    **dict.fromkeys(" .,;:!-", 0.28), "?": 0.61,
    "I": 0.28, "J": 0.56, "L": 0.61, "M": 0.83, "W": 0.94,
}
UPPER_WIDTH = 0.72
DEFAULT_CHAR_WIDTH = 0.56
WIDTH_SAFETY = 1.06
LINE_HEIGHT_FACTOR = 1.25


# ---------------------------------------------------------------------------
# Secțiunile cântării: marcaje, ordine de cântare
# ---------------------------------------------------------------------------

# Tipurile de secțiuni, după codul scurt folosit în marcaje și în ordine
SECTION_NAMES = {
    "S": "Strofa",
    "P": "Pre-refren",
    "R": "Refren",
    "B": "Punte",
    "C": "Final",
}

# Cuvintele acceptate ca etichetă (fără diacritice / cu diacritice, în română
# și în engleză), fiecare cu tipul de secțiune corespunzător
_KIND_WORDS = {
    "s": "S", "strofa": "S", "strofă": "S", "strofe": "S", "v": "S", "verse": "S",
    "r": "R", "refren": "R", "refrenul": "R", "chorus": "R",
    "p": "P", "pre": "P", "prerefren": "P", "pre-refren": "P",
    "prechorus": "P", "pre-chorus": "P",
    "b": "B", "punte": "B", "puntea": "B", "bridge": "B",
    "c": "C", "coda": "C", "final": "C", "finalul": "C", "ending": "C",
    "outro": "C",
}

# Eticheta: un cuvânt, un număr opțional, eventual între paranteze pătrate,
# urmată opțional de „:” și de primul vers (ca la vechiul „R: text”)
_HEADER_RE = re.compile(
    r"^\[?\s*(?P<word>[^\W\d_](?:[^\W\d_]|-)*)\s*(?P<num>\d+)?\s*\]?\s*"
    r"(?P<sep>[:.)])?\s*(?P<rest>.*)$"
)

# Strofă numerotată ca în cântările tipărite: „1. Primul vers...”
_NUMBERED_RE = re.compile(r"^(?P<num>\d+)[.)]\s*(?P<rest>.*)$")


@dataclass
class Section:
    """O secțiune a cântării (strofă, refren etc.) cu versurile ei."""

    kind: str                # una din cheile SECTION_NAMES
    number: int | None       # S are mereu număr; R/B/C/P doar dacă sunt mai multe
    lines: list[str] = field(default_factory=list)

    @property
    def code(self) -> str:
        """Codul scurt: S1, R, R2, B, C."""
        return self.kind + (str(self.number) if self.number else "")

    @property
    def name(self) -> str:
        """Numele afișat: „Strofa 1”, „Refren”, „Refren 2”, „Punte”."""
        base = SECTION_NAMES[self.kind]
        return f"{base} {self.number}" if self.number else base

    @property
    def key(self) -> tuple[str, int]:
        """Cheia de identificare: R și R1 sunt același refren."""
        return (self.kind, self.number or 1)


@dataclass
class Song:
    """Rezultatul citirii versurilor."""

    sections: list[Section]   # secțiunile definite, fiecare o singură dată
    sequence: list[Section]   # ordinea în care apar în text (cu repetări)
    warnings: list[str]


def parse_header(line: str, at_block_start: bool = False):
    """Recunoaște o etichetă de secțiune pe un rând.

    Returnează (tip, număr sau None, text rămas pe rând) sau None dacă rândul
    e un vers obișnuit. Exemple: „S1”, „Strofa 2”, „R”, „Refren:”, „[Punte]”,
    „R: primul vers al refrenului”. Un rând care începe cu eticheta dar
    continuă cu text fără „:” (de ex. „Punte de lumină”) rămâne vers.

    La începutul unui bloc se acceptă și numerotarea din cântările tipărite:
    „2. Primul vers al strofei a doua”.
    """
    text = line.strip()
    m = _HEADER_RE.match(text)
    if m and m["word"].lower() in _KIND_WORDS:
        rest = m["rest"].strip()
        if not rest or m["sep"] == ":":
            num = int(m["num"]) if m["num"] else None
            return _KIND_WORDS[m["word"].lower()], num, rest

    if at_block_start:
        m = _NUMBERED_RE.match(text)
        if m:
            return "S", int(m["num"]), m["rest"].strip()
    return None


def parse_song(lyrics: str) -> Song:
    """Împarte versurile în secțiuni.

    Reguli (gândite să meargă și pe versuri lipite așa cum sunt):
    - un rând cu eticheta (S1, R, B, C...) începe o secțiune nouă;
    - un rând gol închide secțiunea curentă;
    - un bloc fără etichetă e strofa următoare (numerotată automat), iar dacă
      textul lui e identic cu al unei secțiuni de dinainte, e o repetare a ei
      (de ex. refrenul scris de mai multe ori);
    - o etichetă fără versuri sub ea (de ex. doar „R”) înseamnă „aici se
      repetă secțiunea”, definită înainte sau după.
    """
    # 1. Blocuri brute: (tip, număr, linii); tip None = fără etichetă
    blocks = []
    current = None
    at_block_start = True

    for raw_line in lyrics.splitlines():
        line = raw_line.strip()
        if not line:
            current = None
            at_block_start = True
            continue

        header = parse_header(line, at_block_start)
        at_block_start = False
        if header:
            kind, num, rest = header
            current = [kind, num, [rest] if rest else []]
            blocks.append(current)
        elif current is None:
            current = [None, None, [line]]
            blocks.append(current)
        else:
            current[2].append(line)

    # 2. Secțiunile definite și ordinea din text
    warnings = []
    defined: dict[tuple[str, int], Section] = {}
    sequence: list[Section | tuple[str, int]] = []  # cheie = repetare de rezolvat
    used_verse_numbers = {num for kind, num, _ in blocks if kind == "S" and num}
    next_verse = 1

    def find_by_text(lines):
        for sec in defined.values():
            if sec.lines == lines:
                return sec
        return None

    for kind, num, lines in blocks:
        if kind is None:
            same = find_by_text(lines)
            if same:
                sequence.append(same)
                continue
            while next_verse in used_verse_numbers:
                next_verse += 1
            kind, num = "S", next_verse
            used_verse_numbers.add(num)

        if kind == "S" and num is None:
            # „Strofa” fără număr: primește următorul număr liber
            while next_verse in used_verse_numbers:
                next_verse += 1
            num = next_verse
            used_verse_numbers.add(num)

        key = (kind, num or 1)
        if not lines:
            sequence.append(defined.get(key, key))
            continue

        existing = defined.get(key)
        if existing is None:
            section = Section(kind, num, lines)
            defined[key] = section
            sequence.append(section)
        elif existing.lines == lines:
            sequence.append(existing)
        else:
            # aceeași etichetă, alt text: păstrăm ambele în ordinea scrisă,
            # dar în „Ordinea de cântare” codul se referă la primul
            warnings.append(
                f"„{existing.code}” apare de două ori cu versuri diferite. "
                f"Dă-i celei de-a doua alt nume (de ex. {kind}{(num or 1) + 1})."
            )
            sequence.append(Section(kind, num, lines))

    # 3. Rezolvăm repetările care se referă la secțiuni definite mai jos
    resolved = []
    for item in sequence:
        if isinstance(item, tuple):
            if item in defined:
                resolved.append(defined[item])
            else:
                code = item[0] + (str(item[1]) if item[0] == "S" or item[1] > 1 else "")
                warnings.append(f"„{code}” nu are versuri nicăieri în text, am sărit-o.")
        else:
            resolved.append(item)

    return Song(list(defined.values()), resolved, warnings)


def auto_order(song: Song) -> list[Section]:
    """Ordinea de cântare implicită.

    Secțiunile se cântă în ordinea în care sunt scrise. Dacă refrenul e
    scris o singură dată, el se repetă automat după fiecare strofă și după
    punte (cum se cântă de obicei: S1 R S2 R B R). Dacă refrenul e scris sau
    marcat de mai multe ori, înseamnă că ordinea e deja cea dorită și nu mai
    adăugăm nimic.

    Când sunt mai multe refrene (R1, R2), fiecare strofă primește refrenul
    cel mai recent dinaintea ei (sau primul, pentru strofele de dinainte).
    """
    seq = song.sequence
    choruses = [sec for sec in seq if sec.kind == "R"]
    if not choruses or len(choruses) != len({id(sec) for sec in choruses}):
        return list(seq)

    result = []
    current_chorus = choruses[0]
    for i, sec in enumerate(seq):
        if sec.kind == "R":
            current_chorus = sec
        result.append(sec)

        nxt = seq[i + 1] if i + 1 < len(seq) else None
        if sec.kind in ("S", "B"):
            # după strofă/punte vine refrenul, dacă nu urmează deja
            # refrenul sau pre-refrenul care duce spre el
            if nxt is None or nxt.kind not in ("R", "P"):
                result.append(current_chorus)
        elif sec.kind == "P" and (nxt is None or nxt.kind != "R"):
            result.append(current_chorus)
    return result


def order_codes(sections: list[Section]) -> str:
    """Ordinea ca text: „S1 R S2 R B R”."""
    return " ".join(sec.code for sec in sections)


# Element din ordine: „S1”, „R”, „Refren2”, „Rx2” (de două ori)
_ORDER_TOKEN_RE = re.compile(r"^(?P<word>[^\d×*]+?)(?P<num>\d+)?(?:[x×*](?P<times>\d+))?$")


def parse_order(order_text: str, song: Song) -> tuple[list[Section], list[str]]:
    """Citește ordinea de cântare scrisă de utilizator.

    Acceptă coduri separate prin spații sau virgule: „S1 R S2 R B R R”,
    cu nume întregi („Strofa 1 Refren Punte”) și repetări („Rx2”).
    Returnează (secțiunile în ordine, erori).
    """
    defined = {sec.key: sec for sec in song.sections}
    # „Strofa 2” -> „Strofa2”, ca numărul să rămână lângă cuvânt
    text = re.sub(r"([^\W\d_])\s+(\d)", r"\1\2", order_text)
    tokens = [t for t in re.split(r"[\s,;→>|]+", text) if t]

    result, errors = [], []
    for token in tokens:
        m = _ORDER_TOKEN_RE.match(token)
        kind = _KIND_WORDS.get(m["word"].lower()) if m else None
        if kind is None:
            errors.append(f"Nu înțeleg „{token}”.")
            continue
        if kind == "S" and not m["num"]:
            errors.append(f"La „{token}” scrie și numărul strofei (S1, S2...).")
            continue
        key = (kind, int(m["num"]) if m["num"] else 1)
        if key not in defined:
            errors.append(f"„{token}” nu există în versuri.")
            continue
        result.extend([defined[key]] * int(m["times"] or 1))
    return result, errors


@dataclass
class Slide:
    """Conținutul unui slide: textul și eticheta secțiunii (goală la titlu)."""

    text: str
    label: str = ""
    kind: str = ""
    code: str = ""


@dataclass
class Plan:
    """Tot ce trebuie pentru generare și previzualizare."""

    slides: list[Slide]
    song: Song
    auto: list[Section]
    errors: list[str]


def build_plan(lyrics: str, title: str = "", order_text: str = "") -> Plan:
    """Calculează slide-urile: titlul (opțional) + secțiunile în ordinea de
    cântare (cea scrisă de utilizator, sau cea automată dacă e goală)."""
    song = parse_song(lyrics)
    auto = auto_order(song)

    errors = []
    if order_text.strip():
        order, errors = parse_order(order_text, song)
    else:
        order = auto

    slides = [Slide(title.strip())] if title.strip() else []
    slides += [Slide("\n".join(sec.lines), sec.name, sec.kind, sec.code) for sec in order]
    return Plan(slides, song, auto, errors)


def line_width_em(line: str) -> float:
    """Lățimea estimată a unui rând, în multipli ai mărimii fontului."""
    total = 0.0
    for ch in line:
        base = unicodedata.normalize("NFD", ch)[0]  # „ș” -> „s”
        if base in _CHAR_WIDTHS:
            total += _CHAR_WIDTHS[base]
        else:
            total += UPPER_WIDTH if base.isupper() else DEFAULT_CHAR_WIDTH
    return total


def compute_font_size(text: str, reserved_in: float = 0.0) -> int:
    """Calculează cea mai mare mărime de font (pt) la care textul încape pe slide.

    Se aplică două constrângeri independente și se ia minimul dintre ele:
      1. Lățime:  cea mai lungă linie trebuie să încapă pe lățimea utilă.
      2. Înălțime: toate liniile trebuie să încapă pe înălțimea utilă
         (minus `reserved_in` inch, spațiul ocupat de etichetă).

    Estimarea folosește lățimile literelor din Arial Bold, deci nu depinde
    de auto-fit-ul aplicației care deschide fișierul.
    """
    lines = text.splitlines()
    num_lines = max(len(lines), 1)
    widest_em = max((line_width_em(line) for line in lines), default=1.0)
    widest_em = max(widest_em, 1.0)

    # Spațiul util în puncte (1 inch = 72 pt)
    usable_width_pt = (SLIDE_WIDTH_IN - 2 * MARGIN_IN) * 72
    usable_height_pt = (SLIDE_HEIGHT_IN - 2 * MARGIN_IN - reserved_in) * 72

    # Constrângerea pe lățime: widest_em * WIDTH_SAFETY * font <= usable_width
    font_by_width = usable_width_pt / (widest_em * WIDTH_SAFETY)

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


def _label_shape_xml(label: str) -> str:
    """Caseta cu eticheta secțiunii, în colțul din stânga-sus."""
    off = int(MARGIN_IN * EMU_PER_INCH)
    off_y = int(0.3 * EMU_PER_INCH)
    cx = int(5 * EMU_PER_INCH)
    cy = int(LABEL_SPACE_IN * EMU_PER_INCH)
    return f"""<p:sp>
<p:nvSpPr><p:cNvPr id="3" name="Eticheta"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>
<p:spPr>
<a:xfrm><a:off x="{off}" y="{off_y}"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>
<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
<a:noFill/>
</p:spPr>
<p:txBody>
<a:bodyPr wrap="none" lIns="0" tIns="0" rIns="0" bIns="0" anchor="t"><a:noAutofit/></a:bodyPr>
<a:lstStyle/>
<a:p><a:pPr algn="l"/><a:r><a:rPr lang="ro-RO" sz="{LABEL_FONT_PT * 100}" dirty="0">
<a:solidFill><a:srgbClr val="{LABEL_COLOR}"/></a:solidFill>
<a:latin typeface="{FONT_NAME}"/><a:cs typeface="{FONT_NAME}"/></a:rPr>
<a:t>{escape(label)}</a:t></a:r></a:p>
</p:txBody>
</p:sp>"""


def _slide_xml(text: str, font_size_pt: int, label: str = "") -> str:
    """XML-ul unui slide: fundal negru + casetă de text centrată.

    Fiecare linie a strofei devine un paragraf centrat orizontal; caseta
    de text acoperă zona utilă a slide-ului și e ancorată la mijloc
    (centrare verticală). Mărimea fontului e scrisă explicit pe fiecare
    run (sz este în sutimi de punct), fără auto-fit (<a:noAutofit/>).

    Dacă `label` e completat, eticheta secțiunii apare mică, în colț, iar
    caseta versurilor se micșorează egal sus și jos ca să nu o acopere.
    """
    # Poziția și dimensiunea casetei de text, în EMU
    band_in = LABEL_SPACE_IN if label else 0.0
    off_x = int(MARGIN_IN * EMU_PER_INCH)
    off_y = int((MARGIN_IN + band_in) * EMU_PER_INCH)
    ext_cx = int((SLIDE_WIDTH_IN - 2 * MARGIN_IN) * EMU_PER_INCH)
    ext_cy = int((SLIDE_HEIGHT_IN - 2 * MARGIN_IN - 2 * band_in) * EMU_PER_INCH)
    sz = font_size_pt * 100  # OOXML exprimă mărimea fontului în 1/100 pt

    paragraphs = "\n".join(
        f'<a:p><a:pPr algn="ctr"/>'
        f'<a:r><a:rPr lang="ro-RO" sz="{sz}" b="1" dirty="0">'
        f'<a:solidFill><a:srgbClr val="{TEXT_COLOR}"/></a:solidFill>'
        f'<a:latin typeface="{FONT_NAME}"/><a:cs typeface="{FONT_NAME}"/></a:rPr>'
        f'<a:t>{escape(line)}</a:t></a:r></a:p>'
        for line in text.splitlines()
    )
    label_shape = _label_shape_xml(label) if label else ""

    return _XML_DECL + f"""<p:sld xmlns:a="{_NS_A}" xmlns:r="{_NS_R}" xmlns:p="{_NS_P}">
<p:cSld>{_BLACK_BG}
<p:spTree>
<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>
<p:sp>
<p:nvSpPr><p:cNvPr id="2" name="Versuri"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>
<p:spPr>
<a:xfrm><a:off x="{off_x}" y="{off_y}"/><a:ext cx="{ext_cx}" cy="{ext_cy}"/></a:xfrm>
<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
<a:noFill/>
</p:spPr>
<p:txBody>
<a:bodyPr wrap="square" lIns="0" tIns="0" rIns="0" bIns="0" anchor="ctr"><a:noAutofit/></a:bodyPr>
<a:lstStyle/>
{paragraphs}
</p:txBody>
</p:sp>
{label_shape}
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


def write_pptx(slides: list[Slide], title: str, output_path: str,
               show_labels: bool = False) -> int:
    """Scrie fișierul .pptx cu slide-urile date și returnează numărul lor."""
    num_slides = len(slides)

    # Scriem arhiva ZIP cu toate părțile pachetului OOXML
    with zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", _content_types_xml(num_slides))
        zf.writestr("_rels/.rels", _ROOT_RELS)
        zf.writestr("docProps/core.xml", _core_xml(title.strip()))
        zf.writestr("docProps/app.xml", _APP_XML)
        zf.writestr("ppt/presentation.xml", _presentation_xml(num_slides))
        zf.writestr("ppt/_rels/presentation.xml.rels", _presentation_rels_xml(num_slides))
        zf.writestr("ppt/slideMasters/slideMaster1.xml", _SLIDE_MASTER)
        zf.writestr("ppt/slideMasters/_rels/slideMaster1.xml.rels", _SLIDE_MASTER_RELS)
        zf.writestr("ppt/slideLayouts/slideLayout1.xml", _SLIDE_LAYOUT)
        zf.writestr("ppt/slideLayouts/_rels/slideLayout1.xml.rels", _SLIDE_LAYOUT_RELS)
        zf.writestr("ppt/theme/theme1.xml", _THEME)

        for i, slide in enumerate(slides, start=1):
            label = slide.label if show_labels else ""
            font_size = compute_font_size(slide.text, 2 * LABEL_SPACE_IN if label else 0)
            zf.writestr(f"ppt/slides/slide{i}.xml", _slide_xml(slide.text, font_size, label))
            zf.writestr(f"ppt/slides/_rels/slide{i}.xml.rels", _SLIDE_RELS)

    return num_slides


def generate_presentation(lyrics: str, title: str, output_path: str,
                          order_text: str = "", show_labels: bool = False) -> int:
    """Generează fișierul .pptx și returnează numărul de slide-uri create.

    Aruncă ValueError dacă textul nu conține versuri sau dacă ordinea de
    cântare are greșeli.
    """
    plan = build_plan(lyrics, title, order_text)
    if not plan.song.sections:
        raise ValueError("Textul nu conține versuri.")
    if plan.errors:
        raise ValueError("\n".join(plan.errors))
    return write_pptx(plan.slides, title, output_path, show_labels)


# ---------------------------------------------------------------------------
# Interfața grafică (tkinter)
# ---------------------------------------------------------------------------

# Paleta interfeței (întunecată, sobră, cu un singur accent)
UI_BG = "#0F1115"
UI_CARD = "#181B22"
UI_FIELD = "#12141A"        # fundalul casetelor de text
UI_BORDER = "#2A2F3A"
UI_BUTTON = "#232733"
UI_BUTTON_HOVER = "#2E3342"
UI_TEXT = "#E6E8EE"
UI_MUTED = "#9097A6"
UI_ACCENT = "#6366F1"
UI_ACCENT_HOVER = "#4F46E5"
UI_SELECTION = "#3B3F6B"
UI_ERROR = "#F87171"
UI_WARN = "#FBBF24"
UI_FONT = "Segoe UI"

# Culorile fiecărui tip de secțiune: (text, fundal) — în editor și în
# previzualizare, ca strofele, refrenul și puntea să se distingă dintr-o privire
KIND_COLORS = {
    "S": ("#A5B4FC", "#252A4A"),
    "R": ("#6EE7B7", "#173A2F"),
    "P": ("#67E8F9", "#123A42"),
    "B": ("#FCD34D", "#3D3115"),
    "C": ("#D8B4FE", "#34234A"),
}

HELP_TEXT = """\
Cum marchezi secțiunile cântării

Scrie eticheta pe un rând separat, deasupra versurilor:

  S1, S2, S3...   Strofa — versurile care se schimbă de fiecare dată
  R               Refrenul — partea care se repetă identic
  R2              Al doilea refren, dacă cântarea are două diferite
  P               Pre-refren — 1-2 rânduri care duc spre refren (rar)
  B               Puntea (bridge) — o parte diferită, de obicei o dată, spre final
  C               Finalul (coda) — rândurile de încheiere

Merg și cuvinte întregi: „Strofa 2”, „Refren”, „Punte”, „Final”.
Strofele fără etichetă sunt numerotate singure (S1, S2...), deci
de obicei e destul să marchezi refrenul cu R.

Ordinea de cântare

Dacă lași câmpul gol, aplicația cântă secțiunile în ordinea scrisă
și repetă refrenul după fiecare strofă și după punte:
  S1 R S2 R B R

Ca să alegi tu ordinea, scrie codurile în câmp, de exemplu:
  S1 R S2 R B R R      sau      S1 R S2 Rx2 C
„Rx2” înseamnă refrenul de două ori.

Eticheta pe slide

Bifa „Arată numele secțiunii pe slide” scrie, mic și gri, în colțul
fiecărui slide, „Strofa 1”, „Refren” etc. Ajută echipa de proiecție
să se orienteze; pentru public slide-urile arată mai curat fără ea.
"""

SAMPLE_LYRICS = """\
S1
Primul vers al strofei întâi
Al doilea vers al strofei întâi

R
Primul vers al refrenului
Al doilea vers al refrenului

S2
Primul vers al strofei a doua
Al doilea vers al strofei a doua
"""


def _enable_dpi_awareness():
    """Pe Windows, fără asta fereastra e randată la 96 DPI și mărită de
    sistem — textul arată încețoșat pe ecranele cu scalare 125%/150%."""
    if sys.platform != "win32":
        return
    try:
        import ctypes
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except (AttributeError, OSError):
        pass


def _dark_title_bar(window: tk.Misc):
    """Bara de titlu a ferestrei devine întunecată, ca restul interfeței
    (Windows 10 20H1+ / Windows 11; pe alte sisteme nu face nimic)."""
    if sys.platform != "win32":
        return
    try:
        import ctypes
        window.update_idletasks()
        hwnd = ctypes.windll.user32.GetParent(window.winfo_id())
        value = ctypes.c_int(1)
        # DWMWA_USE_IMMERSIVE_DARK_MODE = 20
        ctypes.windll.dwmapi.DwmSetWindowAttribute(
            hwnd, 20, ctypes.byref(value), ctypes.sizeof(value))
    except (AttributeError, OSError):
        pass


class Card(tk.Frame):
    """Un panou cu margine subțire, în care stau grupurile de controale."""

    def __init__(self, parent, **kwargs):
        super().__init__(parent, bg=UI_CARD, highlightthickness=1,
                         highlightbackground=UI_BORDER, **kwargs)


class Chip(tk.Label):
    """Buton mic, colorat, pentru inserarea unei etichete de secțiune."""

    def __init__(self, parent, text, kind, command):
        fg, bg = KIND_COLORS[kind]
        super().__init__(parent, text=text, fg=fg, bg=bg, cursor="hand2",
                         font=(UI_FONT, 9, "bold"), padx=10, pady=3)
        self.bind("<Button-1>", lambda _e: command())
        self.bind("<Enter>", lambda _e: self.configure(bg=fg, fg=bg))
        self.bind("<Leave>", lambda _e: self.configure(bg=bg, fg=fg))


class Vers2GoApp:
    """Fereastra principală a aplicației."""

    THUMB_GAP = 12  # spațiul dintre miniaturile slide-urilor (px)

    def __init__(self, root: tk.Tk):
        self.root = root
        root.title("Vers2Go — Versuri în prezentare")
        root.configure(bg=UI_BG)
        _dark_title_bar(root)

        # La DPI mare, distanțele în pixeli trebuie mărite proporțional
        self.scale = max(1.0, root.winfo_fpixels("1i") / 96)
        root.geometry(f"{self.px(1180)}x{self.px(760)}")
        root.minsize(self.px(900), self.px(600))

        # Iconița ferestrei: opțională — dacă lipsește sau formatul nu e
        # suportat (de ex. .ico pe Linux), aplicația pornește fără ea.
        try:
            root.iconbitmap(resource_path("icon.ico"))
        except tk.TclError:
            pass

        self._setup_styles()
        self._thumb_fonts: dict[int, tkfont.Font] = {}
        self._refresh_job = None
        self.plan: Plan | None = None

        self._build_header()
        body = tk.Frame(root, bg=UI_BG)
        body.pack(fill="both", expand=True, padx=self.px(20), pady=(0, self.px(20)))
        body.columnconfigure(0, weight=5, uniform="col")
        body.columnconfigure(1, weight=4, uniform="col")
        body.rowconfigure(0, weight=1)

        self._build_editor(body).grid(row=0, column=0, sticky="nsew", padx=(0, self.px(10)))
        self._build_presentation(body).grid(row=0, column=1, sticky="nsew", padx=(self.px(10), 0))

        self.lyrics_text.insert("1.0", SAMPLE_LYRICS)
        self.lyrics_text.edit_reset()
        self.refresh()

    def px(self, value: float) -> int:
        """Pixeli scalați după DPI-ul ecranului."""
        return int(round(value * self.scale))

    # --- Construirea interfeței ------------------------------------------

    def _setup_styles(self):
        style = ttk.Style(self.root)
        style.theme_use("clam")
        base = (UI_FONT, 10)
        style.configure(".", background=UI_CARD, foreground=UI_TEXT, font=base)
        style.configure("TEntry", fieldbackground=UI_FIELD, foreground=UI_TEXT,
                        insertcolor=UI_TEXT, bordercolor=UI_BORDER,
                        lightcolor=UI_BORDER, darkcolor=UI_BORDER,
                        selectbackground=UI_SELECTION, selectforeground=UI_TEXT,
                        padding=self.px(7))
        style.map("TEntry", bordercolor=[("focus", UI_ACCENT)],
                  lightcolor=[("focus", UI_ACCENT)])

        style.configure("Accent.TButton", background=UI_ACCENT, foreground="#FFFFFF",
                        bordercolor=UI_ACCENT, lightcolor=UI_ACCENT, darkcolor=UI_ACCENT,
                        focuscolor=UI_ACCENT, font=(UI_FONT, 11, "bold"),
                        padding=(self.px(22), self.px(10)))
        style.map("Accent.TButton",
                  background=[("pressed", UI_ACCENT_HOVER), ("active", UI_ACCENT_HOVER)],
                  bordercolor=[("active", UI_ACCENT_HOVER)])

        style.configure("TButton", background=UI_BUTTON, foreground=UI_TEXT,
                        bordercolor=UI_BORDER, lightcolor=UI_BUTTON, darkcolor=UI_BUTTON,
                        focuscolor=UI_BUTTON, padding=(self.px(12), self.px(6)))
        style.map("TButton", background=[("pressed", UI_BUTTON_HOVER),
                                         ("active", UI_BUTTON_HOVER)],
                  lightcolor=[("active", UI_BUTTON_HOVER)],
                  darkcolor=[("active", UI_BUTTON_HOVER)])

        style.configure("TCheckbutton", background=UI_CARD, foreground=UI_TEXT,
                        focuscolor=UI_CARD, indicatorbackground=UI_FIELD,
                        indicatorforeground=UI_TEXT, bordercolor=UI_BORDER,
                        lightcolor=UI_FIELD, darkcolor=UI_FIELD)
        style.map("TCheckbutton", background=[("active", UI_CARD)],
                  indicatorbackground=[("selected", UI_ACCENT), ("active", UI_BUTTON)])

        style.configure("Vertical.TScrollbar", background=UI_BUTTON, troughcolor=UI_FIELD,
                        bordercolor=UI_FIELD, lightcolor=UI_BUTTON, darkcolor=UI_BUTTON,
                        arrowcolor=UI_MUTED, arrowsize=self.px(12))
        style.map("Vertical.TScrollbar", background=[("active", UI_BUTTON_HOVER)])

    def _label(self, parent, text, size=10, bold=False, color=UI_TEXT, bg=UI_CARD, **kw):
        font = (UI_FONT, size, "bold") if bold else (UI_FONT, size)
        return tk.Label(parent, text=text, font=font, fg=color, bg=bg, anchor="w",
                        justify="left", **kw)

    def _build_header(self):
        header = tk.Frame(self.root, bg=UI_BG)
        header.pack(fill="x", padx=self.px(20), pady=(self.px(16), self.px(14)))

        titles = tk.Frame(header, bg=UI_BG)
        titles.pack(side="left")
        self._label(titles, "Vers2Go", 18, True, bg=UI_BG).pack(anchor="w")
        self._label(titles, "Versurile cântării, gata de proiectat — câte o secțiune pe slide",
                    10, color=UI_MUTED, bg=UI_BG).pack(anchor="w")

        ttk.Button(header, text="Cum marchez secțiunile?",
                   command=self.show_help).pack(side="right")

    def _build_editor(self, parent) -> tk.Frame:
        card = Card(parent)
        inner = tk.Frame(card, bg=UI_CARD)
        inner.pack(fill="both", expand=True, padx=self.px(18), pady=self.px(16))

        self._label(inner, "Titlul cântării", 10, True).pack(fill="x")
        self._label(inner, "Opțional — devine primul slide.", 9, color=UI_MUTED).pack(fill="x")
        self.title_entry = ttk.Entry(inner, font=(UI_FONT, 11))
        self.title_entry.pack(fill="x", pady=(self.px(6), self.px(16)))
        self.title_entry.bind("<KeyRelease>", lambda _e: self.schedule_refresh())

        top = tk.Frame(inner, bg=UI_CARD)
        top.pack(fill="x")
        self._label(top, "Versuri", 10, True).pack(side="left")
        ttk.Button(top, text="Golește", command=self.clear_lyrics).pack(side="right")

        chips = tk.Frame(inner, bg=UI_CARD)
        chips.pack(fill="x", pady=(self.px(8), self.px(8)))
        self._label(chips, "Adaugă:", 9, color=UI_MUTED).pack(side="left", padx=(0, self.px(4)))
        for text, kind, command in (
            ("+ Strofă", "S", self.insert_verse),
            ("+ Refren", "R", lambda: self.insert_label("R")),
            ("+ Alt refren", "R", self.insert_new_chorus),
            ("+ Punte", "B", lambda: self.insert_label("B")),
            ("+ Final", "C", lambda: self.insert_label("C")),
        ):
            Chip(chips, text, kind, command).pack(side="left", padx=(self.px(4), 0))

        text_frame = tk.Frame(inner, bg=UI_FIELD, highlightthickness=1,
                              highlightbackground=UI_BORDER, highlightcolor=UI_ACCENT)
        text_frame.pack(fill="both", expand=True)
        self.lyrics_text = tk.Text(
            text_frame, font=(UI_FONT, 11), wrap="word", undo=True, relief="flat", bd=0,
            padx=self.px(14), pady=self.px(12), fg=UI_TEXT, bg=UI_FIELD,
            insertbackground=UI_TEXT, selectbackground=UI_SELECTION,
            selectforeground=UI_TEXT, spacing1=self.px(1), spacing3=self.px(1),
            highlightthickness=0,
        )
        scroll = ttk.Scrollbar(text_frame, orient="vertical", command=self.lyrics_text.yview)
        self.lyrics_text.configure(yscrollcommand=scroll.set)
        scroll.pack(side="right", fill="y")
        self.lyrics_text.pack(side="left", fill="both", expand=True)
        # bordura casetei devine colorată când scrii în ea
        self.lyrics_text.bind("<FocusIn>", lambda _e: text_frame.configure(highlightbackground=UI_ACCENT))
        self.lyrics_text.bind("<FocusOut>", lambda _e: text_frame.configure(highlightbackground=UI_BORDER))
        self.lyrics_text.bind("<<Modified>>", self._on_text_modified)

        for kind, (fg, bg) in KIND_COLORS.items():
            self.lyrics_text.tag_configure(f"hdr_{kind}", foreground=fg, background=bg,
                                           font=(UI_FONT, 10, "bold"))

        self.editor_hint = self._label(
            inner, "Pune eticheta (S1, R, B, C) pe un rând separat, deasupra versurilor. "
                   "Strofele fără etichetă se numerotează singure.",
            9, color=UI_MUTED)
        self.editor_hint.pack(fill="x", pady=(self.px(8), 0))
        self.editor_hint.bind("<Configure>", lambda e: e.widget.configure(wraplength=e.width))
        return card

    def _build_presentation(self, parent) -> tk.Frame:
        card = Card(parent)
        inner = tk.Frame(card, bg=UI_CARD)
        inner.pack(fill="both", expand=True, padx=self.px(18), pady=self.px(16))

        self._label(inner, "Ordinea de cântare", 10, True).pack(fill="x")
        self._label(inner, "Lasă gol pentru ordinea automată, sau scrie de ex. S1 R S2 R B R.",
                    9, color=UI_MUTED).pack(fill="x")

        row = tk.Frame(inner, bg=UI_CARD)
        row.pack(fill="x", pady=(self.px(6), 0))
        self.order_entry = ttk.Entry(row, font=("Consolas", 11))
        self.order_entry.pack(side="left", fill="x", expand=True)
        self.order_entry.bind("<KeyRelease>", lambda _e: self.schedule_refresh())
        ttk.Button(row, text="Completează automat",
                   command=self.fill_auto_order).pack(side="left", padx=(self.px(8), 0))

        self.order_status = self._label(inner, "", 9, color=UI_MUTED)
        self.order_status.pack(fill="x", pady=(self.px(6), 0))
        self.order_status.bind("<Configure>", lambda e: e.widget.configure(wraplength=e.width))

        # Previzualizarea: miniaturi negre exact în ordinea slide-urilor
        head = tk.Frame(inner, bg=UI_CARD)
        head.pack(fill="x", pady=(self.px(16), self.px(8)))
        self._label(head, "Previzualizare", 10, True).pack(side="left")
        self.count_label = self._label(head, "", 9, color=UI_MUTED)
        self.count_label.pack(side="right")

        preview = tk.Frame(inner, bg=UI_FIELD, highlightthickness=1, highlightbackground=UI_BORDER)
        preview.pack(fill="both", expand=True)
        self.canvas = tk.Canvas(preview, bg=UI_FIELD, highlightthickness=0, bd=0)
        pscroll = ttk.Scrollbar(preview, orient="vertical", command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=pscroll.set)
        pscroll.pack(side="right", fill="y")
        self.canvas.pack(side="left", fill="both", expand=True)
        self.canvas.bind("<Configure>", lambda _e: self.draw_preview())
        self.canvas.bind("<Enter>", lambda _e: self.canvas.bind_all("<MouseWheel>", self._on_wheel))
        self.canvas.bind("<Leave>", lambda _e: self.canvas.unbind_all("<MouseWheel>"))

        self.show_labels = tk.BooleanVar(value=False)
        ttk.Checkbutton(inner, text="Arată numele secțiunii pe slide (mic, în colț)",
                        variable=self.show_labels,
                        command=self.draw_preview).pack(fill="x", pady=(self.px(12), 0))

        self.generate_button = ttk.Button(inner, text="Generează PowerPoint",
                                          style="Accent.TButton", command=self.on_generate)
        self.generate_button.pack(fill="x", pady=(self.px(12), 0))
        return card

    # --- Editare ----------------------------------------------------------

    def _on_text_modified(self, _event=None):
        if self.lyrics_text.edit_modified():
            self.lyrics_text.edit_modified(False)
            self.schedule_refresh()

    def insert_label(self, label: str):
        """Inserează eticheta pe un rând nou, deasupra rândului curent,
        lăsând un rând gol înainte dacă e nevoie."""
        text = self.lyrics_text
        line_start = text.index("insert linestart")
        current = text.get(line_start, f"{line_start} lineend").strip()
        before = text.get(f"{line_start} -1 line linestart", f"{line_start} -1 line lineend").strip()

        prefix = "\n" if before and line_start != "1.0" else ""
        # cursorul rămâne pe rândul de versuri, imediat sub etichetă
        text.insert(line_start, f"{prefix}{label}\n")
        if not current:
            text.mark_set("insert", f"{line_start} +{2 if prefix else 1} lines linestart")
        text.see("insert")
        text.focus_set()

    def insert_verse(self):
        """Inserează eticheta strofei următoare (S1, S2...)."""
        song = parse_song(self.lyrics_text.get("1.0", "end"))
        numbers = [sec.number for sec in song.sections if sec.kind == "S"]
        self.insert_label(f"S{max(numbers, default=0) + 1}")

    def insert_new_chorus(self):
        """Inserează eticheta unui refren nou (R2, R3...), pentru cântările
        care au mai multe refrene diferite."""
        song = parse_song(self.lyrics_text.get("1.0", "end"))
        numbers = [sec.number or 1 for sec in song.sections if sec.kind == "R"]
        self.insert_label(f"R{max(numbers, default=0) + 1}" if numbers else "R")

    def clear_lyrics(self):
        if self.lyrics_text.get("1.0", "end").strip() and not messagebox.askyesno(
                "Golește versurile", "Ștergi tot textul din casetă?"):
            return
        self.lyrics_text.delete("1.0", "end")
        self.lyrics_text.focus_set()

    def fill_auto_order(self):
        if self.plan:
            self.order_entry.delete(0, "end")
            self.order_entry.insert(0, order_codes(self.plan.auto))
            self.refresh()

    def show_help(self):
        win = tk.Toplevel(self.root, bg=UI_CARD)
        win.title("Cum marchez secțiunile")
        win.transient(self.root)
        _dark_title_bar(win)
        text = tk.Text(win, font=(UI_FONT, 10), wrap="word", relief="flat", bd=0,
                       padx=self.px(22), pady=self.px(18), bg=UI_CARD, fg=UI_TEXT,
                       width=72, height=34, highlightthickness=0)
        text.insert("1.0", HELP_TEXT)
        for heading in ("Cum marchezi secțiunile cântării", "Ordinea de cântare",
                        "Eticheta pe slide"):
            start = text.search(heading, "1.0")
            text.tag_add("h", start, f"{start} lineend")
        text.tag_configure("h", font=(UI_FONT, 12, "bold"), spacing1=self.px(6),
                           spacing3=self.px(4))
        text.configure(state="disabled")
        text.pack(fill="both", expand=True)
        ttk.Button(win, text="Am înțeles", command=win.destroy).pack(
            anchor="e", padx=self.px(18), pady=(0, self.px(14)))

    # --- Recalcularea slide-urilor ---------------------------------------

    def schedule_refresh(self):
        """Recalculează după o scurtă pauză de la ultima tastă."""
        if self._refresh_job:
            self.root.after_cancel(self._refresh_job)
        self._refresh_job = self.root.after(250, self.refresh)

    def refresh(self):
        self._refresh_job = None
        lyrics = self.lyrics_text.get("1.0", "end")
        self.plan = build_plan(lyrics, self.title_entry.get(), self.order_entry.get())
        self.highlight_labels()
        self.update_status()
        self.draw_preview()

    def highlight_labels(self):
        """Colorează rândurile cu etichete, ca secțiunile să se vadă clar."""
        text = self.lyrics_text
        for kind in KIND_COLORS:
            text.tag_remove(f"hdr_{kind}", "1.0", "end")

        at_block_start = True
        for i, line in enumerate(text.get("1.0", "end-1c").split("\n"), start=1):
            if not line.strip():
                at_block_start = True
                continue
            header = parse_header(line, at_block_start)
            at_block_start = False
            if header:
                kind, _num, rest = header
                # colorăm doar eticheta, nu și versul scris după „:”
                end = len(line) - len(line.lstrip()) + len(line.strip()) - len(rest)
                text.tag_add(f"hdr_{kind}", f"{i}.0", f"{i}.{end}")

    def update_status(self):
        plan = self.plan
        if plan.errors:
            self.order_status.configure(text="  ".join(plan.errors), fg=UI_ERROR)
        elif plan.song.warnings:
            self.order_status.configure(text="  ".join(plan.song.warnings), fg=UI_WARN)
        elif not plan.song.sections:
            self.order_status.configure(text="", fg=UI_MUTED)
        else:
            prefix = "Ordinea ta: " if self.order_entry.get().strip() else "Ordine automată: "
            codes = " ".join(s.code for s in plan.slides if s.code)
            self.order_status.configure(text=prefix + codes, fg=UI_MUTED)
        n = len(plan.slides)
        self.count_label.configure(text=f"{n} slide-uri" if n != 1 else "1 slide")

    # --- Previzualizarea -------------------------------------------------

    def _thumb_font(self, size_px: int) -> tkfont.Font:
        if size_px not in self._thumb_fonts:
            self._thumb_fonts[size_px] = tkfont.Font(family=FONT_NAME, size=-size_px,
                                                     weight="bold")
        return self._thumb_fonts[size_px]

    def draw_preview(self):
        """Desenează miniaturile slide-urilor, la scară, cu aceeași mărime
        relativă a fontului ca în fișierul generat."""
        c = self.canvas
        c.delete("all")
        if not self.plan:
            return

        width = c.winfo_width()
        if width < 50:
            return
        gap = self.px(self.THUMB_GAP)
        cols = 2 if width > self.px(380) else 1
        tw = (width - gap * (cols + 1)) / cols
        th = tw * SLIDE_HEIGHT_IN / SLIDE_WIDTH_IN
        caption_h = self.px(22)
        pt_to_px = tw / (SLIDE_WIDTH_IN * 72)  # pixeli de miniatură per punct
        show_labels = self.show_labels.get()

        if not self.plan.slides:
            c.create_text(width / 2, self.px(60), text="Slide-urile apar aici pe măsură ce scrii.",
                          fill=UI_MUTED, font=(UI_FONT, 10))
            c.configure(scrollregion=(0, 0, width, self.px(120)))
            return

        for i, slide in enumerate(self.plan.slides):
            col, row = i % cols, i // cols
            x = gap + col * (tw + gap)
            y = gap + row * (th + caption_h + gap)

            # legenda: numărul slide-ului + secțiunea, colorată după tip
            fg = KIND_COLORS[slide.kind][0] if slide.kind else UI_MUTED
            c.create_text(x, y + caption_h / 2 - self.px(2), anchor="w",
                          text=f"{i + 1}", fill=UI_MUTED, font=(UI_FONT, 9))
            c.create_text(x + self.px(18), y + caption_h / 2 - self.px(2), anchor="w",
                          text=slide.label or "Titlu", fill=fg, font=(UI_FONT, 9, "bold"))

            top = y + caption_h
            # conturul subțire desparte slide-ul negru de fundalul întunecat
            c.create_rectangle(x, top, x + tw, top + th, fill="#000000", outline=UI_BORDER)

            label = slide.label if show_labels else ""
            font_pt = compute_font_size(slide.text, 2 * LABEL_SPACE_IN if label else 0)
            size_px = max(4, round(font_pt * pt_to_px))
            c.create_text(x + tw / 2, top + th / 2, text=slide.text, fill="#FFFFFF",
                          justify="center", font=self._thumb_font(size_px),
                          width=tw - 2 * MARGIN_IN * 72 * pt_to_px)
            if label:
                c.create_text(x + MARGIN_IN * 72 * pt_to_px, top + 0.3 * 72 * pt_to_px,
                              anchor="nw", text=label, fill="#" + LABEL_COLOR,
                              font=(FONT_NAME, -max(4, round(LABEL_FONT_PT * pt_to_px))))

        rows = (len(self.plan.slides) + cols - 1) // cols
        c.configure(scrollregion=(0, 0, width, gap + rows * (th + caption_h + gap)))

    def _on_wheel(self, event):
        self.canvas.yview_scroll(int(-event.delta / 120), "units")

    # --- Generarea -------------------------------------------------------

    def on_generate(self):
        """Handler-ul butonului: validează, cere destinația și generează."""
        self.refresh()
        plan = self.plan
        title = self.title_entry.get()

        # Validare: textarea goală => mesaj de eroare, nu crash
        if not plan.song.sections:
            messagebox.showerror(
                "Text lipsă",
                "Lipește mai întâi versurile cântării în caseta de text.",
            )
            return
        if plan.errors:
            messagebox.showerror(
                "Ordinea de cântare",
                "Ordinea de cântare are greșeli:\n\n" + "\n".join(plan.errors)
                + "\n\nCorectează-le sau golește câmpul pentru ordinea automată.",
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
            num_slides = write_pptx(plan.slides, title, output_path, self.show_labels.get())
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

        if messagebox.askyesno(
            "Gata",
            f"Prezentarea a fost salvată ({num_slides} slide-uri):\n{output_path}\n\n"
            "O deschid acum?",
        ) and hasattr(os, "startfile"):
            os.startfile(output_path)


def main():
    _enable_dpi_awareness()
    root = tk.Tk()
    Vers2GoApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
