# Vers2Go — Versuri în prezentare PowerPoint

Aplicație desktop (Python + tkinter) care transformă versurile unei melodii
într-o prezentare PowerPoint (`.pptx`): fiecare strofă devine un slide cu
fundal negru și text alb, bold, centrat, cu mărimea fontului calculată
automat astfel încât textul să încapă pe slide.

**Zero dependențe externe:** fișierul `.pptx` (o arhivă ZIP cu XML în format
Office Open XML) este generat direct cu biblioteca standard Python. Aplicația
funcționează deci și pe sisteme unde politicile de securitate (Application
Control / WDAC) blochează DLL-urile native din pachetele pip.

## Cum funcționează

1. Lipești versurile în caseta de text (strofele se separă prin **linii goale**).
2. Opțional, completezi titlul melodiei — acesta devine primul slide.
3. Apeși **„Generează PowerPoint"** și alegi unde se salvează fișierul `.pptx`.

Mărimea fontului este calculată explicit în cod (pe baza numărului de linii
și a celei mai lungi linii din strofă), nu prin auto-fit-ul din PowerPoint,
deci prezentarea arată corect și în LibreOffice Impress sau Google Slides.

### Refrenul (marcajul `R:`)

O strofă care începe cu `R:` este recunoscută ca **refren**: marcajul este
eliminat de pe slide, iar refrenul este inserat automat **după fiecare
strofă** a melodiei. Exemplu:

```
Strofa întâi...

R: Refrenul melodiei...

Strofa a doua...
```

produce slide-urile: *strofa 1 → refren → strofa 2 → refren*. Marcajul e
acceptat și cu literă mică (`r:`), iar dacă melodia are mai multe refrene
marcate, fiecare strofă primește refrenul cel mai recent definit înaintea ei.

## Rulare în development

Necesită doar Python 3.10+ (tkinter este inclus în instalarea standard de
Python pe Windows; pe Linux poate fi nevoie de `sudo apt install python3-tk`).
Nu este nevoie de `pip install` — nu există dependențe:

```bash
python vers2go.py
```

## Generarea executabilului (PyInstaller)

Pe **Windows**:

```bash
python -m pip install pyinstaller
python -m PyInstaller --onefile --windowed --name Vers2Go --icon icon.ico --add-data "icon.ico;." vers2go.py
```

Pe **Linux** separatorul din `--add-data` este `:` în loc de `;`
(iar `--icon` cu .ico nu se aplică — iconițele de executabil sunt
un concept Windows):

```bash
python -m PyInstaller --onefile --windowed --name Vers2Go --add-data "icon.ico:." vers2go.py
```

- `--onefile` — produce un singur fișier executabil, fără foldere auxiliare.
- `--windowed` — aplicația pornește fără fereastră de consolă în fundal.
- `--icon icon.ico` — iconița fișierului .exe (cea din Explorer / taskbar).
- `--add-data "icon.ico;."` — împachetează iconița și pentru fereastra
  aplicației (afișată în colțul ferestrei la rulare).
- Folosește `python -m pip` / `python -m PyInstaller` (nu `pip` / `pyinstaller`
  direct) — pe sistemele cu Application Control, executabilele-shim `pip.exe`
  și `pyinstaller.exe` pot fi blocate, dar modulele rulate prin `python -m`
  funcționează.

Executabilul rezultat se găsește în folderul `dist/` (`dist/Vers2Go.exe`
pe Windows, `dist/Vers2Go` pe Linux) și rulează fără ca utilizatorul să
aibă Python sau vreo bibliotecă instalată.

> **Important:** PyInstaller **nu face cross-compilare**. Executabilul se
> generează doar pentru sistemul de operare pe care rulezi comanda:
> pentru un `.exe` de Windows rulezi PyInstaller pe Windows, pentru un
> binar de Linux rulezi aceeași comandă pe Linux.

> **Notă pentru sistemele cu Application Control:** executabilul produs de
> PyInstaller este nesemnat, deci pe un PC cu politică strictă (ca cel de
> development) poate fi el însuși blocat la rulare. Pe astfel de sisteme
> rulează aplicația direct cu `python vers2go.py`; executabilul e util
> pentru PC-urile fără restricții (de ex. laptopul de proiecție).

## Structura proiectului

| Fișier | Rol |
|---|---|
| `vers2go.py` | Aplicația completă (GUI + generarea prezentării) |
| `icon.ico` | Iconița aplicației (fereastră + executabil) |
| `requirements.txt` | Gol — nu există dependențe de runtime |
