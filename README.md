# Vers2Go — Versuri în prezentare PowerPoint

Aplicație desktop (Python + tkinter) care transformă versurile unei cântări
într-o prezentare PowerPoint (`.pptx`): fiecare secțiune (strofă, refren,
punte, final) devine un slide cu fundal negru și text alb, bold, centrat,
cu mărimea fontului calculată automat astfel încât textul să încapă pe slide.

**Zero dependențe externe:** fișierul `.pptx` (o arhivă ZIP cu XML în format
Office Open XML) este generat direct cu biblioteca standard Python. Aplicația
funcționează deci și pe sisteme unde politicile de securitate (Application
Control / WDAC) blochează DLL-urile native din pachetele pip.

## Cum funcționează

1. Lipești versurile în caseta din stânga (secțiunile se separă prin
   **rânduri goale**) și, opțional, marchezi secțiunile (vezi mai jos).
2. Opțional, completezi titlul cântării — acesta devine primul slide.
3. În dreapta vezi în timp real slide-urile, în ordinea în care vor ieși.
4. Apeși **„Generează PowerPoint"** și alegi unde se salvează fișierul `.pptx`.

Mărimea fontului este calculată explicit în cod (pe baza numărului de linii
și a lățimii literelor din cel mai lung rând), nu prin auto-fit-ul din
PowerPoint, deci prezentarea arată corect și în LibreOffice Impress sau
Google Slides.

### Secțiunile cântării

Eticheta se scrie pe un rând separat, deasupra versurilor (butoanele
„+ Strofă", „+ Refren" etc. din aplicație o inserează la cursor):

| Etichetă | Secțiune | Ce înseamnă |
|---|---|---|
| `S1`, `S2`… | Strofa | versurile care se schimbă de la o repetare la alta |
| `R` | Refrenul | partea care se repetă identic |
| `R2`, `R3`… | Alt refren | când cântarea are mai multe refrene diferite |
| `P` | Pre-refren | 1-2 rânduri care duc spre refren (rar) |
| `B` | Puntea (bridge) | o parte diferită, de obicei o dată, spre final |
| `C` | Finalul (coda) | rândurile de încheiere |

Sunt acceptate și cuvinte întregi (`Strofa 2`, `Refren`, `Punte`, `Final`,
`Bridge`, `Chorus`), între paranteze (`[Refren]`) sau cu două puncte și
primul vers pe același rând (`R: Primul vers`, ca în versiunile vechi).
Blocurile fără etichetă devin strofe numerotate automat, iar un bloc
identic cu unul de dinainte e o repetare a lui. O etichetă fără versuri
sub ea (de ex. doar `R`) înseamnă „aici se repetă secțiunea".

Exemplu:

```
S1
Prima strofă...

R
Refrenul...

S2
A doua strofă...

B
Puntea...
```

### Ordinea de cântare

- **Automată** (câmpul gol): secțiunile se cântă în ordinea scrisă, iar
  dacă refrenul e scris o singură dată, el se repetă după fiecare strofă și
  după punte. Exemplul de mai sus devine *S1 R S2 R B R*. Cu mai multe
  refrene, fiecare strofă primește refrenul cel mai recent dinaintea ei.
- **Scrisă de tine**: de exemplu `S1 R S2 R B R R` sau `S1 R S2 Rx2 C`
  (`Rx2` = refrenul de două ori). Butonul „Completează automat" scrie
  ordinea automată în câmp, ca s-o poți modifica.

Bifa **„Arată numele secțiunii pe slide"** (oprită implicit) scrie mic și
gri, în colțul fiecărui slide, „Strofa 1", „Refren" etc. — util echipei de
proiecție.

## Rulare în development

Necesită doar Python 3.10+ (tkinter este inclus în instalarea standard de
Python pe Windows; pe Linux poate fi nevoie de `sudo apt install python3-tk`).
Nu este nevoie de `pip install` — nu există dependențe:

```bash
python vers2go.py
```

Testele (citirea secțiunilor, ordinea de cântare, validitatea fișierului):

```bash
python -m unittest test_vers2go
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

## Release-uri automate (installer profesional)

Repo-ul are un workflow GitHub Actions care, la fiecare tag de versiune,
construiește automat pe serverele GitHub atât executabilul cât și un
**installer Windows** (Inno Setup: wizard de instalare, scurtătură în
Start Menu / Desktop, uninstaller) și le publică la
[Releases](https://github.com/papura-octavian/Vers2Go/releases).

Publicarea unei versiuni noi:

```bash
git tag v1.0.1
git push origin v1.0.1
```

După ~3-4 minute, versiunea apare la Releases cu două fișiere:

- `Vers2Go-Setup-1.0.1.exe` — installerul (recomandat pentru utilizatori);
- `Vers2Go.exe` — executabilul portabil, fără instalare.

> Fișierele nu sunt semnate digital (certificatele de code-signing costă),
> deci la prima rulare Windows SmartScreen poate afișa avertismentul
> „Windows protected your PC" — se trece cu *More info → Run anyway*.

## Structura proiectului

| Fișier | Rol |
|---|---|
| `vers2go.py` | Aplicația completă (GUI + generarea prezentării) |
| `test_vers2go.py` | Testele (`python -m unittest test_vers2go`) |
| `icon.ico` | Iconița aplicației (fereastră + executabil) |
| `installer.iss` | Scriptul Inno Setup pentru installerul Windows |
| `.github/workflows/release.yml` | Build + release automat la fiecare tag `v*` |
| `requirements.txt` | Gol — nu există dependențe de runtime |
