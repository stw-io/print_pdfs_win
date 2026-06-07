# PDF Ordner-Drucker (Windows) – `print_pdfs_win.py`

Dieses Tool druckt alle PDFs in einem Ordner unter Windows, ohne jede Datei manuell zu öffnen. Es nutzt SumatraPDF für Silent Printing und setzt Optionen wie Duplex, Farbe, Seitenbereich, Kopien und Papierformat pro Druckjob über SumatraPDF `-print-settings`.

---

## Voraussetzungen

1. Windows 10/11
2. Python 3.9+
3. [SumatraPDF](https://www.sumatrapdfreader.org/) installiert
4. Optionales Python-Paket `pypdf` für Funktionen, die PDF-Seiten zählen oder temporäre PDFs erzeugen:
   - erforderlich für `--print-empty`
   - erforderlich für `--duplex fake`

```powershell
python -m pip install --upgrade pip
python -m pip install pypdf
```

Wenn du weder `--print-empty` noch `--duplex fake` verwendest, reichen Python-Standardbibliothek und SumatraPDF aus.

---

## Schnellstart

Alle PDFs im Ordner mit dem Windows-Standarddrucker drucken:

```powershell
python .\print_pdfs_win.py "C:\PDFs"
```

Trockenlauf ohne Druckauftrag:

```powershell
python .\print_pdfs_win.py "C:\PDFs" --dry-run
```

---

## Alle Parameter mit Beispielen

### `folder` (Pflichtparameter)

Ordner, dessen PDFs gedruckt werden sollen.

```powershell
python .\print_pdfs_win.py "C:\PDFs"
```

### `--recursive`

Bezieht Unterordner ein.

```powershell
python .\print_pdfs_win.py "C:\PDFs" --recursive
```

### `--printer "DRUCKERNAME"`

Druckt auf einen bestimmten Drucker. Ohne diesen Parameter wird der Windows-Standarddrucker per PowerShell ermittelt.

```powershell
python .\print_pdfs_win.py "C:\PDFs" --printer "HP LaserJet Pro MFP M479fnw"
```

### `--sumatra "PFAD"`

Verwendet eine bestimmte SumatraPDF.exe, falls sie nicht automatisch gefunden wird.

```powershell
python .\print_pdfs_win.py "C:\PDFs" --sumatra "C:\Program Files\SumatraPDF\SumatraPDF.exe"
```

### `--duplex`

Legt den Duplex-Modus fest. Mögliche Werte:

- `default` – keine Duplex-Einstellung an Sumatra übergeben
- `simplex` – einseitig drucken
- `duplex` – Duplex, Treiber entscheidet Bindung
- `long-edge` – Duplex über lange Kante
- `short-edge` – Duplex über kurze Kante
- `tumble` – Alias für `short-edge`
- `fake` – manueller Duplex-Workflow, siehe Abschnitt [Manueller Duplex mit `--duplex fake`](#manueller-duplex-mit---duplex-fake)

```powershell
python .\print_pdfs_win.py "C:\PDFs" --duplex long-edge
python .\print_pdfs_win.py "C:\PDFs" --duplex simplex
python .\print_pdfs_win.py "C:\PDFs" --duplex fake
```

### `--fake-pass`

Nur relevant mit `--duplex fake`. Steuert, welcher Teil des manuellen Duplex-Workflows läuft. Mögliche Werte:

- `both` – Standard: Vorderseiten und danach Rückseiten drucken
- `front` – nur Vorderseiten/ungerade Seiten drucken
- `back` – nur Rückseiten/gerade Seiten drucken, ohne vorher Vorderseiten zu drucken

```powershell
# Vollständiger manueller Duplex-Lauf
python .\print_pdfs_win.py "C:\PDFs" --duplex fake --fake-pass both

# Nur Vorderseiten drucken
python .\print_pdfs_win.py "C:\PDFs" --duplex fake --fake-pass front

# Rückseiten später manuell anstoßen
python .\print_pdfs_win.py "C:\PDFs" --duplex fake --fake-pass back
```

### `--color`

Legt den Farbmodus fest. Mögliche Werte:

- `default` – keine Farbeinstellung an Sumatra übergeben
- `color` – Farbe
- `mono` – Schwarzweiß/monochrom

```powershell
python .\print_pdfs_win.py "C:\PDFs" --color mono
python .\print_pdfs_win.py "C:\PDFs" --color color
```

### `--pages "SEITEN"`

Druckt nur bestimmte Seiten. Sumatra akzeptiert z.B. einzelne Seiten, Bereiche und offene Bereiche.

```powershell
python .\print_pdfs_win.py "C:\PDFs" --pages "1,3,5"
python .\print_pdfs_win.py "C:\PDFs" --pages "1-3,5,7-"
```

Hinweis: `--pages` kann nicht mit `--duplex fake` kombiniert werden.

### `--paper "FORMAT"`

Übergibt ein Papierformat an SumatraPDF (`paper=<Wert>`). Unterstützte/typische Werte sind z.B. `A4`, `A5`, `A3`, `Letter` oder `Legal`. Für intern erzeugte leere Seiten verwendet das Skript dieselbe Papiergröße.

```powershell
python .\print_pdfs_win.py "C:\PDFs" --paper A4
python .\print_pdfs_win.py "C:\PDFs" --paper A5
python .\print_pdfs_win.py "C:\PDFs" --paper Letter
```

### `--print-empty`

Ergänzt bei `--pages` fehlende Seiten als leere Seiten. Das ist nützlich, wenn z.B. Seiten `1-10` angefordert werden, ein PDF aber nur 8 Seiten hat.

```powershell
python .\print_pdfs_win.py "C:\PDFs" --pages "1,3,7-10" --print-empty
python .\print_pdfs_win.py "C:\PDFs" --pages "1-10" --print-empty --paper A4
```

Wichtig:

- Für `--print-empty` wird `pypdf` benötigt.
- Mit `--print-empty` müssen Seitenbereiche endlich sein. `--pages "7-"` ist dann nicht erlaubt; verwende z.B. `--pages "7-10"`.
- Wenn keine angeforderte Seite im PDF existiert, wird das Original-PDF übersprungen und es werden nur die benötigten leeren Seiten gedruckt.
- `--print-empty` kann nicht mit `--duplex fake` kombiniert werden.

### `--copies ANZAHL`

Druckt mehrere Kopien pro Job.

```powershell
python .\print_pdfs_win.py "C:\PDFs" --copies 2
python .\print_pdfs_win.py "C:\PDFs" --copies 3 --color mono
```

### `--filter "GLOB"`

Druckt nur PDFs, deren Dateiname zum Glob-Pattern passt. Der Parameter kann mehrfach angegeben werden.

```powershell
python .\print_pdfs_win.py "C:\PDFs" --filter "*_invoice.pdf"
python .\print_pdfs_win.py "C:\PDFs" --filter "2026_*.pdf"
python .\print_pdfs_win.py "C:\PDFs" --filter "*_invoice.pdf" --filter "*_deliverynote.pdf"
```

### `--exclude "GLOB"`

Schließt PDFs aus, deren Dateiname zum Glob-Pattern passt. Der Parameter kann mehrfach angegeben und mit `--filter` kombiniert werden.

```powershell
python .\print_pdfs_win.py "C:\PDFs" --exclude "*_draft.pdf"
python .\print_pdfs_win.py "C:\PDFs" --filter "*_invoice.pdf" --exclude "*_credit*.pdf"
```

Hinweis: `--filter` und `--exclude` prüfen nur den Dateinamen, nicht den vollständigen Pfad.

### `--dry-run`

Zeigt nur an, welche Druckaufträge erzeugt würden, ohne SumatraPDF zum Drucken aufzurufen.

```powershell
python .\print_pdfs_win.py "C:\PDFs" --dry-run
python .\print_pdfs_win.py "C:\PDFs" --duplex fake --fake-pass back --dry-run
```

### `--reverse`

Druckt die gefilterte PDF-Liste in umgekehrter Reihenfolge.

```powershell
python .\print_pdfs_win.py "C:\PDFs" --reverse
python .\print_pdfs_win.py "C:\PDFs" --filter "2026_*.pdf" --reverse
```

---

## Manueller Duplex mit `--duplex fake`

`--duplex fake` ist für Drucker gedacht, bei denen Duplex manuell über das erneute Einlegen des Papierstapels ausgeführt werden muss, z.B. HP LaserJet Pro MFP M479fnw.

### Vollständiger Ablauf (`--fake-pass both`)

```powershell
python .\print_pdfs_win.py "C:\PDFs" --duplex fake --fake-pass both
```

1. Das Skript druckt alle Vorderseiten: pro PDF nur ungerade Seiten in normaler Dateireihenfolge.
2. Danach fordert das Skript dazu auf, den Druckstapel um 180° in der horizontalen Ebene zu drehen und wieder einzulegen.
3. Das Skript wartet auf eine Eingabe:
   - `y` startet den Rückseitenlauf.
   - `n` bricht ab.
   - Jede andere Eingabe wird nicht als Abbruch gewertet; das Skript fragt erneut nach `y` oder `n`.
4. Der Rückseitenlauf druckt die PDFs in umgekehrter Dateireihenfolge und innerhalb jeder Datei die geraden Seiten rückwärts.
5. Hat ein PDF eine ungerade Seitenzahl, erzeugt das Skript für den Rückseitenlauf temporär eine zusätzliche leere Seite am Ende, damit die Seitenreihenfolge beim Stapel passt.

### Nur Vorderseiten drucken (`--fake-pass front`)

```powershell
python .\print_pdfs_win.py "C:\PDFs" --duplex fake --fake-pass front
```

Dieser Modus druckt nur ungerade Seiten. Es gibt keine Rückseiten-Abfrage.

### Nur Rückseiten manuell starten (`--fake-pass back`)

```powershell
python .\print_pdfs_win.py "C:\PDFs" --duplex fake --fake-pass back
```

Dieser Modus ist für den Fall gedacht, dass die Vorderseiten bereits gedruckt wurden und der Rückseitenlauf später separat gestartet werden soll. Das Skript druckt dann direkt die geraden Seiten rückwärts und die Dateien rückwärts. Es gibt keine vorherige Vorderseitenrunde und keine Fortfahren-Abfrage.

### Einschränkungen für `--duplex fake`

- `--duplex fake` kann nicht mit `--pages` kombiniert werden.
- `--duplex fake` kann nicht mit `--print-empty` kombiniert werden.
- Für `--duplex fake` wird `pypdf` benötigt, weil das Skript Seiten zählen und bei Bedarf temporäre PDFs mit einer leeren Ausgleichsseite erzeugen muss.

---

## Kombinierte Beispiele

Schwarzweiß, A5, zwei Kopien:

```powershell
python .\print_pdfs_win.py "C:\PDFs" --color mono --paper A5 --copies 2
```

Nur Rechnungen, ohne Stornos, in umgekehrter Reihenfolge:

```powershell
python .\print_pdfs_win.py "C:\PDFs" --filter "*_invoice.pdf" --exclude "*_credit*.pdf" --reverse
```

Fehlende Seiten als leere Seiten ergänzen:

```powershell
python .\print_pdfs_win.py "C:\PDFs" --pages "1-10" --print-empty --paper A4
```

Manuellen Duplex in zwei getrennten Befehlen ausführen:

```powershell
python .\print_pdfs_win.py "C:\PDFs" --duplex fake --fake-pass front
# Papierstapel manuell drehen und wieder einlegen
python .\print_pdfs_win.py "C:\PDFs" --duplex fake --fake-pass back
```

---

## Troubleshooting

### `SumatraPDF.exe nicht gefunden`

Installiere SumatraPDF oder gib den Pfad explizit an:

```powershell
python .\print_pdfs_win.py "C:\PDFs" --sumatra "C:\Program Files\SumatraPDF\SumatraPDF.exe"
```

### `Für --print-empty wird das Python-Paket 'pypdf' benötigt`

Installiere `pypdf`:

```powershell
python -m pip install pypdf
```

### `Für Duplex 'fake' wird das Python-Paket 'pypdf' benötigt`

Installiere `pypdf`:

```powershell
python -m pip install pypdf
```

### Seiten-Parsing-Fehler

Erlaubte Beispiele:

```text
1
1,3,5
1-3,5,7-
```

Mit `--print-empty` sind offene Bereiche wie `7-` nicht erlaubt. Verwende dann einen endlichen Bereich, z.B.:

```text
7-10
```
