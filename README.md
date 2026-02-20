# PDF Ordner-Drucker (Windows) – print_pdfs.py

Dieses Tool druckt alle PDFs in einem Ordner unter Windows – ohne manuell jedes PDF zu öffnen.

---

## Features

- Druckt alle PDFs in einem Ordner (optional rekursiv)
- Parameter für:
  - Duplex (simplex / long-edge / short-edge)
  - Farbe (color / mono)
  - Seiten-Auswahl (z.B. "1-3,5,7-" oder "1,3")
  - Kopien
  - PDF-Filter (z.B. nur `*_invoice.pdf`)
  - Umgekehrte Reihenfolge der (gefilterten) Dateien (`--reverse`)
- Nutzt SumatraPDF.exe für zuverlässiges "silent printing"
- Setzt Duplex/Farbe/Seiten/Kopien pro Job per SumatraPDF `-print-settings` (keine Admin-Rechte nötig)

---

## Voraussetzungen

1) Windows 10/11  
2) Python 3.9+  
3) SumatraPDF installiert  
4) Python-Pakete (keine Admin-Rechte nötig):
   - (keine zwingenden Pakete; Script nutzt Standardbibliothek)

SumatraPDF: https://www.sumatrapdfreader.org/

---

## Installation

```powershell
python -m pip install --upgrade pip
```

## Nutzung
Alle PDFs im Ordner drucken (Default-Drucker)
```powershell
python .\print_pdfs_win.py "C:\PDFs"
Rekursiv (Unterordner)
```powershell
python .\print_pdfs_win.py "C:\PDFs" --recursive
```
Auf bestimmten Drucker
```powershell
python .\print_pdfs_win.py "C:\PDFs" --printer "HP LaserJet M404"
```
Duplex + Schwarzweiß
```powershell
python .\print_pdfs_win.py "C:\PDFs" --duplex long-edge --color mono
```
Nur bestimmte Seiten (z.B. 1-3, 5, ab 7 bis Ende)
```powershell
python .\print_pdfs_win.py "C:\PDFs" --pages "1-3,5,7-"
```
Mehrere Kopien
```powershell
python .\print_pdfs_win.py "C:\PDFs" --copies 2
```
Dry-Run (nur anzeigen, nicht drucken)
```powershell
python .\print_pdfs_win.py "C:\PDFs" --dry-run
```

## Manueller Duplex (Laserjet 479fnw)
180 Grad rotieren, ohne wenden
## Hinweise / Einschränkungen

- Duplex/Farbe:

    Das Script setzt Windows-DEVMODE Felder (Duplex/Color). Ob der Drucker/ Treiber das wirklich übernimmt, hängt stark vom Druckertreiber ab.

    Manche Treiber ignorieren diese Einstellungen oder nutzen eigene Optionen.

    Das Script ändert die Drucker-Defaults temporär und stellt sie anschließend wieder her.

    Zwischen "Setzen" und "Wiederherstellen" wartet das Script standardmäßig 1.5s (--restore-delay), damit der Spooler den Job sicher übernommen hat.

- SumatraPDF:

    Das Script erwartet SumatraPDF.exe. Falls es nicht gefunden wird:

    python .\print_pdfs_win.py "C:\PDFs" --sumatra "C:\Program Files\SumatraPDF\SumatraPDF.exe"

- Troubleshooting

"SumatraPDF.exe nicht gefunden"

Installiere SumatraPDF oder verwende --sumatra.

"pywin32 nicht installiert"

pip install pywin32

Seiten-Parsing Fehler

Erlaubtes Format: "1-3,5,7-" (Komma-getrennt, Bereiche mit "-")


# PDF-Filter (nur bestimmte Dateien drucken)

Du kannst Dateien per Glob-Pattern **einschließen** (`--filter`) oder **ausschließen** (`--exclude`).
Beide Optionen können **mehrfach** angegeben werden.

## Beispiele

### Nur Rechnungen drucken (z.B. *_invoice.pdf)
```powershell
python .\print_pdfs_win.py "C:\PDFs" --filter "*_invoice.pdf"
```

Nur PDFs, die mit "2026_" anfangen
```powershell
python .\print_pdfs_win.py "C:\PDFs" --filter "2026_*.pdf"
```
Mehrere Filter (z.B. Rechnungen ODER Lieferscheine)
```powershell
python .\print_pdfs_win.py "C:\PDFs" --filter "*_invoice.pdf" --filter "*_deliverynote.pdf"
```
Entwürfe ausschließen
```powershell
python .\print_pdfs_win.py "C:\PDFs" --exclude "*_draft.pdf"
```
Kombination: nur Rechnungen, aber ohne Stornos
```powershell
python .\print_pdfs_win.py "C:\PDFs" --filter "*_invoice.pdf" --exclude "*_credit*.pdf"
```

Hinweis: Die Patterns gelten nur für den Dateinamen (nicht den Pfad).