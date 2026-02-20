#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
print_pdfs.py (Windows)

Druckt alle PDFs in einem Ordner.
- nutzt SumatraPDF für zuverlässiges "silent printing"
- setzt Duplex/Farbe/Seiten/Kopien per SumatraPDF "-print-settings" (kein SetPrinter/keine Admin-Rechte nötig)
- unterstützt Include/Exclude Filter über Glob-Patterns (z.B. "*_invoice.pdf")

Beispiele:
  python .\print_pdfs.py "C:\PDFs" --duplex long-edge --color mono --pages "1,3"
  python .\print_pdfs.py "C:\PDFs" --filter "*_invoice.pdf" --exclude "*_draft.pdf"
"""

from __future__ import annotations

import argparse
import fnmatch
import os
import shutil
import subprocess
import sys
from pathlib import Path
from typing import List, Optional


# -------------------- SumatraPDF Finden --------------------

def find_sumatra(sumatra_arg: Optional[str]) -> Optional[str]:
    """
    Findet SumatraPDF.exe.
    - wenn --sumatra gesetzt, wird das verwendet
    - sonst PATH
    - sonst typische Installationsorte
    """
    if sumatra_arg:
        p = Path(sumatra_arg)
        if p.is_file():
            return str(p)
        w = shutil.which(sumatra_arg)
        if w:
            return w
        return None

    w = shutil.which("SumatraPDF.exe")
    if w:
        return w

    pf = os.environ.get("ProgramFiles", r"C:\Program Files")
    pf86 = os.environ.get("ProgramFiles(x86)", r"C:\Program Files (x86)")
    lad = os.environ.get("LOCALAPPDATA", "")

    candidates = [
        Path(pf) / "SumatraPDF" / "SumatraPDF.exe",
        Path(pf86) / "SumatraPDF" / "SumatraPDF.exe",
        Path(lad) / "SumatraPDF" / "SumatraPDF.exe",
    ]
    for c in candidates:
        if c.is_file():
            return str(c)
    return None


# -------------------- PDF Liste + Filter --------------------

def list_pdfs(
    folder: Path,
    recursive: bool,
    include_patterns: Optional[List[str]],
    exclude_patterns: Optional[List[str]],
) -> List[Path]:
    pattern = "**/*.pdf" if recursive else "*.pdf"
    files = sorted(folder.glob(pattern))
    files = [f for f in files if f.is_file()]

    def matches_any(path: Path, patterns: List[str]) -> bool:
        name = path.name
        return any(fnmatch.fnmatch(name, p) for p in patterns)

    # Include-Filter (wenn gesetzt, nur diese)
    if include_patterns:
        files = [f for f in files if matches_any(f, include_patterns)]

    # Exclude-Filter (wenn gesetzt, diese rauswerfen)
    if exclude_patterns:
        files = [f for f in files if not matches_any(f, exclude_patterns)]

    return files


# -------------------- Sumatra Drucken --------------------

def build_sumatra_print_settings(duplex: str, color: str, pages: Optional[str], copies: int) -> Optional[str]:
    """
    Baut den String für SumatraPDF -print-settings.
    Typische Tokens:
      - Seiten: "1,3" oder "1-3,5"
      - Duplex: simplex | duplex | duplexlong | duplexshort
      - Farbe:  color | monochrome
      - Kopien: "3x"
    """
    parts: List[str] = []

    # Seiten
    if pages:
        # Sumatra akzeptiert 1-basierte Seitenangaben wie "1,3" oder "1-3,5"
        parts.append(pages)

    # Duplex Mapping zu Sumatra
    duplex_map = {
        "default": None,
        "simplex": "simplex",
        "duplex": "duplex",          # Treiber entscheidet oft long/short
        "long-edge": "duplexlong",
        "short-edge": "duplexshort",
        "tumble": "duplexshort",
    }
    d = duplex_map.get(duplex)
    if d:
        parts.append(d)

    # Farbe
    color_map = {
        "default": None,
        "color": "color",
        "mono": "monochrome",
    }
    c = color_map.get(color)
    if c:
        parts.append(c)

    # Kopien
    if copies and copies > 1:
        parts.append(f"{copies}x")

    return ",".join(parts) if parts else None


def print_with_sumatra(sumatra_exe: str, pdf_path: Path, printer_name: str, print_settings: Optional[str]) -> None:
    """
    Druckt per SumatraPDF silent mit optionalen -print-settings.
    """
    cmd = [
        sumatra_exe,
        "-print-to", printer_name,
        "-silent",
        "-exit-on-print",
    ]
    if print_settings:
        cmd += ["-print-settings", print_settings]

    subprocess.run(cmd + [str(pdf_path)], check=True)


# -------------------- Main --------------------

def main() -> int:
    parser = argparse.ArgumentParser(description="Druckt alle PDFs in einem Ordner (Windows, SumatraPDF).")
    parser.add_argument("folder", type=str, help="Ordner mit PDFs")
    parser.add_argument("--recursive", action="store_true", help="Unterordner einbeziehen")
    parser.add_argument("--printer", type=str, default=None, help="Druckername (Standard: Windows Standarddrucker)")
    parser.add_argument("--sumatra", type=str, default=None, help="Pfad zu SumatraPDF.exe (optional)")

    parser.add_argument(
        "--duplex",
        type=str,
        default="default",
        choices=["default", "simplex", "duplex", "long-edge", "short-edge", "tumble"],
        help="Duplex-Modus (über Sumatra -print-settings)",
    )
    parser.add_argument(
        "--color",
        type=str,
        default="default",
        choices=["default", "color", "mono"],
        help="Farbmodus (über Sumatra -print-settings)",
    )
    parser.add_argument("--pages", type=str, default=None, help='Seiten z.B. "1-3,5,7-" oder "1,3" (1-basiert)')
    parser.add_argument("--copies", type=int, default=1, help="Anzahl Kopien (Standard 1)")

    parser.add_argument(
        "--filter",
        action="append",
        default=None,
        help='Nur PDFs drucken, die zu diesem Glob passen (kann mehrfach angegeben werden), z.B. "*_invoice.pdf"',
    )
    parser.add_argument(
        "--exclude",
        action="append",
        default=None,
        help='PDFs ausschließen, die zu diesem Glob passen (kann mehrfach angegeben werden), z.B. "*_draft.pdf"',
    )

    parser.add_argument("--dry-run", action="store_true", help="Nur anzeigen, nichts drucken")
    args = parser.parse_args()

    folder = Path(args.folder).expanduser().resolve()
    if not folder.is_dir():
        print(f"Fehler: Ordner existiert nicht: {folder}", file=sys.stderr)
        return 2

    # Drucker: wenn nicht angegeben, versuchen wir über Windows default.
    printer = args.printer
    if not printer:
        # Kein pywin32 mehr nötig: wir lesen Default-Drucker via PowerShell aus
        try:
            ps = [
                "powershell",
                "-NoProfile",
                "-Command",
                "(Get-CimInstance Win32_Printer | Where-Object {$_.Default -eq $true} | Select-Object -First 1 -ExpandProperty Name)"
            ]
            res = subprocess.run(ps, capture_output=True, text=True, check=True)
            printer = (res.stdout or "").strip()
        except Exception:
            printer = ""

        if not printer:
            print(
                "Fehler: Konnte keinen Standarddrucker ermitteln. Bitte gib --printer \"Druckername\" an.",
                file=sys.stderr,
            )
            return 2

    sumatra = find_sumatra(args.sumatra)
    if not sumatra:
        print(
            "Fehler: SumatraPDF.exe nicht gefunden.\n"
            "Installiere SumatraPDF oder gib den Pfad an: --sumatra \"C:\\Path\\SumatraPDF.exe\"",
            file=sys.stderr,
        )
        return 2

    pdfs = list_pdfs(folder, args.recursive, args.filter, args.exclude)
    if not pdfs:
        print(f"Keine passenden PDFs gefunden in: {folder}")
        return 0

    sumatra_settings = build_sumatra_print_settings(args.duplex, args.color, args.pages, args.copies)

    print(f"Ordner:   {folder}")
    print(f"PDFs:     {len(pdfs)}")
    print(f"Drucker:  {printer}")
    print(f"Sumatra:  {sumatra}")
    print(f"Duplex:   {args.duplex}")
    print(f"Farbe:    {args.color}")
    print(f"Seiten:   {args.pages or 'alle'}")
    print(f"Kopien:   {args.copies}")
    print(f"Filter:   {', '.join(args.filter) if args.filter else '-'}")
    print(f"Exclude:  {', '.join(args.exclude) if args.exclude else '-'}")
    print(f"Settings: {sumatra_settings or '-'}")
    print(f"Dry-Run:  {args.dry_run}")
    print("")

    for i, pdf in enumerate(pdfs, start=1):
        print(f"[{i}/{len(pdfs)}] {pdf.name}")
        try:
            if args.dry_run:
                print(f"  -> würde drucken: {pdf}")
            else:
                print_with_sumatra(sumatra, pdf, printer, sumatra_settings)
                print("  -> Druckauftrag an Spooler übergeben.")
        except subprocess.CalledProcessError as e:
            print(f"  !! Druck fehlgeschlagen (Sumatra Exitcode): {e.returncode}", file=sys.stderr)
        except Exception as e:
            print(f"  !! Fehler: {e}", file=sys.stderr)

    print("\nFertig.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())