#!/usr/bin/env python3
"""
Génère une lettre de motivation personnalisée par contact à partir du modèle .docx.

Variables du modèle (remplacées partout dans le document) :
  {{LABO_NOM}}      colonne 'labo_nom' (ex. "LIP – équipe Avalon") : en-tête destinataire et objet
  {{LABO_COURT}}    colonne 'labo_court' (ex. "l'équipe Avalon (Inria / LIP)") : phrases « rejoindre … »
  {{LABORATOIRE}}   colonne 'laboratoire' du CSV (nom complet)
  {{DESTINATAIRE}}  colonne 'salutation' (ex. "Monsieur le Professeur Gaaloul")
  {{VILLE}}         colonne 'ville'
  {{THEME}}         colonne 'theme' (phrase qui suit « vos travaux sur »)
  {{DATE}}          date du jour en français (ou --date)

Usage :
  python fill_letters.py                      # génère letters/*.docx pour tous les contacts avec email
  python fill_letters.py --pdf                # génère aussi les PDF (nécessite LibreOffice : soffice)
  python fill_letters.py --only "Gaaloul"     # un seul contact (filtre sur le nom)
"""
import argparse
import csv
import datetime
import os
import re
import shutil
import subprocess
import sys
import unicodedata
import zipfile
from xml.sax.saxutils import escape

MOIS = ["janvier", "février", "mars", "avril", "mai", "juin", "juillet",
        "août", "septembre", "octobre", "novembre", "décembre"]


def slugify(text: str) -> str:
    text = unicodedata.normalize("NFKD", text).encode("ascii", "ignore").decode("ascii")
    return re.sub(r"[^A-Za-z0-9]+", "_", text).strip("_").lower()


def contact_slug(row: dict) -> str:
    """Nom de fichier commun avec send_applications.py (laboratoire + chercheur)."""
    return slugify(f"{row.get('laboratoire', '')[:40]}_{row.get('chercheur', '')}")


def find_soffice():
    """Cherche LibreOffice dans le PATH puis aux emplacements habituels (Windows, macOS)."""
    found = shutil.which("soffice") or shutil.which("libreoffice")
    if found:
        return found
    candidates = [
        r"C:\Program Files\LibreOffice\program\soffice.exe",
        r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
        "/Applications/LibreOffice.app/Contents/MacOS/soffice",
    ]
    for path in candidates:
        if os.path.exists(path):
            return path
    return None


def fill_docx(template: str, out_path: str, values: dict) -> None:
    with zipfile.ZipFile(template) as zin, zipfile.ZipFile(out_path, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename.startswith("word/") and item.filename.endswith(".xml"):
                xml = data.decode("utf-8")
                for key, val in values.items():
                    xml = xml.replace("{{" + key + "}}", escape(val))
                data = xml.encode("utf-8")
            zout.writestr(item, data)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--csv", default="contacts_labs.csv")
    ap.add_argument("--template", default="Lettre_motivation_template_sahar_saidani.docx")
    ap.add_argument("--out", default="letters")
    ap.add_argument("--date", default=None, help="ex. '24 octobre 2026' (défaut : aujourd'hui)")
    ap.add_argument("--pdf", action="store_true", help="convertit aussi en PDF via LibreOffice")
    ap.add_argument("--only", default=None, help="ne traiter que les contacts dont le nom contient ce texte")
    ap.add_argument("--all", action="store_true", help="inclure aussi les lignes sans email")
    args = ap.parse_args()

    today = datetime.date.today()
    date_txt = args.date or f"{today.day} {MOIS[today.month - 1]} {today.year}"
    os.makedirs(args.out, exist_ok=True)

    soffice = find_soffice()
    word_convert = None
    if args.pdf and not soffice:
        try:
            from docx2pdf import convert as word_convert  # utilise Microsoft Word (Windows / macOS)
        except ImportError:
            sys.exit(
                "Aucun convertisseur PDF trouvé. Deux options :\n"
                "  1) installer LibreOffice (gratuit) : https://www.libreoffice.org\n"
                "  2) si Microsoft Word est installé : pip install docx2pdf, puis relancer cette commande."
            )
        print("[i] LibreOffice absent : conversion en PDF avec Microsoft Word (docx2pdf).")

    n = 0
    with open(args.csv, encoding="utf-8-sig") as f:
        for row in csv.DictReader(f):
            if not args.all and not (row.get("email") or "").strip():
                continue
            if args.only and args.only.lower() not in row.get("chercheur", "").lower():
                continue
            values = {
                "LABORATOIRE": row["laboratoire"].strip(),
                "LABO_NOM": (row.get("labo_nom") or row["laboratoire"]).strip(),
                "LABO_COURT": (row.get("labo_court") or row["laboratoire"]).strip(),
                "DESTINATAIRE": row["salutation"].strip(),
                "VILLE": row.get("ville", "").strip(),
                "THEME": row["theme"].strip(),
                "DATE": date_txt,
            }
            base = contact_slug(row)
            docx_path = os.path.join(args.out, base + ".docx")
            fill_docx(args.template, docx_path, values)
            n += 1
            if args.pdf:
                if soffice:
                    subprocess.run([soffice, "--headless", "--convert-to", "pdf", "--outdir", args.out, docx_path],
                                   check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
                else:
                    word_convert(os.path.abspath(docx_path), os.path.abspath(docx_path[:-5] + ".pdf"))
    print(f"{n} lettre(s) générée(s) dans '{args.out}/'.")


if __name__ == "__main__":
    main()