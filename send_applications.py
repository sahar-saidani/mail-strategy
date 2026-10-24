#!/usr/bin/env python3
"""
Script d'envoi d'emails de candidature (stage de recherche) personnalisés via Gmail API.

Fonctionnalités :
- Extraction du texte du template Word (.docx) avec variables {salutation}, {laboratoire}, {chercheur}, {theme}
- Personnalisation par laboratoire / chercheur (colonnes du CSV contacts_labs.csv)
- Pièces jointes : CV (.pdf), lettre de motivation (.pdf, personnalisée par laboratoire si disponible)
  et lettre de recommandation (.pdf, optionnelle)
- Authentification officielle Gmail API (OAuth 2.0 avec credentials.json)
- Mode prévisualisation (--preview) sans aucun envoi
- Mode test d'envoi (--test-send votre_email@gmail.com) pour vérifier le rendu réel
- Journalisation (sent_emails.csv) et reprise automatique pour éviter les doublons
- Pause paramétrable entre chaque envoi pour respecter les quotas Gmail et la délivrabilité

Les lignes du CSV sans adresse email valide sont ignorées (aucune adresse n'est devinée).
"""

import argparse
import base64
import csv
import json
import os
import re
import time
import urllib.parse
import urllib.request
import urllib.error
import webbrowser
from datetime import datetime
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.utils import formatdate
from http.server import BaseHTTPRequestHandler, HTTPServer
import socket
import zipfile
import xml.etree.ElementTree as ET

from fill_letters import contact_slug  # même convention de nom que le générateur de lettres

# Configuration par défaut des fichiers dans le dossier
DEFAULT_CSV = "contacts_labs.csv"
DEFAULT_DOCX = "Email_template_sahar_saidani.docx"
DEFAULT_CV = "CV_recherche_sahar_saidani.pdf"
DEFAULT_COVER_LETTER = "Lettre_motivation_sahar_saidani.pdf"          # lettre par défaut (si pas de lettre dédiée)
DEFAULT_LETTERS_DIR = "letters"                                       # lettres personnalisées : letters/<slug>.pdf
DEFAULT_RECOMMENDATION = "Lettre_Recommandation_Heithem_ABBES_Sahar_Saidani.pdf"
DEFAULT_CREDENTIALS = "credentials.json"
DEFAULT_TOKEN = "token.json"
DEFAULT_LOG = "sent_emails.csv"
DEFAULT_SUBJECT = "Stage de recherche (février 2027) – {labo_nom} – Sahar Saidani"
GMAIL_SCOPES = ["https://www.googleapis.com/auth/gmail.send"]


# ==============================================================================
# 1. EXTRACTION DU TEXTE DEPUIS LE DOCUMENT WORD (.DOCX)
# ==============================================================================

def extract_docx_text(docx_path: str) -> str:
    """
    Extrait le texte structuré par paragraphes depuis un fichier .docx.
    Fonctionne sans dépendance externe via standard library (zipfile + xml).
    """
    if not os.path.exists(docx_path):
        raise FileNotFoundError(f"Le fichier template Word '{docx_path}' est introuvable.")

    try:
        # Essayer d'abord la bibliothèque python-docx si disponible
        import docx
        doc = docx.Document(docx_path)
        paragraphs = [p.text for p in doc.paragraphs if p.text.strip()]
        return "\n\n".join(paragraphs)
    except ImportError:
        # Décompression native du document XML OpenXML (.docx)
        with zipfile.ZipFile(docx_path, 'r') as z:
            xml_content = z.read('word/document.xml')
            tree = ET.fromstring(xml_content)
            ns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
            paragraphs = []
            w_ns = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
            for p in tree.findall('.//w:p', ns):
                texts = []
                for node in p.iter():
                    if node.tag == w_ns + 't' and node.text:
                        texts.append(node.text)
                    elif node.tag == w_ns + 'br':
                        texts.append("\n")   # saut de ligne (ex. signature sur plusieurs lignes)
                p_text = "".join(texts).strip()
                if p_text:
                    paragraphs.append(p_text)
            return "\n\n".join(paragraphs)


# ==============================================================================
# 2. PERSONNALISATION DU TEMPLATE
# ==============================================================================

def personalize_email(template_text: str, contact: dict) -> str:
    """
    Remplace les variables du template :
      {salutation}  -> ex. 'Monsieur le Professeur Gaaloul' (défaut : 'Madame, Monsieur')
      {laboratoire} -> nom complet du laboratoire
      {labo_nom}    -> nom court (ex. 'LIP – équipe Avalon')
      {labo_court}  -> formulation avec article, après « rejoindre » (ex. "l'équipe Avalon (Inria / LIP)")
      {chercheur}   -> nom du chercheur
      {theme}       -> thème de recherche (phrase qui suit 'vos travaux sur')
    """
    values = {
        "salutation": contact.get("salutation") or "Madame, Monsieur",
        "laboratoire": contact.get("laboratoire") or "votre laboratoire",
        "labo_nom": contact.get("labo_nom") or contact.get("laboratoire") or "votre laboratoire",
        "labo_court": contact.get("labo_court") or contact.get("laboratoire") or "votre laboratoire",
        "chercheur": contact.get("chercheur") or "",
        "theme": contact.get("theme") or "le DevOps et le cloud",
    }
    text = template_text
    for key, val in values.items():
        text = text.replace("{" + key + "}", val.strip())
    return text


def format_subject(subject_template: str, contact: dict) -> str:
    """Génère l'objet du mail avec variables {labo_nom}, {laboratoire} et {chercheur}."""
    return subject_template.format(
        laboratoire=(contact.get("laboratoire") or "").strip(),
        labo_nom=(contact.get("labo_nom") or contact.get("laboratoire") or "").strip(),
        chercheur=(contact.get("chercheur") or "").strip(),
    ).strip()


# ==============================================================================
# 3. CRÉATION DU MESSAGE MIME AVEC PIÈCES JOINTES
# ==============================================================================

def create_mime_message(
    to_email: str,
    subject: str,
    body_text: str,
    attachment_paths: list[str]
) -> MIMEMultipart:
    """Construit un message MIME multipart contenant le corps et les pièces jointes PDF."""
    msg = MIMEMultipart("mixed")
    msg["To"] = to_email
    msg["Subject"] = subject
    msg["Date"] = formatdate(localtime=False)

    # Corps du texte en UTF-8
    body_part = MIMEText(body_text, "plain", "utf-8")
    msg.attach(body_part)

    # Pièces jointes
    for path in attachment_paths:
        if not os.path.exists(path):
            raise FileNotFoundError(f"Pièce jointe introuvable : {path}")

        filename = os.path.basename(path)
        with open(path, "rb") as f:
            pdf_data = f.read()

        part = MIMEApplication(pdf_data, _subtype="pdf")
        part.add_header(
            "Content-Disposition",
            "attachment",
            filename=("utf-8", "", filename)
        )
        msg.attach(part)

    return msg


def encode_mime_to_raw(msg: MIMEMultipart) -> str:
    """Encode le message MIME au format base64 urlsafe attendu par l'API Gmail."""
    raw_bytes = msg.as_bytes()
    return base64.urlsafe_b64encode(raw_bytes).decode("ascii")


def resolve_attachments(contact: dict, cv_path: str, cover_letter_path: str,
                        letters_dir: str, recommendation_path: str | None) -> list[str]:
    """
    Liste des pièces jointes d'un contact :
    CV + lettre dédiée (letters/<slug>.pdf) sinon lettre par défaut + recommandation (si le fichier existe).
    """
    letter = os.path.join(letters_dir, contact_slug(contact) + ".pdf")
    if not os.path.exists(letter):
        letter = cover_letter_path
    paths = [cv_path, letter]
    if recommendation_path and os.path.exists(recommendation_path):
        paths.append(recommendation_path)
    return paths


# ==============================================================================
# 4. GESTION DE L'AUTHENTIFICATION GMAIL API (OAUTH 2.0)
# ==============================================================================

class OAuthRedirectHandler(BaseHTTPRequestHandler):
    """Serveur HTTP temporaire pour recevoir le code d'autorisation OAuth2."""
    auth_code = None
    error = None

    def do_GET(self):
        parsed_url = urllib.parse.urlparse(self.path)
        query = urllib.parse.parse_qs(parsed_url.query)

        if "code" in query:
            OAuthRedirectHandler.auth_code = query["code"][0]
            self.send_response(200)
            self.send_header("Content-Type", "text/html; charset=utf-8")
            self.end_headers()
            html = """
            <html>
                <body style="font-family: sans-serif; text-align: center; padding-top: 50px;">
                    <h2 style="color: #2e7d32;">✓ Authentification Gmail réussie !</h2>
                    <p>Vous pouvez fermer cet onglet et retourner dans votre terminal.</p>
                </body>
            </html>
            """
            self.wfile.write(html.encode("utf-8"))
        elif "error" in query:
            OAuthRedirectHandler.error = query["error"][0]
            self.send_response(400)
            self.send_header("Content-Type", "text/html; charset=utf-8")
            self.end_headers()
            self.wfile.write(f"Erreur d'authentification : {OAuthRedirectHandler.error}".encode("utf-8"))
        else:
            self.send_response(404)
            self.end_headers()

    def log_message(self, format, *args):
        # Silence HTTP logs
        return


def find_free_port(preferred_port=8080) -> int:
    """Trouve un port libre pour le serveur de redirection."""
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        try:
            s.bind(('127.0.0.1', preferred_port))
            return preferred_port
        except OSError:
            s.bind(('127.0.0.1', 0))
            return s.getsockname()[1]


def get_gmail_access_token(credentials_path: str = DEFAULT_CREDENTIALS, token_path: str = DEFAULT_TOKEN) -> str:
    """
    Obtient un access_token valide.
    Gère le chargement, le refresh automatique et la première authentification interactive via navigateur.
    """
    if not os.path.exists(credentials_path):
        print("\n" + "="*80)
        print("[-] ERREUR : Le fichier 'credentials.json' est introuvable !")
        print("    Veuillez suivre le guide de configuration Google Cloud Console")
        print("    et placer le fichier 'credentials.json' dans ce répertoire.")
        print("="*80 + "\n")
        raise FileNotFoundError(f"Identifiants introuvables : {credentials_path}")

    with open(credentials_path, "r", encoding="utf-8") as f:
        cred_data = json.load(f)

    # Récupérer client_id et client_secret (type 'installed' ou 'web')
    client_info = cred_data.get("installed") or cred_data.get("web")
    if not client_info:
        raise ValueError("Format invalide de credentials.json : clé 'installed' ou 'web' manquante.")

    client_id = client_info["client_id"]
    client_secret = client_info["client_secret"]
    token_uri = client_info.get("token_uri", "https://oauth2.googleapis.com/token")
    auth_uri = client_info.get("auth_uri", "https://accounts.google.com/o/oauth2/auth")

    tokens = {}
    if os.path.exists(token_path):
        try:
            with open(token_path, "r", encoding="utf-8") as f:
                tokens = json.load(f)
        except Exception:
            tokens = {}

    now = time.time()
    access_token = tokens.get("access_token")
    expires_at = tokens.get("expires_at", 0)
    refresh_token = tokens.get("refresh_token")

    # Si le token est encore valide (avec marge de 60s)
    if access_token and (expires_at - now > 60):
        return access_token

    # Si un refresh_token est disponible, on rafraîchit l'access_token
    if refresh_token:
        try:
            req_data = urllib.parse.urlencode({
                "client_id": client_id,
                "client_secret": client_secret,
                "refresh_token": refresh_token,
                "grant_type": "refresh_token"
            }).encode("utf-8")

            req = urllib.request.Request(token_uri, data=req_data, method="POST")
            req.add_header("Content-Type", "application/x-www-form-urlencoded")
            with urllib.request.urlopen(req) as resp:
                data = json.loads(resp.read().decode("utf-8"))

            tokens["access_token"] = data["access_token"]
            tokens["expires_at"] = now + data.get("expires_in", 3600)
            if "refresh_token" in data:
                tokens["refresh_token"] = data["refresh_token"]

            with open(token_path, "w", encoding="utf-8") as f:
                json.dump(tokens, f, indent=2)

            return tokens["access_token"]
        except Exception as e:
            print(f"[!] Échec du rafraîchissement du token ({e}). Nouvelle authentification requise.")

    # Nouvelle authentification interactive requise
    port = find_free_port(8080)
    redirect_uri = f"http://127.0.0.1:{port}/"

    auth_params = {
        "client_id": client_id,
        "redirect_uri": redirect_uri,
        "response_type": "code",
        "scope": " ".join(GMAIL_SCOPES),
        "access_type": "offline",
        "prompt": "consent"
    }
    authorization_url = f"{auth_uri}?{urllib.parse.urlencode(auth_params)}"

    print("\n" + "="*80)
    print("🔑 Authentification requise avec votre compte Google / Gmail.")
    print("   Ouverture automatique de la page d'autorisation dans votre navigateur...")
    print(f"   Si le navigateur ne s'ouvre pas, visitez cette URL :\n   {authorization_url}")
    print("="*80 + "\n")

    webbrowser.open(authorization_url)

    # Lancement du serveur local pour intercepter le code de callback
    server = HTTPServer(("127.0.0.1", port), OAuthRedirectHandler)
    while OAuthRedirectHandler.auth_code is None and OAuthRedirectHandler.error is None:
        server.handle_request()
    server.server_close()

    if OAuthRedirectHandler.error:
        raise RuntimeError(f"Erreur d'autorisation OAuth: {OAuthRedirectHandler.error}")

    auth_code = OAuthRedirectHandler.auth_code
    print("[+] Code d'autorisation reçu ! Échange avec les jetons Google...")

    # Échange du code d'autorisation contre les tokens
    token_payload = urllib.parse.urlencode({
        "code": auth_code,
        "client_id": client_id,
        "client_secret": client_secret,
        "redirect_uri": redirect_uri,
        "grant_type": "authorization_code"
    }).encode("utf-8")

    req = urllib.request.Request(token_uri, data=token_payload, method="POST")
    req.add_header("Content-Type", "application/x-www-form-urlencoded")
    with urllib.request.urlopen(req) as resp:
        token_response = json.loads(resp.read().decode("utf-8"))

    token_response["expires_at"] = time.time() + token_response.get("expires_in", 3600)

    # Sauvegarde locale du token.json
    with open(token_path, "w", encoding="utf-8") as f:
        json.dump(token_response, f, indent=2)

    print(f"[+] Jeton d'accès sauvegardé dans '{token_path}'.")
    return token_response["access_token"]


def send_gmail_message(access_token: str, raw_b64_message: str) -> dict:
    """Envoie un message via l'API REST Gmail."""
    send_url = "https://gmail.googleapis.com/gmail/v1/users/me/messages/send"
    payload = json.dumps({"raw": raw_b64_message}).encode("utf-8")

    req = urllib.request.Request(send_url, data=payload, method="POST")
    req.add_header("Authorization", f"Bearer {access_token}")
    req.add_header("Content-Type", "application/json")

    try:
        with urllib.request.urlopen(req) as response:
            return json.loads(response.read().decode("utf-8"))
    except urllib.error.HTTPError as e:
        err_msg = e.read().decode("utf-8", errors="replace")
        raise RuntimeError(f"Erreur API Gmail (HTTP {e.code}): {err_msg}")


# ==============================================================================
# 5. CHARGEMENT ET SUIVI DES ENVOIS (JOURNALISATION)
# ==============================================================================

def _extract_email(value: str) -> str:
    """Retourne la première adresse email valide trouvée dans une chaîne (sinon '')."""
    if not value:
        return ""
    match = re.search(r"[\w.+-]+@[\w.-]+\.\w{2,}", value)
    return match.group(0) if match else ""


def load_contacts(csv_path: str) -> tuple[list[dict], int]:
    """
    Charge et nettoie les contacts depuis le fichier CSV.

    Colonnes attendues (contacts_labs.csv) :
      laboratoire, chercheur, salutation, email, theme, ville, source
    Anciennes colonnes toujours acceptées : company -> laboratoire, hr / contact -> chercheur.

    Retourne (contacts, lignes_ignorées) : les lignes sans adresse email valide
    ne sont pas envoyables et sont ignorées (aucune adresse n'est devinée).
    """
    if not os.path.exists(csv_path):
        raise FileNotFoundError(f"Fichier de contacts CSV '{csv_path}' introuvable.")

    contacts = []
    skipped = 0
    with open(csv_path, "r", encoding="utf-8-sig") as f:
        reader = csv.DictReader(f)
        for row in reader:
            if not row or not any(v is not None and str(v).strip() for v in row.values()):
                continue
            laboratoire = (row.get("laboratoire") or row.get("company") or "").strip()
            chercheur = (row.get("chercheur") or row.get("hr") or row.get("contact") or "").strip()
            email = _extract_email((row.get("email") or row.get("suggested_contact") or "").strip())
            if email:
                contacts.append({
                    "laboratoire": laboratoire,
                    "chercheur": chercheur,
                    "salutation": (row.get("salutation") or "").strip(),
                    "theme": (row.get("theme") or "").strip(),
                    "ville": (row.get("ville") or "").strip(),
                    "labo_nom": (row.get("labo_nom") or "").strip(),
                    "labo_court": (row.get("labo_court") or "").strip(),
                    "email": email,
                })
            else:
                skipped += 1
    return contacts, skipped


def load_sent_emails(log_path: str) -> set[str]:
    """Charge la liste des emails déjà contactés pour éviter les doublons."""
    if not os.path.exists(log_path):
        return set()

    sent = set()
    with open(log_path, "r", encoding="utf-8") as f:
        reader = csv.reader(f)
        for row in reader:
            if row and len(row) >= 2 and row[0] != "timestamp":
                # N'ignorer que les envois réussis : une ligne en erreur peut être retentée
                status = row[5] if len(row) >= 6 else "SUCCESS"
                if status == "SUCCESS":
                    sent.add(row[1].strip().lower())
    return sent


def log_sent_email(log_path: str, email: str, chercheur: str, laboratoire: str, msg_id: str, status: str = "SUCCESS"):
    """Enregistre un envoi dans le fichier CSV de suivi."""
    file_exists = os.path.exists(log_path)
    with open(log_path, "a", encoding="utf-8", newline="") as f:
        writer = csv.writer(f)
        if not file_exists:
            writer.writerow(["timestamp", "email", "chercheur", "laboratoire", "message_id", "status"])
        writer.writerow([datetime.now().isoformat(), email, chercheur, laboratoire, msg_id, status])


# ==============================================================================
# 6. FONCTIONS DE TEST ET DE PRÉVISUALISATION
# ==============================================================================

def run_preview(
    csv_path: str,
    docx_path: str,
    cv_path: str,
    cover_letter_path: str,
    letters_dir: str,
    recommendation_path: str,
    subject_template: str,
    limit: int = 3
):
    """
    Vérifie tous les fichiers, simule la personnalisation et affiche l'output
    complet pour les premiers contacts sans rien envoyer.
    """
    print("\n" + "="*80)
    print("🔍 VÉRIFICATION ET PRÉVISUALISATION DU CONTENU (--preview)")
    print("="*80)

    # 1. Vérification du CV (obligatoire)
    if os.path.exists(cv_path):
        print(f"[✓] CV trouvé : '{cv_path}' ({os.path.getsize(cv_path) / 1024:.1f} KB)")
    else:
        print(f"[✗] ERREUR : CV manquant : '{cv_path}'")
        return
    if os.path.exists(recommendation_path):
        print(f"[✓] Lettre de recommandation trouvée : '{recommendation_path}'")
    else:
        print(f"[i] Lettre de recommandation absente ('{recommendation_path}') : elle ne sera pas jointe.")

    # 2. Vérification et lecture du template docx
    try:
        raw_template = extract_docx_text(docx_path)
        print(f"[✓] Template Word extrait avec succès depuis '{docx_path}'")
    except Exception as e:
        print(f"[✗] ERREUR lors de la lecture du template Word : {e}")
        return

    # 3. Chargement des contacts
    try:
        contacts, skipped = load_contacts(csv_path)
        print(f"[✓] Contacts CSV chargés : {len(contacts)} destinataires avec email dans '{csv_path}'")
        if skipped:
            print(f"[i] {skipped} ligne(s) sans adresse email valide ignorées (contact manuel requis).")
    except Exception as e:
        print(f"[✗] ERREUR lors de la lecture du fichier CSV : {e}")
        return

    # 4. Lettres de motivation : dédiées ou par défaut
    dedicated = 0
    for c in contacts:
        if os.path.exists(os.path.join(letters_dir, contact_slug(c) + ".pdf")):
            dedicated += 1
    print(f"[i] Lettres personnalisées trouvées : {dedicated}/{len(contacts)} (les autres utilisent '{cover_letter_path}').")
    if dedicated < len(contacts) and not os.path.exists(cover_letter_path):
        print(f"[✗] ATTENTION : lettre par défaut '{cover_letter_path}' absente et certains contacts n'ont pas de lettre dédiée.")
        print("    Générez-les avec : python fill_letters.py --pdf")

    # 5. Affichage des aperçus
    print("\n" + "-"*80)
    print(f"📄 APERÇU DES {min(limit, len(contacts))} PREMIERS EMAILS GÉNÉRÉS :")
    print("-"*80)

    for idx, c in enumerate(contacts[:limit], start=1):
        subj = format_subject(subject_template, c)
        body = personalize_email(raw_template, c)
        atts = resolve_attachments(c, cv_path, cover_letter_path, letters_dir, recommendation_path)

        print(f"\n--- [EXEMPLE #{idx}] Destinataire : {c['chercheur']} <{c['email']}> | {c['laboratoire']} ---")
        print(f"Objet : {subj}")
        print(f"Pièces jointes : {atts}")
        print("Corps du message :")
        print("-" * 40)
        print(body)
        print("-" * 40)

    print("\n" + "="*80)
    print(f"✓ Test de génération terminé avec succès pour {len(contacts)} contacts !")
    print("  Pour envoyer un email test réel dans votre boîte de réception :")
    print("    python send_applications.py --test-send votre_email@gmail.com")
    print("  Pour lancer l'envoi officiel :")
    print("    python send_applications.py --send")
    print("="*80 + "\n")


def run_test_send(
    target_email: str,
    csv_path: str,
    docx_path: str,
    cv_path: str,
    cover_letter_path: str,
    letters_dir: str,
    recommendation_path: str,
    subject_template: str,
    credentials_path: str,
    token_path: str
):
    """
    Envoie un unique email de test complet avec les pièces jointes à l'adresse indiquée
    afin de vérifier le rendu exact dans votre propre boîte de réception.
    """
    print("\n" + "="*80)
    print(f"🧪 ENVOI D'UN EMAIL TEST RÉEL VERS : {target_email}")
    print("="*80)

    raw_template = extract_docx_text(docx_path)
    contacts, _ = load_contacts(csv_path)
    sample_contact = contacts[0] if contacts else {
        "laboratoire": "Laboratoire Exemple", "chercheur": "Nom Exemple",
        "salutation": "Madame, Monsieur", "theme": "le DevOps et le cloud"}

    # Préparation du message
    subj = "[TEST] " + format_subject(subject_template, sample_contact)
    body = personalize_email(raw_template, sample_contact)
    attachments = resolve_attachments(sample_contact, cv_path, cover_letter_path, letters_dir, recommendation_path)

    for path in attachments:
        if not os.path.exists(path):
            print(f"[✗] ERREUR : Fichier introuvable : {path}")
            return

    mime_msg = create_mime_message(target_email, subj, body, attachments)
    raw_b64 = encode_mime_to_raw(mime_msg)

    # Authentification et envoi via Gmail API
    try:
        print("[*] Connexion à l'API Gmail...")
        token = get_gmail_access_token(credentials_path, token_path)
        print(f"[*] Envoi de l'email test à {target_email} avec {len(attachments)} pièce(s) jointe(s)...")
        res = send_gmail_message(token, raw_b64)
        print(f"\n[✓] SUCCÈS ! Email test envoyé avec succès (Gmail Message ID : {res.get('id')})")
        print(f"    Consultez votre boîte '{target_email}' pour vérifier l'objet, le texte et les PDF.")
    except Exception as e:
        print(f"\n[✗] ERREUR lors de l'envoi test : {e}")


# ==============================================================================
# 7. ENVOI EN MASSE (PRODUCTION)
# ==============================================================================

def run_bulk_send(
    csv_path: str,
    docx_path: str,
    cv_path: str,
    cover_letter_path: str,
    letters_dir: str,
    recommendation_path: str,
    subject_template: str,
    credentials_path: str,
    token_path: str,
    log_path: str,
    delay: float = 3.0,
    limit: int = None,
    auto_confirm: bool = False
):
    """Envoi en masse des candidatures avec confirmation de sécurité et temporisation."""
    # 1. Vérification des prérequis
    if not os.path.exists(cv_path):
        print(f"[✗] ERREUR : CV introuvable : {cv_path}")
        return

    raw_template = extract_docx_text(docx_path)
    all_contacts, skipped = load_contacts(csv_path)
    sent_emails = load_sent_emails(log_path)

    # Filtrer les contacts déjà traités
    pending_contacts = [c for c in all_contacts if c["email"].lower() not in sent_emails]

    # Vérifier que chaque contact a bien une lettre (dédiée ou par défaut)
    missing = []
    for c in pending_contacts:
        for path in resolve_attachments(c, cv_path, cover_letter_path, letters_dir, recommendation_path):
            if not os.path.exists(path):
                missing.append((c["email"], path))
    if missing:
        print("[✗] ERREUR : pièces jointes manquantes, envoi annulé :")
        for e, p in missing[:10]:
            print(f"    {e} -> {p}")
        print("    Générez les lettres avec : python fill_letters.py --pdf")
        return

    if limit:
        pending_contacts = pending_contacts[:limit]

    print("\n" + "="*80)
    print("🚀 PRÉPARATION DE L'ENVOI OFFICIEL DES CANDIDATURES")
    print("="*80)
    print(f"Total contacts avec email dans le CSV : {len(all_contacts)}")
    print(f"Lignes sans email valide (ignorées)   : {skipped}")
    print(f"Déjà envoyés précédemment             : {len(sent_emails)}")
    print(f"Restants à envoyer                    : {len(pending_contacts)}")
    print(f"Délai entre chaque mail    : {delay} secondes")
    print(f"Fichier de suivi           : {log_path}")
    print("="*80)

    if not pending_contacts:
        print("[i] Aucun email en attente d'envoi. Tous les contacts ont déjà été traités.")
        return

    # Demande de confirmation explicite de l'utilisateur
    if not auto_confirm:
        print(f"\n⚠️  ATTENTION : Vous êtes sur le point d'envoyer {len(pending_contacts)} emails réels à des chercheurs !")
        confirm = input("Êtes-vous sûr de vouloir continuer ? Tapez 'OUI' pour valider : ").strip()
        if confirm != "OUI":
            print("[-] Envoi annulé par l'utilisateur.")
            return

    # Authentification Gmail API
    token = get_gmail_access_token(credentials_path, token_path)

    print("\n[*] Démarrage de l'envoi...")
    success_count = 0
    error_count = 0

    for idx, c in enumerate(pending_contacts, start=1):
        email = c["email"]
        chercheur = c["chercheur"]
        laboratoire = c["laboratoire"]

        subj = format_subject(subject_template, c)
        body = personalize_email(raw_template, c)

        print(f"[{idx}/{len(pending_contacts)}] Envoi à {chercheur} ({laboratoire}) <{email}> ...", end=" ", flush=True)

        try:
            # Re-validation du token si besoin (si la session dure plus d'une heure)
            token = get_gmail_access_token(credentials_path, token_path)

            attachments = resolve_attachments(c, cv_path, cover_letter_path, letters_dir, recommendation_path)
            mime_msg = create_mime_message(email, subj, body, attachments)
            raw_b64 = encode_mime_to_raw(mime_msg)
            res = send_gmail_message(token, raw_b64)

            msg_id = res.get("id", "UNKNOWN")
            log_sent_email(log_path, email, chercheur, laboratoire, msg_id, status="SUCCESS")
            print(f"[✓ OK (ID: {msg_id})]")
            success_count += 1
        except Exception as e:
            print(f"[✗ ERREUR : {e}]")
            log_sent_email(log_path, email, chercheur, laboratoire, "", status=f"ERROR: {e}")
            error_count += 1

        # Pause pour éviter de déclencher les filtres anti-spam de Google
        if idx < len(pending_contacts):
            time.sleep(delay)

    print("\n" + "="*80)
    print(f"🏁 ENVOI TERMINÉ : {success_count} réussis, {error_count} échecs.")
    print(f"   L'historique complet est enregistré dans '{log_path}'.")
    print("="*80 + "\n")


# ==============================================================================
# 8. POINT D'ENTRÉE CLI
# ==============================================================================

def main():
    parser = argparse.ArgumentParser(
        description="Envoi d'emails de candidature (stage de recherche) personnalisés via Gmail API."
    )
    group = parser.add_mutually_exclusive_group(required=True)
    group.add_argument(
        "--preview",
        action="store_true",
        help="Affiche le texte personnalisé et valide les pièces jointes sans rien envoyer."
    )
    group.add_argument(
        "--test-send",
        metavar="VOTRE_EMAIL",
        help="Envoie un seul email réel avec les PDF attachés à votre adresse personnelle pour tester le rendu."
    )
    group.add_argument(
        "--send",
        action="store_true",
        help="Lance l'envoi officiel aux contacts du CSV (demande confirmation)."
    )

    # Options de configuration des fichiers
    parser.add_argument("--csv", default=DEFAULT_CSV, help=f"Chemin vers le fichier CSV (défaut: {DEFAULT_CSV})")
    parser.add_argument("--docx", default=DEFAULT_DOCX, help=f"Chemin vers le template Word (défaut: {DEFAULT_DOCX})")
    parser.add_argument("--cv", default=DEFAULT_CV, help=f"Chemin vers le CV PDF (défaut: {DEFAULT_CV})")
    parser.add_argument("--cover-letter", default=DEFAULT_COVER_LETTER, help=f"Lettre PDF par défaut (défaut: {DEFAULT_COVER_LETTER})")
    parser.add_argument("--letters-dir", default=DEFAULT_LETTERS_DIR, help=f"Dossier des lettres personnalisées en PDF (défaut: {DEFAULT_LETTERS_DIR})")
    parser.add_argument("--recommendation", default=DEFAULT_RECOMMENDATION, help=f"Lettre de recommandation PDF, jointe si elle existe (défaut: {DEFAULT_RECOMMENDATION})")
    parser.add_argument("--credentials", default=DEFAULT_CREDENTIALS, help=f"Fichier d'identifiants OAuth (défaut: {DEFAULT_CREDENTIALS})")
    parser.add_argument("--token", default=DEFAULT_TOKEN, help=f"Fichier de token OAuth généré (défaut: {DEFAULT_TOKEN})")
    parser.add_argument("--log", default=DEFAULT_LOG, help=f"Fichier de suivi des envois (défaut: {DEFAULT_LOG})")
    parser.add_argument("--subject", default=DEFAULT_SUBJECT, help=f"Modèle d'objet (défaut: '{DEFAULT_SUBJECT}')")
    parser.add_argument("--delay", type=float, default=3.0, help="Délai en secondes entre chaque envoi (défaut: 3.0)")
    parser.add_argument("--limit", type=int, default=None, help="Limiter le nombre d'emails traités.")
    parser.add_argument("--yes", action="store_true", help="Ignorer la demande de confirmation pour l'envoi en masse.")

    args = parser.parse_args()

    if args.preview:
        run_preview(
            csv_path=args.csv,
            docx_path=args.docx,
            cv_path=args.cv,
            cover_letter_path=args.cover_letter,
            letters_dir=args.letters_dir,
            recommendation_path=args.recommendation,
            subject_template=args.subject,
            limit=args.limit or 3
        )
    elif args.test_send:
        run_test_send(
            target_email=args.test_send,
            csv_path=args.csv,
            docx_path=args.docx,
            cv_path=args.cv,
            cover_letter_path=args.cover_letter,
            letters_dir=args.letters_dir,
            recommendation_path=args.recommendation,
            subject_template=args.subject,
            credentials_path=args.credentials,
            token_path=args.token
        )
    elif args.send:
        run_bulk_send(
            csv_path=args.csv,
            docx_path=args.docx,
            cv_path=args.cv,
            cover_letter_path=args.cover_letter,
            letters_dir=args.letters_dir,
            recommendation_path=args.recommendation,
            subject_template=args.subject,
            credentials_path=args.credentials,
            token_path=args.token,
            log_path=args.log,
            delay=args.delay,
            limit=args.limit,
            auto_confirm=args.yes
        )


if __name__ == "__main__":
    main()
