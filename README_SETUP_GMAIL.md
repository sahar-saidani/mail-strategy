# Guide Complet : Envoi Automatisé des Candidatures via l'API Gmail

Ce projet permet d'automatiser l'envoi personnalisé de vos candidatures de stage PFE aux recruteurs listés dans votre fichier CSV, avec votre CV et votre lettre de motivation en pièces jointes.

---

## 📁 Fichiers du projet

- `send_applications.py` : Script principal Python d'envoi et de test.
- `Email_template_Chaabane_Boussadia.docx` : Modèle de mail extrait et personnalisé automatiquement.
- `wave1_contacts.csv` : Liste des 73 contacts (colonnes : `company`, `hr`, `email`).
- `wave2_contacts.csv` : Liste de la 2ᵉ vague (34 contacts). Colonnes : `company`, `country`, `sector`, `fit_reason`, `suggested_contact`, `hr`, `email` (emails et contacts RH récupérés depuis `../PFEcontacts.xlsx`).
- `Chaabane_BOUSSADIA CV.pdf` : Votre CV (première pièce jointe).
- `Cover_Letter_Chaabane_Boussadia.pdf` : Votre lettre de motivation (deuxième pièce jointe).
- `sent_emails.csv` : Fichier généré automatiquement pour suivre l'historique et éviter les doublons.

---

## ⚙️ Étape 1 : Configuration de l'API Gmail (Google Cloud Console)

Pour envoyer des emails via votre propre compte Gmail en toute sécurité, suivez ces étapes (environ 5 minutes) :

### 1. Créer un projet Google Cloud
1. Rendez-vous sur la [Google Cloud Console](https://console.cloud.google.com/).
2. Connectez-vous avec votre compte Gmail personnel.
3. Cliquez sur le sélecteur de projet en haut à gauche, puis sur **"Nouveau projet"** (New Project).
4. Nommez-le par exemple `Candidatures-PFE` et cliquez sur **"Créer"**.

### 2. Activer l'API Gmail
1. Dans le menu de gauche, allez dans **"API et services"** > **"Bibliothèque"** (Library).
2. Recherchez **"Gmail API"**.
3. Cliquez sur **"Gmail API"** puis sur le bouton bleu **"Activer"** (Enable).

### 3. Configurer l'écran de consentement OAuth (Consent Screen)
1. Dans le menu de gauche, allez dans **"API et services"** > **"Écran de consentement OAuth"** (OAuth consent screen).
2. Choisissez le type d'utilisateur **"Externe"** (External), puis cliquez sur **"Créer"**.
3. Remplissez les champs obligatoires :
   - **Nom de l'application** : `Mail Sender PFE`
   - **Adresse e-mail d'assistance utilisateur** : votre adresse Gmail
   - **Coordonnées du développeur** : votre adresse Gmail
4. Cliquez sur **"Enregistrer et continuer"**.
5. À l'étape **"Champs d'application"** (Scopes), cliquez sur **"Ajouter ou supprimer des champs d'application"**, cherchez ou cochez :
   - `https://www.googleapis.com/auth/gmail.send` (ou saisissez manuellement ce scope).
   - Cliquez sur **"Mettre à jour"** puis **"Enregistrer et continuer"**.
6. ⚠️ **TRÈS IMPORTANT - Utilisateurs test (Test users)** :
   - Cliquez sur **"+ Add users"** (Ajouter des utilisateurs).
   - Saisissez **votre propre adresse Gmail** (celle avec laquelle vous allez envoyer les candidatures).
   - Cliquez sur **"Enregistrer et continuer"**.

### 4. Créer les identifiants OAuth (credentials.json)
1. Dans le menu de gauche, allez dans **"Identifiants"** (Credentials).
2. Cliquez sur **"+ Créer des identifiants"** en haut > **"ID client OAuth"** (OAuth client ID).
3. Type d'application : sélectionnez **"Application pour ordinateur"** (Desktop app).
4. Nom : `Desktop Client` (ou laissez par défaut).
5. Cliquez sur **"Créer"**.
6. Une fenêtre s'ouvre : cliquez sur **"Télécharger le fichier JSON"** (Download JSON).
7. Renommez le fichier téléchargé exactement en **`credentials.json`** et placez-le dans le dossier du projet :
   ```
   /home/chaabane/Bureau/Stage Et/mail & strategy/credentials.json
   ```

---

## 🧪 Étape 2 : Tester avant d'envoyer aux entreprises

Le script intègre deux niveaux de sécurité pour vérifier parfaitement le résultat avant tout envoi officiel.

### Test 1 : Prévisualisation visuelle (sans aucun envoi)
Cette commande lit le document Word, charge les contacts CSV, remplace les balises `[]` par les vrais noms de recruteurs et d'entreprises, et vérifie que les 2 PDF existent :

```bash
python3 send_applications.py --preview
```

Vous pouvez aussi prévisualiser plus de contacts en ajoutant `--limit 10` :
```bash
python3 send_applications.py --preview --limit 10
```

### Test 2 : Envoi d'un email test réel dans votre propre boîte
Cette commande effectue l'authentification Gmail et vous envoie **1 seul email réel** à votre adresse de test avec les 2 PDF attachés (CV + Lettre de motivation). Vous pourrez ainsi ouvrir l'email sur votre téléphone ou PC et vérifier exactement la mise en page :

```bash
python3 send_applications.py --test-send votre_email@gmail.com
```

> **Note lors du premier envoi** : Une page web Google va s'ouvrir automatiquement dans votre navigateur pour vous demander d'autoriser l'application. Connectez-vous avec votre compte Gmail et cliquez sur **"Continuer"** / **"Autoriser"**. Un fichier `token.json` sera généré automatiquement et vous n'aurez plus jamais besoin de vous reconnecter.

---

## 🚀 Étape 3 : Lancer l'envoi officiel aux entreprises

Une fois le rendu test validé dans votre boîte de réception :

```bash
python3 send_applications.py --send
```

> Le script détecte automatiquement les deux formats de CSV (wave 1 et wave 2) et ignore les lignes sans adresse email valide. Pour lancer la 2ᵉ vague :
> `python3 send_applications.py --send --csv wave2_contacts.csv`

### Sécurités incluses :
1. **Confirmation requise** : Le script vous affiche le nombre total de recruteurs ciblés et vous demande de taper explicitement `OUI` avant de démarrer.
2. **Délai anti-spam** : Par défaut, une pause de 3 secondes est respectée entre chaque email (modifiable avec `--delay 5`).
3. **Anti-doublon / Reprise sur coupure** : Chaque email envoyé avec succès est tracé dans `sent_emails.csv`. Si le script est interrompu (ex. coupure internet), relancer simplement `python3 send_applications.py --send` ignorera automatiquement les entreprises déjà contactées !
4. **Envoi par lot (optionnel)** : Vous pouvez envoyer par tranche de 10 ou 20 candidatures avec `--limit` :
   ```bash
   python3 send_applications.py --send --limit 20
   ```

---

## ⚙️ Options avancées

| Option | Description | Exemple |
|---|---|---|
| `--preview` | Prévisualiser le texte et vérifier les fichiers | `python3 send_applications.py --preview` |
| `--test-send EMAIL` | Envoyer un email test complet à cette adresse | `python3 send_applications.py --test-send test@gmail.com` |
| `--send` | Lancer l'envoi réel aux entreprises | `python3 send_applications.py --send` |
| `--delay SECONDES` | Délai entre les envois (défaut : 3s) | `python3 send_applications.py --send --delay 4` |
| `--limit N` | Limiter à N emails | `python3 send_applications.py --send --limit 15` |
| `--subject "..."` | Personnaliser l'objet de l'email | `python3 send_applications.py --send --subject "Stage PFE – {company} – Chaabane Boussadia"` |
| `--yes` | Ignorer la confirmation interactive | `python3 send_applications.py --send --yes` |
