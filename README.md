# Publipostage — Banque Alimentaire 22

Outil de publipostage : génère des fiches PS (docx + pdf) à partir d'un CSV
de bénéficiaires et d'un template Word, avec envoi d'emails optionnel.
Packagé en exécutable autonome (PyInstaller) pour Windows et Linux —
aucune installation de Python requise côté volontaire.

## Sommaire

- [Architecture](#architecture)
- [Développement local](#développement-local)
- [Build manuel (local)](#build-manuel-local)
- [Build automatique (GitHub Actions)](#build-automatique-github-actions)
- [Où vivent les données](#où-vivent-les-données)
- [Dépannage](#dépannage)
- [Installation](#Installation pour les volontaires)

---

## Architecture
Publipostage_comptable/
├── app.py                      # Serveur Flask, point d'entrée
├── publipostage.spec           # Config PyInstaller
├── requirements.txt
├── config/
│   ├── config.yaml             # Config applicative (chemins, table DB...)
│   └── .env.default            # Généré par la CI uniquement, jamais committé
├── scripts/
│   └── publipostage_stock_ed7.py  # Logique métier, importée par app.py
├── templates/
│   └── index.html              # Interface web
├── .github/
│   └── workflows/
│       └── build.yml            # CI : build Windows + Ubuntu sur tag
└── .github/workflows/
    └── build.yml              
├── installer/
│   ├── creer_raccourci_linux.sh
│   └── creer_raccourci_windows.ps1
```

Points clés à retenir si tu reprends ce projet plus tard :

- **`run_publipostage()`** (dans `scripts/publipostage_stock_ed7.py`) est
  importée directement par `app.py` — pas de `subprocess`. C'est volontaire :
  un exécutable PyInstaller ne peut pas garantir qu'un `python3` existe sur
  la machine cible.
- **`RESOURCE_DIR`** (ressources en lecture seule, embarquées dans le binaire)
  et **`DATA_DIR`** (données utilisateur, persistantes, gérées par
  `platformdirs`) sont deux notions séparées. Ne jamais écrire dans
  `RESOURCE_DIR` à l'exécution.
- Les identifiants SMTP vivent dans `DATA_DIR/.env`, pré-remplis
  automatiquement au premier lancement depuis `config/.env.default`
  (embarqué dans le binaire, généré par la CI à partir des secrets GitHub).
  Le volontaire n'a jamais besoin d'y toucher.

---

## Développement local

```bash
cd Publipostage_comptable
python3 -m venv venv-dev
source venv-dev/bin/activate
pip install -r requirements.txt

cp .env.example .env   # puis remplir SMTP_USER / SMTP_PASSWORD pour tester
python3 app.py
```

L'app s'ouvre automatiquement sur `http://127.0.0.1:5000`.

En dev, `app.py` détecte l'absence de `sys._MEIPASS` et utilise le dossier
du projet comme `RESOURCE_DIR` — pas besoin de build pour itérer.

---

## Build manuel (local)

Toujours utiliser un **venv dédié au build**, séparé du venv de dev, pour
être sûr que `requirements.txt` est exhaustif (sinon des libs installées
"par accident" en dev masquent des dépendances manquantes pour le binaire).

```bash
cd Publipostage_comptable
python3 -m venv venv-build
source venv-build/bin/activate
pip install -r requirements.txt
pip install pyinstaller

# Build propre (supprime le cache PyInstaller au besoin)
rm -rf build dist
pyinstaller publipostage.spec
```

Le binaire sort dans `dist/publipostage` (Linux) ou `dist/publipostage.exe`
(Windows, si buildé sur Windows — PyInstaller ne fait pas de cross-compilation).

### Tester le binaire

```bash
./dist/publipostage
```

Vérifier :
- le navigateur s'ouvre sur `http://127.0.0.1:5000`
- `/list/csv`, `/list/templates` répondent sans erreur
- une vraie campagne de test passe (CSV + docx factices) — valide que
  `run_publipostage` s'exécute bien en mode importé

Pour un test plus poussé (machine "propre", sans l'environnement de dev) :

```bash
docker run --rm -v $(pwd)/dist:/test -it ubuntu:24.04 /test/publipostage
```

### Localiser le dossier de données utilisateur

```bash
python3 -c "from platformdirs import user_data_dir; print(user_data_dir('Publipostage', 'BanqueAlimentaire22'))"
```

---

## Build automatique (GitHub Actions)

Le workflow `.github/workflows/build.yml` se déclenche **uniquement sur un
tag de version** (`v*`) — pas à chaque commit.

### Secrets requis (une seule fois, par repo)

**Settings → Secrets and variables → Actions → New repository secret** :

| Nom | Exemple |
|---|---|
| `SMTP_USER` | `ba220.informatique@banquealimentaire.org` |
| `SMTP_PASSWORD` | mot de passe d'application Gmail (16 caractères) |
| `SMTP_SERVER` | `smtp.gmail.com` |
| `SMTP_PORT` | `587` |

Vérifier aussi : **Settings → Actions → General → Workflow permissions**
doit être sur **"Read and write permissions"** (nécessaire pour que la CI
puisse créer la Release).

### Déclencher un build + release

```bash
git add .
git commit -m "..."
git push

git tag v1.0.0
git push origin v1.0.0
```

Suivre l'avancement dans l'onglet **Actions** du repo. À la fin, une
**Release** apparaît avec `publipostage-linux` et `publipostage-windows.exe`
attachés, prêts à être distribués aux volontaires.

### Obtenir un mot de passe d'application Gmail (si besoin d'en regénérer un)

Prérequis : validation en 2 étapes activée sur le compte Google.

1. https://myaccount.google.com/apppasswords
2. Sélectionner une application → "Autre (Nom personnalisé)" → nommer
   "Publipostage BA22"
3. Générer, copier le code de 16 caractères immédiatement (non récupérable
   ensuite)
4. Mettre à jour le secret `SMTP_PASSWORD` sur GitHub, puis retag une
   nouvelle version

---

## Où vivent les données

| Type | Emplacement | Géré par |
|---|---|---|
| Code, templates, config par défaut | `_MEIPASS` (lecture seule, dans le binaire) | PyInstaller |
| CSV, templates docx/html importés, PDF générés, logs, base SQLite, `.env` | `DATA_DIR` (`platformdirs`) | écrits à l'exécution |

Le volontaire **n'a jamais besoin de connaître ces chemins** — tout passe
par l'interface web (upload de fichiers, config SMTP pré-remplie).

> ⚠️ **Exception : le fichier `.env`** doit être créé manuellement dans `DATA_DIR` lors de la première installation sur chaque machine :
>
> - **Linux** : `~/.local/share/BanqueAlimentaire22/Publipostage/.env`
> - **Windows** : `C:\Users\<user>\AppData\Local\Publipostage\.env`
>
> Contenu attendu :
> ```
> SMTP_USER=ba220.informatique@banquealimentaire.org
> SMTP_PASSWORD=<mot de passe d'application Gmail>
> SMTP_FROM=ba220.benevoles@banquealimentaire.org
> FLASK_SECRET_KEY=<clé aléatoire>
> ```

---

## Dépannage

| Symptôme | Cause probable | Solution |
|---|---|---|
| `ModuleNotFoundError` au lancement du binaire | Dépendance absente de `requirements.txt` ou non détectée par l'analyse statique PyInstaller | `pip install <module>` dans le venv-build, l'ajouter à `requirements.txt`, éventuellement à `hiddenimports` dans le `.spec` |
| `Unable to find '.../static'` au build | Dossier listé dans `datas` du `.spec` mais absent du disque | Retirer la ligne du `.spec` si le dossier n'existe pas vraiment |
| `Invalid value '...' for dtype 'float64'` | Colonne CSV vide lue par pandas comme `float64`, puis écriture d'une chaîne dedans | Déjà corrigé dans `run_publipostage()` via `.astype("object")` sur `Statut`/`Date_envoi` |
| Email non envoyé, pas d'erreur visible | `.env` absent ou mal placé dans `DATA_DIR` | Vérifier via le bouton "Tester la connexion" de l'interface |
| `Resource not accessible by integration` sur la release CI | Permissions par défaut du `GITHUB_TOKEN` trop restrictives | `permissions: contents: write` sur le job `release` (déjà en place) + vérifier les Workflow permissions du repo |
| Conversion PDF échoue | LibreOffice absent du PATH sur la machine du volontaire | Prérequis à documenter dans le guide d'installation : LibreOffice doit être installé séparément |

## Installation pour les volontaires

Après avoir téléchargé et extrait le zip correspondant à votre système :

### Linux

1. Extraire l'archive `publipostage-linux.zip`
2. Ouvrir un terminal dans le dossier extrait
3. Lancer le script de configuration :
```bash
   ./creer_raccourci_linux.sh
```
4. Une icône **Publipostage** apparaît sur le Bureau. Double-cliquez dessus pour lancer l'application.

> Si le raccourci ne se lance pas au premier double-clic (message "non fiable"), faites un clic droit dessus puis choisissez *Autoriser le lancement* (ou *Trust/Launch*, selon votre environnement de bureau).

### Windows

1. Extraire l'archive `publipostage-windows.zip`
2. Faire un clic droit sur `creer_raccourci_windows.ps1` → **Exécuter avec PowerShell**
3. Une icône **Publipostage** apparaît sur le Bureau. Double-cliquez dessus pour lancer l'application.

> Si Windows refuse l'exécution du script (fichier bloqué car téléchargé d'Internet) :
> - clic droit sur le fichier → **Propriétés** → cocher **Débloquer** → OK, puis relancer,
> - ou en PowerShell : `Unblock-File .\creer_raccourci_windows.ps1`

### Notes techniques

- Les scripts de raccourci (`installer/creer_raccourci_linux.sh` et `installer/creer_raccourci_windows.ps1`) sont inclus automatiquement dans chaque paquet par la CI (`.github/workflows/build.yml`).
- Ils utilisent un chemin **relatif à leur propre emplacement** : le dossier extrait peut être déplacé n'importe où, le raccourci pointera toujours vers l'exécutable correspondant.
- Le raccourci n'a besoin d'être créé **qu'une seule fois** par installation ; il reste valide tant que le dossier extrait n'est pas déplacé ou supprimé.
