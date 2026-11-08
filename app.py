# ==============================================================================
# app.py
# Changelog :
#   ed1 : ajout d'un bouton "exit"
#   ed1 : ajout d'un help
#   ed2 : with_attachment : le template .docx (et le PDF généré) sont
#          optionnels — si la case est décochée dans le formulaire,
#          aucun document n'est généré ni joint à l'email.
#   web1 : espaces séparés (shared/ + users/<adresse>/) si AUTH_ENABLED=1,
#          campagnes rattachées à leur compte, noms de fichiers contrôlés,
#          /shutdown désactivé et stats Brevo réservées aux admins sur le web.
#
# ==============================================================================
from flask import Flask, render_template, request, redirect, url_for, jsonify, session, send_file, g, abort
from werkzeug.utils import secure_filename

import os
import sqlite3
import yaml
import sys
from pathlib import Path
from platformdirs import user_data_dir
from datetime import datetime
from dotenv import load_dotenv
import webbrowser
import threading
import signal
import mammoth
import requests
import secrets
import re
from types import SimpleNamespace

APP_NAME = "Publipostage"
APP_AUTHOR = "BanqueAlimentaire22"

# --------------------------------------------
# Répertoire des ressources embarquées (lecture seule)
# Répertoire _MEIPASS
# --------------------------------------------
if hasattr(sys, '_MEIPASS'):
    RESOURCE_DIR = Path(sys._MEIPASS)
else:
    RESOURCE_DIR = Path(__file__).resolve().parent

CONFIG_PATH = RESOURCE_DIR / "config" / "config.yaml"

VERSION_PATH = RESOURCE_DIR / "VERSION"
try:
    BUILD_VERSION = VERSION_PATH.read_text(encoding="utf-8").strip()
except FileNotFoundError:
    BUILD_VERSION = "dev"  # exécution locale hors build (python app.py direct)

try:
    with open(CONFIG_PATH, "r", encoding="utf-8") as f:
        config = yaml.safe_load(f)
except FileNotFoundError:
    print(f"Erreur : fichier de configuration introuvable ({CONFIG_PATH}).\n"
          "L'application ne peut pas démarrer.")
    sys.exit(1)

sys.path.insert(0, str(RESOURCE_DIR / "scripts"))
from publipostage_stock_ed7 import run_publipostage  # noqa: E402

app = Flask(
    __name__,
    template_folder=str(RESOURCE_DIR / "templates"),
    static_folder=str(RESOURCE_DIR / "static"),
)
from werkzeug.middleware.proxy_fix import ProxyFix
app.wsgi_app = ProxyFix(app.wsgi_app, x_proto=1, x_host=1)

# --------------------------------------------
# Répertoire des données utilisateur (écriture, persistant)
# --------------------------------------------
DATA_DIR = Path(user_data_dir(APP_NAME, APP_AUTHOR))

DB_PATH = DATA_DIR / config["database"]["path"]
DB_TABLE = config["database"]["table_name"]
DB_PATH.parent.mkdir(parents=True, exist_ok=True)

DEFAULT_TEMPLATE_NAME = config["default_files"]["template_file"]
DEFAULT_SEND_EMAILS = config["processing"].get("send_emails_default", False)

# --------------------------------------------
# Identifiants SMTP
# --------------------------------------------
load_dotenv(DATA_DIR / ".env", override=False)
SMTP_USER = os.getenv("SMTP_USER")          # obsolète depuis l'API Brevo, conservé pour compatibilité
SMTP_PASSWORD = os.getenv("SMTP_PASSWORD")  # idem
SMTP_FROM = os.getenv("SMTP_FROM")
BREVO_API_KEY = os.getenv("BREVO_API_KEY")
if not BREVO_API_KEY or not SMTP_FROM:
    print(f"⚠️  Brevo non configuré (BREVO_API_KEY ou SMTP_FROM manquant) — .env attendu ici : {DATA_DIR / '.env'}", flush=True)
else:
    print(f"✅ Brevo configuré, expéditeur : {SMTP_FROM}", flush=True)
SMTP_SERVER = os.getenv("SMTP_SERVER") or config.get("email", {}).get("smtp_server", "smtp.gmail.com")
SMTP_PORT = int(os.getenv("SMTP_PORT") or config.get("email", {}).get("smtp_port", 587))

# Clé Flask : lue APRÈS le chargement du .env (sinon une clé aléatoire était
# générée à chaque démarrage sur le PC)
app.secret_key = os.getenv("FLASK_SECRET_KEY") or secrets.token_hex(32)

# --------------------------------------------
# Authentification (version web) — désactivable pour l'usage local sur PC
# --------------------------------------------
AUTH_ENABLED = os.getenv("AUTH_ENABLED", "1") == "1"
if AUTH_ENABLED:
    from auth import init_auth, admin_required
    init_auth(app)
else:
    def admin_required():
        return None


# --------------------------------------------
# Espaces de travail
#   - AUTH_ENABLED=0 (PC) : disposition historique, tout sous DATA_DIR
#   - AUTH_ENABLED=1 (web) : DATA_DIR/shared pour les modèles et pièces
#     jointes communs, DATA_DIR/users/<adresse> pour les CSV, PDF, DOCX
#     générés et logs de chaque compte
# --------------------------------------------
SHARED_KEEP = 50    # modèles partagés : purge large, ils servent à tous
USER_KEEP = 5

def _user_key():
    """Identifiant du compte connecté, utilisable comme nom de dossier."""
    user = g.get("user") if AUTH_ENABLED else None
    if not user:
        return None
    return re.sub(r"[^a-z0-9@._-]", "_", user["email"].lower())


def _dirs_for(user_key):
    """Dossiers de travail pour un compte (None = disposition PC historique)."""
    paths = config["paths"]
    if AUTH_ENABLED and user_key:
        shared = DATA_DIR / "shared"
        own = DATA_DIR / "users" / user_key
    else:
        shared = own = DATA_DIR
    d = SimpleNamespace(
        csv=own / paths["csv_dir"],
        pdf=own / paths["pdf_dir"],
        doc=own / paths["doc_dir"],
        logs=own / paths["logs_dir"],
        documents=shared / paths["document_templates_dir"],
        messages=shared / paths["message_templates_dir"],
        attachments=shared / paths["attachments_dir"],
    )
    for p in vars(d).values():
        p.mkdir(parents=True, exist_ok=True)
    return d


def dirs():
    """Dossiers de l'utilisateur de la requête en cours (créés si besoin)."""
    if "dirs" not in g:
        g.dirs = _dirs_for(_user_key())
    return g.dirs


def _child(directory: Path, filename):
    """Fichier `filename` DANS `directory` — refuse tout chemin (../, /...)."""
    if not filename:
        return None
    name = Path(filename).name
    if name != filename or name in {"", ".", ".."}:
        abort(400)
    return directory / name


def init_db():
    conn = sqlite3.connect(str(DB_PATH))
    c = conn.cursor()
    c.execute(f'''
        CREATE TABLE IF NOT EXISTS {DB_TABLE} (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            date TEXT NOT NULL,
            fichier_csv TEXT NOT NULL,
            template_docx TEXT NOT NULL,
            statut TEXT NOT NULL,
            logs TEXT
        )
    ''')
    cols = [r[1] for r in c.execute(f"PRAGMA table_info({DB_TABLE})")]
    if "utilisateur" not in cols:
        c.execute(f"ALTER TABLE {DB_TABLE} ADD COLUMN utilisateur TEXT")
    conn.commit()
    conn.close()


def _owns_campaign(campaign_id):
    """La campagne appartient-elle au compte connecté ? (toujours vrai sur PC)"""
    if not AUTH_ENABLED:
        return True
    conn = sqlite3.connect(str(DB_PATH))
    row = conn.execute(f"SELECT utilisateur FROM {DB_TABLE} WHERE id = ?", (campaign_id,)).fetchone()
    conn.close()
    return bool(row) and row[0] == _user_key()


init_db()
campaign_logs = {}


# --------------------------------------------
# Helpers
# --------------------------------------------

def _parse_form_params(form):
    print(f"🔍 form.keys() = {list(form.keys())}", flush=True)
    print(f"🔍 reply_to brut = '{form.get('reply_to')}'", flush=True)
    """
    Extrait et normalise les paramètres communs aux routes /upload et /preview.
    Retourne un dict prêt à l'emploi.
    """
    csv_filename = form.get('csv')
    send_emails = form.get('send_emails') == 'on'
    email_template_filename = form.get('email_template') or None

    # docx vide ("— Sans pièce jointe —") = mail sans pièce jointe
    docx_filename = form.get('docx') or None

        # Pièces jointes génériques (multi-sélection)
    attachment_filenames = form.getlist('attachments')
    d = dirs()
    attachment_paths = [_child(d.attachments, f) for f in attachment_filenames if f]
    print(f"🔍 attachment_filenames = {attachment_filenames}", flush=True)
    print(f"🔍 attachment_paths = {attachment_paths}", flush=True)

    csv_path = _child(d.csv, csv_filename)
    docx_path = _child(d.documents, docx_filename)
    email_template_path = _child(d.messages, email_template_filename)
    reply_to = form.get("reply_to") or None

    return {
        "csv_filename": csv_filename,
        "docx_filename": docx_filename,
        "send_emails": send_emails,
        "with_attachment": bool(docx_filename),  # déduit du choix dans la liste
        "email_template_filename": email_template_filename,
        "csv_path": csv_path,
        "docx_path": docx_path,
        "email_template_path": email_template_path,
        "reply_to": reply_to,
        "attachment_paths": attachment_paths,
    }


# --------------------------------------------
# Routes
# --------------------------------------------

@app.route("/help")
def help_page():
    return render_template("help.html")


@app.route("/shutdown", methods=["POST"])
def shutdown():
    """Arrête le serveur Flask proprement (usage local uniquement)."""
    if AUTH_ENABLED:
        abort(404)
    os.kill(os.getpid(), signal.SIGTERM)
    return ("", 204)


@app.route('/')
def index():
    return render_template(
        'index.html',
        default_send_emails=DEFAULT_SEND_EMAILS,
        build_version=BUILD_VERSION,
        auth_enabled=AUTH_ENABLED,
    )

@app.route('/version')
def version():
    return jsonify({"build_version": BUILD_VERSION})


@app.route('/list/csv')
def list_csv_files():
    files = sorted(
        f.name for f in dirs().csv.iterdir()
        if f.suffix.lower() in {'.csv', '.xlsx', '.xls'}
    )
    return jsonify(files)


@app.route('/list/templates')
def list_template_files():
    files = sorted(f.name for f in dirs().documents.glob("*.docx"))
    return jsonify(files)

@app.route('/list/attachments')
def list_attachments():
    files = sorted(
        f.name for f in dirs().attachments.iterdir()
        if f.suffix.lower() in {'.pdf', '.docx'}
    )
    return jsonify(files)


@app.route('/upload/attachment', methods=['POST'])
def upload_attachment():
    success, message = _save_uploaded_file(
        request.files.get('file'), dirs().attachments, {'.pdf', '.docx'}
    )
    if success:
        _purge_old_uploads(dirs().attachments, ("*.pdf", "*.docx"), keep=SHARED_KEEP)
    return jsonify({"success": success, "message": message})

@app.route('/delete/attachment/<filename>', methods=['POST'])
def delete_attachment(filename):
    path = _child(dirs().attachments, filename)
    if path.exists():
        path.unlink()
        return jsonify({"success": True, "message": f"{filename} supprimé"})
    return jsonify({"success": False, "message": "Fichier introuvable"})

@app.route('/delete/csv/<filename>', methods=['POST'])


def delete_csv(filename):
    path = _child(dirs().csv, filename)
    if path.exists():
        path.unlink()
        return jsonify({"success": True, "message": f"{filename} supprimé"})
    return jsonify({"success": False, "message": "Fichier introuvable"})

@app.route('/delete/template/<filename>', methods=['POST'])
def delete_template(filename):
    path = _child(dirs().documents, filename)
    if path.exists():
        path.unlink()
        return jsonify({"success": True, "message": f"{filename} supprimé"})
    return jsonify({"success": False, "message": "Fichier introuvable"})

@app.route('/delete/email_template/<filename>', methods=['POST'])
def delete_email_template(filename):
    path = _child(dirs().messages, filename)
    if path.exists():
        path.unlink()
        return jsonify({"success": True, "message": f"{filename} supprimé"})
    return jsonify({"success": False, "message": "Fichier introuvable"})

@app.route('/list/email_templates')
def list_email_templates():
    files = sorted(
        f.name for f in dirs().messages.iterdir()
        if f.suffix.lower() in {'.html', '.docx'}
    )
    return jsonify(files)


def _save_uploaded_file(file_storage, target_dir: Path, allowed_extensions):
    if not file_storage or file_storage.filename == "":
        return False, "Aucun fichier sélectionné"
    filename = secure_filename(file_storage.filename)
    ext = Path(filename).suffix.lower()
    if ext not in allowed_extensions:
        return False, f"Extension non autorisée ({ext}). Attendu : {', '.join(allowed_extensions)}"
    target_dir.mkdir(parents=True, exist_ok=True)
    file_storage.save(target_dir / filename)
    return True, f"Fichier '{filename}' importé avec succès"


@app.route('/upload/csv', methods=['POST'])
def upload_csv():
    success, message = _save_uploaded_file(
        request.files.get('file'), dirs().csv, {'.csv', '.xlsx', '.xls'}
    )
    if success:
        _purge_old_uploads(dirs().csv, ("*.csv", "*.xlsx", "*.xls"), keep=USER_KEEP)
    return jsonify({"success": success, "message": message})


@app.route('/upload/template', methods=['POST'])
def upload_template():
    success, message = _save_uploaded_file(request.files.get('file'), dirs().documents, {'.docx'})
    if success:
        _purge_old_uploads(dirs().documents, ("*.docx",), keep=SHARED_KEEP)
    return jsonify({"success": success, "message": message})


@app.route('/upload/email_template', methods=['POST'])
def upload_email_template():
    success, message = _save_uploaded_file(request.files.get('file'), dirs().messages, {'.html', '.docx'})
    if success:
        _purge_old_uploads(dirs().messages, ("*.html", "*.docx"), keep=SHARED_KEEP)
    return jsonify({"success": success, "message": message})


@app.route('/upload', methods=['POST'])
def upload_files():
    print(f"🔍 /upload form = {dict(request.form)}", flush=True)
    p = _parse_form_params(request.form)

    if not p["csv_filename"]:
        return redirect(request.url)
    if not p["csv_path"].exists():
        return redirect(request.url)
    if p["with_attachment"] and (not p["docx_path"] or not p["docx_path"].exists()):
        return redirect(request.url)
    if p["send_emails"] and not p["email_template_filename"]:
        return redirect(request.url)

    # Valeur affichée dans l'historique
    template_label = p["docx_filename"] or "(mail seul)"

    conn = sqlite3.connect(str(DB_PATH))
    c = conn.cursor()
    c.execute(f'''
        INSERT INTO {DB_TABLE} (date, fichier_csv, template_docx, statut, logs, utilisateur)
        VALUES (?, ?, ?, ?, ?, ?)
    ''', (datetime.now().strftime('%Y-%m-%d %H:%M:%S'), p["csv_filename"], template_label, "En attente", "", _user_key()))
    campaign_id = c.lastrowid
    conn.commit()
    conn.close()

    campaign_logs[campaign_id] = {"status": "En attente", "logs": []}
    threading.Thread(
        target=lancer_campagne,
        args=(
            p["csv_path"], p["docx_path"],
            p["send_emails"], True,
            p["email_template_path"], campaign_id,
            p["reply_to"], p["attachment_paths"], dirs(),
        ),
        daemon=True,
    ).start()

    return redirect(url_for('campaign_result', campaign_id=campaign_id))


def _update_campaign(campaign_id, statut, logs_lines):
    conn = sqlite3.connect(str(DB_PATH))
    c = conn.cursor()
    c.execute(
        f'UPDATE {DB_TABLE} SET statut = ?, logs = ? WHERE id = ?',
        (statut, "\n".join(logs_lines), campaign_id),
    )
    conn.commit()
    conn.close()

def _purge_old_uploads(directory: Path, patterns, keep: int = 5) -> None:
    """Ne garde que les `keep` fichiers les plus récents par pattern dans `directory`."""
    for pattern in patterns:
        files = sorted(directory.glob(pattern), key=lambda f: f.stat().st_mtime)
        for old in files[:-keep]:
            old.unlink(missing_ok=True)

def _purge_old_logs(logs_dir: Path, keep: int = 3) -> None:
    """Supprime les anciens fichiers log_*.txt et statuts_*.csv, ne garde que les `keep` plus récents de chaque."""
    for pattern in ("log_*.txt", "statuts_*.csv"):
        files = sorted(logs_dir.glob(pattern), key=lambda f: f.stat().st_mtime)
        for old in files[:-keep]:
            try:
                old.unlink(missing_ok=True)
            except Exception:
                pass  # fichier verrouillé ou permissions insuffisantes


def lancer_campagne(csv_path, docx_path, send_emails, personalize, email_template_path, campaign_id, reply_to=None, attachment_paths=None, d=None):
    print(f"🔍 reply_to dans lancer_campagne = '{reply_to}'", flush=True)
    campaign_logs[campaign_id]["status"] = "En cours"
    label = docx_path.name if docx_path else "mail seul"
    _update_campaign(campaign_id, "En cours", [f"Début du traitement : {csv_path.name}, {label}"])

    try:
        result = run_publipostage(
            csv_path=csv_path,
            template_path=docx_path,        # None = mail sans pièce jointe
            pdf_dir=d.pdf,
            doc_dir=d.doc,
            logs_dir=d.logs,
            send_emails=send_emails,
            personalize=personalize,
            email_template_path=email_template_path,
            smtp_user=SMTP_USER,
            smtp_password=SMTP_PASSWORD,
            smtp_server=SMTP_SERVER,
            smtp_port=SMTP_PORT,
            smtp_from=SMTP_FROM,
            reply_to=reply_to,
            api_key=BREVO_API_KEY,
            attachment_paths=attachment_paths or [],
        )
        statut = "Terminé" if result["success"] else "Erreur"
        campaign_logs[campaign_id]["status"] = statut
        campaign_logs[campaign_id]["logs"] = result["logs"]
        _update_campaign(campaign_id, statut, result["logs"])

    except Exception as e:
        logs = [f"Erreur : {e}"]
        campaign_logs[campaign_id]["status"] = "Erreur"
        campaign_logs[campaign_id]["logs"] = logs
        _update_campaign(campaign_id, "Erreur", logs)

    finally:                                       # on ne garde que les 3 dernières campagnes
        _purge_old_logs(d.logs, keep=3)

@app.route('/logs/files')
def list_log_files():
    """Liste les fichiers log_*.txt et statuts_*.csv du dossier de logs du compte."""
    files = []
    for pattern in ("log_*.txt", "statuts_*.csv"):
        for f in sorted(dirs().logs.glob(pattern), key=lambda x: x.stat().st_mtime, reverse=True):
            files.append({"name": f.name, "size": f.stat().st_size})
    return jsonify(files)


@app.route('/logs/files/<filename>')
def download_log_file(filename):
    """Télécharge un fichier log ou statuts depuis le dossier de logs du compte."""
    safe = secure_filename(filename)
    path = dirs().logs / safe
    if not path.exists() or not (safe.startswith("log_") or safe.startswith("statuts_")):
        return "Fichier introuvable", 404
    return send_file(path, as_attachment=True)

@app.route('/logs/<int:campaign_id>')
def get_logs(campaign_id):
    if not _owns_campaign(campaign_id):
        return jsonify({"logs": [], "status": "Inconnu"})
    conn = sqlite3.connect(str(DB_PATH))
    c = conn.cursor()
    c.execute(f'SELECT logs, statut FROM {DB_TABLE} WHERE id = ?', (campaign_id,))
    result = c.fetchone()
    conn.close()

    if result:
        logs, status = result
        return jsonify({"logs": logs.split('\n') if logs else [], "status": status})
    return jsonify({"logs": [], "status": "Inconnu"})


@app.route('/history')
def history():
    conn = sqlite3.connect(str(DB_PATH))
    c = conn.cursor()
    if AUTH_ENABLED:
        c.execute(f'SELECT * FROM {DB_TABLE} WHERE utilisateur = ? ORDER BY date DESC LIMIT 3', (_user_key(),))
    else:
        c.execute(f'SELECT * FROM {DB_TABLE} ORDER BY date DESC LIMIT 3')
    campagnes = c.fetchall()
    conn.close()
    return jsonify([{
        "id": row[0],
        "date": row[1],
        "fichier_csv": row[2],
        "template_docx": row[3],
        "statut": row[4]
    } for row in campagnes])

@app.route('/brevo/stats')
def brevo_stats():
    """Récupère les événements des emails transactionnels depuis l'API Brevo."""
    admin_required()   # tous les envois partent du même expéditeur : réservé aux admins
    if not BREVO_API_KEY:
        return jsonify({"error": "BREVO_API_KEY non configurée dans le .env"}), 400

    import requests as req
    try:
        sender = request.args.get("sender", "").strip()
        email_filter = request.args.get("email", "").strip()
        event_filter = request.args.get("event", "").strip()
        start_date = request.args.get("startDate", "").strip()
        end_date = request.args.get("endDate", "").strip()

        params = {"limit": 500, "sort": "desc"}
        if email_filter:
            params["email"] = email_filter
        if event_filter:
            params["event"] = event_filter
        if start_date:
            params["startDate"] = start_date
        if end_date:
            params["endDate"] = end_date

        response = req.get(
            "https://api.brevo.com/v3/smtp/statistics/events",
            headers={"api-key": BREVO_API_KEY, "accept": "application/json"},
            params=params,
            timeout=10,
        )
        response.raise_for_status()
        data = response.json()

        # Filtre sur l'expéditeur côté Python
        if sender and "events" in data:
            data["events"] = [
                e for e in data["events"]
                if sender.lower() in (e.get("from", "") or "").lower()
            ]

        return jsonify(data)
    except Exception as e:
        return jsonify({"error": str(e)}), 500


def _brevo_fetch_events(params, max_pages=40, page_size=500):
    """Récupère tous les événements Brevo en paginant (offset)."""
    events = []
    offset = 0
    for _ in range(max_pages):
        p = dict(params, limit=page_size, offset=offset, sort="desc")
        r = requests.get(
            "https://api.brevo.com/v3/smtp/statistics/events",
            headers={"api-key": BREVO_API_KEY, "accept": "application/json"},
            params=p,
            timeout=15,
        )
        r.raise_for_status()
        batch = r.json().get("events", [])
        if not batch:
            break
        events.extend(batch)
        offset += len(batch)
    return events


@app.route('/brevo/synthese')
def brevo_synthese():
    """Synthèse par expéditeur : nombre de messages uniques par type d'événement."""
    admin_required()
    if not BREVO_API_KEY:
        return jsonify({"error": "BREVO_API_KEY non configurée dans le .env"}), 400

    sender = request.args.get("sender", "").strip().lower()
    params = {}
    if request.args.get("startDate"):
        params["startDate"] = request.args["startDate"]
    if request.args.get("endDate"):
        params["endDate"] = request.args["endDate"]

    try:
        events = _brevo_fetch_events(params)
    except Exception as e:
        return jsonify({"error": str(e)}), 500

    par_expediteur = {}
    for e in events:
        exp = (e.get("from") or "—").lower()
        if sender and sender not in exp:
            continue
        cle = (e.get("messageId"), e.get("email"))
        par_expediteur.setdefault(exp, {}).setdefault(e.get("event", "?"), set()).add(cle)

    resultat = [
        {"from": exp, "counts": {ev: len(cles) for ev, cles in evs.items()}}
        for exp, evs in sorted(par_expediteur.items())
    ]
    return jsonify({"expediteurs": resultat, "total_events": len(events)})

@app.route('/preview', methods=['POST'])
def preview():
    print(f"🔍 form complet = {dict(request.form)}", flush=True)
    p = _parse_form_params(request.form)

    preview_count = int(request.form.get('preview_count', 3))

    result = run_publipostage(
        csv_path=p["csv_path"],
        template_path=p["docx_path"],       # None si pas de pièce jointe
        pdf_dir=dirs().pdf,
        doc_dir=dirs().doc,
        logs_dir=dirs().logs,
        send_emails=p["send_emails"],
        personalize=True,
        email_template_path=p["email_template_path"],
        smtp_user=SMTP_USER,
        smtp_password=SMTP_PASSWORD,
        smtp_server=SMTP_SERVER,
        smtp_port=SMTP_PORT,
        smtp_from=SMTP_FROM,
        reply_to=p["reply_to"],
        api_key=BREVO_API_KEY,
        dry_run=True,
        preview_count=preview_count,
    )

    session['preview_pdf'] = result.get("preview_pdf")
    print(f"🔍 form passé au template = {dict(request.form)}", flush=True)
    return render_template(
        'preview.html',
        previews=result.get("previews", []),
        form=request.form,
        send_emails=p["send_emails"],
        has_pdf=bool(result.get("preview_pdf")),
        attachment_filenames=[p.name for p in p["attachment_paths"]],
        total_count=result.get("total_count", len(result.get("previews", []))),
        logs=result.get("logs", []),
    )


@app.route('/preview/pdf')
def preview_pdf():
    pdf_path = session.get('preview_pdf')
    if not pdf_path or not Path(pdf_path).exists():
        return "PDF non disponible", 404
    if dirs().pdf.resolve() not in Path(pdf_path).resolve().parents:
        abort(403)
    return send_file(pdf_path, mimetype='application/pdf')


@app.route('/campaign/<int:campaign_id>')
def campaign_result(campaign_id):
    if not _owns_campaign(campaign_id):
        abort(404)
    return render_template('campaign_result.html', campaign_id=campaign_id)


@app.route('/campaign/<int:campaign_id>/download-log')
def download_log(campaign_id):
    if not _owns_campaign(campaign_id):
        abort(404)
    conn = sqlite3.connect(str(DB_PATH))
    c = conn.cursor()
    c.execute(f'SELECT logs FROM {DB_TABLE} WHERE id = ?', (campaign_id,))
    row = c.fetchone()
    conn.close()
    if not row:
        return "Campagne introuvable", 404

    from flask import Response
    content = row[0] or ""
    filename = f"log_campagne_{campaign_id}.txt"
    return Response(
        content,
        mimetype="text/plain",
        headers={"Content-Disposition": f"attachment; filename={filename}"}
    )


def open_browser():
    webbrowser.open("http://127.0.0.1:5000")


if __name__ == '__main__':
    debug_mode = os.getenv("FLASK_DEBUG", "0") == "1" #true

    if not os.environ.get("WERKZEUG_RUN_MAIN"):
        threading.Timer(1.0, open_browser).start()

    app.run(host='127.0.0.1', port=5000, debug=debug_mode, use_reloader=False)
