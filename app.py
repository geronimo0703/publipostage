# ==============================================================================
# app.py
# Changelog :
#   ed1 : ajout d'un bouton "exit"
#   ed1 : ajout d'un help
#   ed2 : with_attachment : le template .docx (et le PDF généré) sont
#          optionnels — si la case est décochée dans le formulaire,
#          aucun document n'est généré ni joint à l'email.
#
# ==============================================================================
from flask import Flask, render_template, request, redirect, url_for, jsonify, session, send_file
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
app.secret_key = os.getenv("FLASK_SECRET_KEY", "publipostage-dev-key")

# --------------------------------------------
# Répertoire des données utilisateur (écriture, persistant)
# --------------------------------------------
DATA_DIR = Path(user_data_dir(APP_NAME, APP_AUTHOR))

CSV_DIR = DATA_DIR / config["paths"]["csv_dir"]
DOCUMENT_TEMPLATES_DIR = DATA_DIR / config["paths"]["document_templates_dir"]
MESSAGE_TEMPLATES_DIR = DATA_DIR / config["paths"]["message_templates_dir"]
PDF_DIR = DATA_DIR / config["paths"]["pdf_dir"]
DOC_DIR = DATA_DIR / config["paths"]["doc_dir"]
LOGS_DIR = DATA_DIR / config["paths"]["logs_dir"]
DB_PATH = DATA_DIR / config["database"]["path"]
DB_TABLE = config["database"]["table_name"]

DEFAULT_TEMPLATE_NAME = config["default_files"]["template_file"]
DEFAULT_SEND_EMAILS = config["processing"].get("send_emails_default", False)

for d in (CSV_DIR, DOCUMENT_TEMPLATES_DIR, MESSAGE_TEMPLATES_DIR, PDF_DIR, DOC_DIR, LOGS_DIR, DB_PATH.parent):
    d.mkdir(parents=True, exist_ok=True)

# --------------------------------------------
# Identifiants SMTP
# --------------------------------------------
load_dotenv(DATA_DIR / ".env", override=True)
SMTP_USER = os.getenv("SMTP_USER")
SMTP_PASSWORD = os.getenv("SMTP_PASSWORD")
SMTP_FROM = os.getenv("SMTP_FROM")
print(f"🔍 SMTP_USER = '{SMTP_USER}'", flush=True)
print(f"🔍 SMTP_FROM = '{SMTP_FROM}'", flush=True)
if not SMTP_USER or not SMTP_PASSWORD:
    print(f"⚠️  SMTP non configuré — fichier .env attendu ici : {DATA_DIR / '.env'}", flush=True)
else:
    print(f"✅ SMTP chargé pour : {SMTP_USER}", flush=True)
SMTP_SERVER = config.get("email", {}).get("smtp_server", "smtp.gmail.com")
SMTP_PORT = config.get("email", {}).get("smtp_port", 587)


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
    conn.commit()
    conn.close()


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

    csv_path = CSV_DIR / csv_filename if csv_filename else None
    docx_path = DOCUMENT_TEMPLATES_DIR / docx_filename if docx_filename else None
    email_template_path = MESSAGE_TEMPLATES_DIR / email_template_filename if email_template_filename else None
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
    }


# --------------------------------------------
# Routes
# --------------------------------------------

@app.route("/help")
def help_page():
    return render_template("help.html")


@app.route("/shutdown", methods=["POST"])
def shutdown():
    """Arrête le serveur Flask proprement."""
    os.kill(os.getpid(), signal.SIGTERM)
    return ("", 204)


@app.route('/')
def index():
    return render_template(
        'index.html',
        default_send_emails=DEFAULT_SEND_EMAILS,
        build_version=BUILD_VERSION,
    )

@app.route('/version')
def version():
    return jsonify({"build_version": BUILD_VERSION})


@app.route('/list/csv')
def list_csv_files():
    files = sorted(
        f.name for f in CSV_DIR.iterdir()
        if f.suffix.lower() in {'.csv', '.xlsx', '.xls'}
    )
    return jsonify(files)


@app.route('/list/templates')
def list_template_files():
    files = sorted(f.name for f in DOCUMENT_TEMPLATES_DIR.glob("*.docx"))
    return jsonify(files)


@app.route('/list/email_templates')
def list_email_templates():
    files = sorted(
        f.name for f in MESSAGE_TEMPLATES_DIR.iterdir()
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
        request.files.get('file'), CSV_DIR, {'.csv', '.xlsx', '.xls'}
    )
    if success:
        _purge_old_uploads(CSV_DIR, ("*.csv", "*.xlsx", "*.xls"), keep=5)
    return jsonify({"success": success, "message": message})


@app.route('/upload/template', methods=['POST'])
def upload_template():
    success, message = _save_uploaded_file(request.files.get('file'), DOCUMENT_TEMPLATES_DIR, {'.docx'})
    if success:
        _purge_old_uploads(DOCUMENT_TEMPLATES_DIR, ("*.docx",), keep=5)
    return jsonify({"success": success, "message": message})


@app.route('/upload/email_template', methods=['POST'])
def upload_email_template():
    success, message = _save_uploaded_file(request.files.get('file'), MESSAGE_TEMPLATES_DIR, {'.html', '.docx'})
    if success:
        _purge_old_uploads(MESSAGE_TEMPLATES_DIR, ("*.html", "*.docx"), keep=5)
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
        INSERT INTO {DB_TABLE} (date, fichier_csv, template_docx, statut, logs)
        VALUES (?, ?, ?, ?, ?)
    ''', (datetime.now().strftime('%Y-%m-%d %H:%M:%S'), p["csv_filename"], template_label, "En attente", ""))
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
            p["reply_to"],
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
            old.unlink(missing_ok=True)

def lancer_campagne(csv_path, docx_path, send_emails, personalize, email_template_path, campaign_id,reply_to=None):
    print(f"🔍 reply_to dans lancer_campagne = '{reply_to}'", flush=True)
    campaign_logs[campaign_id]["status"] = "En cours"
    label = docx_path.name if docx_path else "mail seul"
    _update_campaign(campaign_id, "En cours", [f"Début du traitement : {csv_path.name}, {label}"])

    try:
        result = run_publipostage(
            csv_path=csv_path,
            template_path=docx_path,        # None = mail sans pièce jointe
            pdf_dir=PDF_DIR,
            doc_dir=DOC_DIR,
            logs_dir=LOGS_DIR,
            send_emails=send_emails,
            personalize=personalize,
            email_template_path=email_template_path,
            smtp_user=SMTP_USER,
            smtp_password=SMTP_PASSWORD,
            smtp_server=SMTP_SERVER,
            smtp_port=SMTP_PORT,
            smtp_from=SMTP_FROM,
            reply_to=reply_to,
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
        _purge_old_logs(LOGS_DIR, keep=3)

@app.route('/logs/files')
def list_log_files():
    """Liste les fichiers log_*.txt et statuts_*.csv dans LOGS_DIR."""
    files = []
    for pattern in ("log_*.txt", "statuts_*.csv"):
        for f in sorted(LOGS_DIR.glob(pattern), key=lambda x: x.stat().st_mtime, reverse=True):
            files.append({"name": f.name, "size": f.stat().st_size})
    return jsonify(files)


@app.route('/logs/files/<filename>')
def download_log_file(filename):
    """Télécharge un fichier log ou statuts depuis LOGS_DIR."""
    safe = secure_filename(filename)
    path = LOGS_DIR / safe
    if not path.exists() or not (safe.startswith("log_") or safe.startswith("statuts_")):
        return "Fichier introuvable", 404
    return send_file(path, as_attachment=True)

@app.route('/logs/<int:campaign_id>')
def get_logs(campaign_id):
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


@app.route('/preview', methods=['POST'])
def preview():
    print(f"🔍 form complet = {dict(request.form)}", flush=True)
    p = _parse_form_params(request.form)

    result = run_publipostage(
        csv_path=p["csv_path"],
        template_path=p["docx_path"],       # None si pas de pièce jointe
        pdf_dir=PDF_DIR,
        doc_dir=DOC_DIR,
        logs_dir=LOGS_DIR,
        send_emails=p["send_emails"],
        personalize=True,
        email_template_path=p["email_template_path"],
        smtp_user=SMTP_USER,
        smtp_password=SMTP_PASSWORD,
        smtp_server=SMTP_SERVER,
        smtp_port=SMTP_PORT,
        smtp_from=SMTP_FROM,
        reply_to=p["reply_to"],
        dry_run=True,
    )

    session['preview_pdf'] = result.get("preview_pdf")
    print(f"🔍 form passé au template = {dict(request.form)}", flush=True)
    return render_template(
        'preview.html',
        previews=result.get("previews", []),
        form=request.form,
        send_emails=p["send_emails"],
        has_pdf=bool(result.get("preview_pdf")),
    )


@app.route('/preview/pdf')
def preview_pdf():
    pdf_path = session.get('preview_pdf')
    if not pdf_path or not Path(pdf_path).exists():
        return "PDF non disponible", 404
    return send_file(pdf_path, mimetype='application/pdf')


@app.route('/campaign/<int:campaign_id>')
def campaign_result(campaign_id):
    return render_template('campaign_result.html', campaign_id=campaign_id)


@app.route('/campaign/<int:campaign_id>/download-log')
def download_log(campaign_id):
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
    debug_mode = True #os.getenv("FLASK_DEBUG", "0") == "1"

    if not os.environ.get("WERKZEUG_RUN_MAIN"):
        threading.Timer(1.0, open_browser).start()

    app.run(host='127.0.0.1', port=5000, debug=debug_mode)
