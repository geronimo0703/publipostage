# ==============================================================================
# app.py
# Changelog :
#   ed1 : ajout d'un bouton "exit"
#   ed1 : ajour d'un help
#
# ==============================================================================
from flask import Flask, render_template, request, redirect, url_for, jsonify, session
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

try:
    with open(CONFIG_PATH, "r", encoding="utf-8") as f:
        config = yaml.safe_load(f)
except FileNotFoundError:
    print(f"Erreur : fichier de configuration introuvable ({CONFIG_PATH}).\n"
          "L'application ne peut pas démarrer.")
    sys.exit(1)

# scripts/ est maintenant importé comme un module Python (et non plus
# exécuté en subprocess), indispensable pour fonctionner dans le binaire
# PyInstaller où aucun interpréteur python3 n'est garanti sur la machine
# du volontaire.
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
DOC_TEMPLATES_DIR = DATA_DIR / config["paths"]["document_templates_dir"]
PDF_DIR = DATA_DIR / config["paths"]["pdf_dir"]
DOC_DIR = DATA_DIR / config["paths"]["doc_dir"]
LOGS_DIR = DATA_DIR / config["paths"]["logs_dir"]
DB_PATH = DATA_DIR / config["database"]["path"]
DB_TABLE = config["database"]["table_name"]

DEFAULT_TEMPLATE_NAME = config["default_files"]["template_file"]
DEFAULT_SEND_EMAILS = config["processing"].get("send_emails_default", False)

# Créer les répertoires nécessaires
for d in (CSV_DIR, DOC_TEMPLATES_DIR, PDF_DIR, DOC_DIR, LOGS_DIR, DB_PATH.parent):
    d.mkdir(parents=True, exist_ok=True)

# --------------------------------------------
# Identifiants SMTP (.env à côté des données utilisateur persistantes,
# pour que le volontaire puisse les configurer sans rouvrir le binaire)
# --------------------------------------------
load_dotenv(DATA_DIR / ".env")
SMTP_USER = os.getenv("SMTP_USER")
SMTP_PASSWORD = os.getenv("SMTP_PASSWORD")
# Diagnostic au démarrage
if not SMTP_USER or not SMTP_PASSWORD:
    print(f"⚠️  SMTP non configuré — fichier .env attendu ici : {DATA_DIR / '.env'}", flush=True)
else:
    print(f"✅ SMTP chargé pour : {SMTP_USER}", flush=True)
SMTP_SERVER = config.get("email", {}).get("smtp_server", "smtp.gmail.com")
SMTP_PORT = config.get("email", {}).get("smtp_port", 587)


def init_db():
    conn = sqlite3.connect(str(DB_PATH))
    c = conn.cursor()
    # DB_TABLE vient de notre propre config.yaml local (pas une entrée utilisateur),
    # on peut donc l'interpoler directement dans le nom de la table.
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
    )


@app.route('/list/csv')
def list_csv_files():
    files = sorted(f.name for f in CSV_DIR.glob("*.csv"))
    return jsonify(files)


@app.route('/list/templates')
def list_template_files():
    files = sorted(f.name for f in DOC_TEMPLATES_DIR.glob("*.docx"))
    return jsonify(files)


@app.route('/list/email_templates')
def list_email_templates():
    files = sorted(f.name for f in DOC_TEMPLATES_DIR.glob("*.html"))
    return jsonify(files)


def _save_uploaded_file(file_storage, target_dir: Path, allowed_extensions):
    """
    Sauvegarde un fichier uploadé dans target_dir après vérification de
    l'extension. Retourne (succès: bool, message: str).
    """
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
    success, message = _save_uploaded_file(request.files.get('file'), CSV_DIR, {'.csv'})
    return jsonify({"success": success, "message": message})


@app.route('/upload/template', methods=['POST'])
def upload_template():
    success, message = _save_uploaded_file(request.files.get('file'), DOC_TEMPLATES_DIR, {'.docx'})
    return jsonify({"success": success, "message": message})


@app.route('/upload/email_template', methods=['POST'])
def upload_email_template():
    success, message = _save_uploaded_file(request.files.get('file'), DOC_TEMPLATES_DIR, {'.html'})
    return jsonify({"success": success, "message": message})


@app.route('/upload', methods=['POST'])
def upload_files():
    csv_filename = request.form.get('csv')
    docx_filename = request.form.get('docx')
    send_emails = request.form.get('send_emails') == 'on'
    personalize = True
    email_template_filename = request.form.get('email_template') or None

    if not csv_filename or not docx_filename:
        return redirect(request.url)

    csv_path = CSV_DIR / csv_filename
    docx_path = DOC_TEMPLATES_DIR / docx_filename

    if not csv_path.exists() or not docx_path.exists():
        return redirect(request.url)

    if send_emails and not email_template_filename:
        return redirect(request.url)

    email_template_path = (
        DOC_TEMPLATES_DIR / email_template_filename if email_template_filename else None
    )

    conn = sqlite3.connect(str(DB_PATH))
    c = conn.cursor()
    c.execute(f'''
        INSERT INTO {DB_TABLE} (date, fichier_csv, template_docx, statut, logs)
        VALUES (?, ?, ?, ?, ?)
    ''', (datetime.now().strftime('%Y-%m-%d %H:%M:%S'), csv_filename, docx_filename, "En attente", ""))
    campaign_id = c.lastrowid
    conn.commit()
    conn.close()

    campaign_logs[campaign_id] = {"status": "En attente", "logs": []}
    threading.Thread(
        target=lancer_campagne,
        args=(csv_path, docx_path, send_emails, personalize, email_template_path, campaign_id),
        daemon=True,
    ).start()

    return redirect(url_for('index'))

def _update_campaign(campaign_id, statut, logs_lines):
    conn = sqlite3.connect(str(DB_PATH))
    c = conn.cursor()
    c.execute(
        f'UPDATE {DB_TABLE} SET statut = ?, logs = ? WHERE id = ?',
        (statut, "\n".join(logs_lines), campaign_id),
    )
    conn.commit()
    conn.close()


def lancer_campagne(csv_path, docx_path, send_emails, personalize, email_template_path, campaign_id):
    campaign_logs[campaign_id]["status"] = "En cours"
    _update_campaign(campaign_id, "En cours", [f"Début du traitement : {csv_path.name}, {docx_path.name}"])

    try:
        result = run_publipostage(
            csv_path=csv_path,
            template_path=docx_path,
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
    # Sélectionner uniquement les 3 dernières campagnes
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


def open_browser():
    webbrowser.open("http://127.0.0.1:5000")


if __name__ == '__main__':
    debug_mode = os.getenv("FLASK_DEBUG", "0") == "1"

    if not os.environ.get("WERKZEUG_RUN_MAIN"):
        threading.Timer(1.0, open_browser).start()

    app.run(host='127.0.0.1', port=5000, debug=debug_mode)
