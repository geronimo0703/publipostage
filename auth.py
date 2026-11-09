"""
auth.py — Connexion par lien magique pour le publipostage BA22.

- Liste blanche d'adresses (table users), gérée par les admins via /admin/utilisateurs
- Lien valable 15 minutes, à usage unique (nonce stocké en SQLite)
- Limite de demandes de lien par adresse et par heure
- Toutes les pages de l'application sont protégées, sauf la connexion

Intégration dans app.py :
    from auth import init_auth
    init_auth(app)

Variables .env utilisées :
    SECRET_KEY      clé longue et aléatoire (python -c "import secrets; print(secrets.token_hex(32))")
    ADMINS          adresses admin séparées par des virgules (ajoutées automatiquement au démarrage)
    BREVO_API_KEY   clé API Brevo (envoi transactionnel)
    MAIL_FROM       adresse expéditeur, vérifiée dans Brevo
    MAIL_FROM_NAME  nom affiché de l'expéditeur (facultatif)
    COOKIE_SECURE   1 = HTTPS seulement, 0 = http (local), auto = selon la requête
                    (utile quand le même serveur répond en http sur le réseau
                    local ET en https via ngrok)
    ALLOWED_HOSTS   noms d'hôte acceptés, séparés par des virgules
                    (ex. ba220-20:8000,xxx.ngrok-free.app) ; vide = pas de contrôle
"""

import json
import os
import secrets
import sqlite3
import time
import urllib.error
import urllib.request
from datetime import timedelta

from flask import (Blueprint, abort, current_app, flash, g, redirect,
                   render_template_string, request, session, url_for)
from flask.sessions import SecureCookieSessionInterface
from itsdangerous import BadSignature, SignatureExpired, URLSafeTimedSerializer

auth_bp = Blueprint("auth", __name__)

DOMAIN = "@banquealimentaire.org"
LINK_MAX_AGE = 15 * 60          # durée de validité d'un lien (secondes)
MAX_LINKS_PER_HOUR = 5          # demandes de lien par adresse et par heure
PUBLIC_ENDPOINTS = {"auth.login", "auth.magic", "static"}

SCHEMA = """
CREATE TABLE IF NOT EXISTS users (
    email     TEXT PRIMARY KEY,
    is_admin  INTEGER NOT NULL DEFAULT 0,
    added_at  REAL NOT NULL
);
CREATE TABLE IF NOT EXISTS nonces (
    nonce       TEXT PRIMARY KEY,
    email       TEXT NOT NULL,
    created_at  REAL NOT NULL,
    used_at     REAL
);
"""


# ---------------------------------------------------------------- base SQLite

def _db():
    if "auth_db" not in g:
        g.auth_db = sqlite3.connect(current_app.config["AUTH_DB"])
        g.auth_db.row_factory = sqlite3.Row
    return g.auth_db


def _close_db(exc=None):
    db = g.pop("auth_db", None)
    if db is not None:
        db.close()


def _serializer():
    return URLSafeTimedSerializer(current_app.secret_key, salt="magic-link")


# ---------------------------------------------------------------- initialisation

def init_auth(app):
    if not app.secret_key:
        app.secret_key = os.environ["SECRET_KEY"]

    app.config.setdefault("AUTH_DB", os.path.join(app.root_path, "data", "auth.db"))
    os.makedirs(os.path.dirname(app.config["AUTH_DB"]), exist_ok=True)

    app.config.update(
        PERMANENT_SESSION_LIFETIME=timedelta(days=30),
        SESSION_COOKIE_HTTPONLY=True,
        SESSION_COOKIE_SAMESITE="Lax",
    )
    cookie_secure = os.getenv("COOKIE_SECURE", "1").strip().lower()
    if cookie_secure == "auto":
        app.session_interface = _AdaptiveSecureSession()
    else:
        app.config["SESSION_COOKIE_SECURE"] = cookie_secure == "1"

    allowed_hosts = {h.strip().lower() for h in os.getenv("ALLOWED_HOSTS", "").split(",") if h.strip()}

    with sqlite3.connect(app.config["AUTH_DB"]) as db:
        db.executescript(SCHEMA)
        for email in os.getenv("ADMINS", "").split(","):
            email = email.strip().lower()
            if email:
                db.execute(
                    "INSERT INTO users(email, is_admin, added_at) VALUES (?, 1, ?) "
                    "ON CONFLICT(email) DO UPDATE SET is_admin = 1",
                    (email, time.time()),
                )

    app.teardown_appcontext(_close_db)
    app.register_blueprint(auth_bp)

    @app.before_request
    def _require_login():
        # les liens magiques sont construits à partir du nom d'hôte de la
        # requête : on refuse tout hôte inconnu (en-tête Host falsifié)
        if allowed_hosts and request.host.lower() not in allowed_hosts:
            abort(400)
        g.user = None
        email = session.get("user")
        if email:
            row = _db().execute("SELECT * FROM users WHERE email = ?", (email,)).fetchone()
            if row:
                g.user = row
            else:                       # retiré de la liste blanche entre-temps
                session.clear()
        if g.user is None and request.endpoint not in PUBLIC_ENDPOINTS:
            if request.endpoint is None:            # 404 : on laisse Flask répondre
                return None
            if request.method == "GET":
                session["next"] = request.full_path
            return redirect(url_for("auth.login"))
        return None

    @app.context_processor
    def _inject_user():
        return {"current_user": g.get("user")}


class _AdaptiveSecureSession(SecureCookieSessionInterface):
    """Cookie « Secure » seulement quand la requête arrive en HTTPS."""
    def get_cookie_secure(self, app):
        return request.is_secure


def admin_required():
    if not g.user or not g.user["is_admin"]:
        abort(403)


# ---------------------------------------------------------------- envoi du mail

BREVO_URL = "https://api.brevo.com/v3/smtp/email"


def _send_link(email, link):
    """Envoie le lien magique via l'API transactionnelle Brevo."""
    text = (
        "Bonjour,\n\n"
        "Pour vous connecter à l'outil de publipostage, cliquez sur ce lien :\n\n"
        f"{link}\n\n"
        "Ce lien est valable 15 minutes et ne peut servir qu'une seule fois.\n"
        "Si vous n'avez rien demandé, ignorez simplement ce message.\n\n"
        "Banque Alimentaire des Côtes d'Armor"
    )
    html = (
        "<p>Bonjour,</p>"
        "<p>Pour vous connecter à l'outil de publipostage, cliquez sur ce lien :</p>"
        f'<p><a href="{link}">Me connecter au publipostage</a></p>'
        "<p>Ce lien est valable 15 minutes et ne peut servir qu'une seule fois.<br>"
        "Si vous n'avez rien demandé, ignorez simplement ce message.</p>"
        "<p>Banque Alimentaire des Côtes d'Armor</p>"
    )
    payload = {
        "sender": {"email": os.getenv("MAIL_FROM") or os.environ["SMTP_FROM"],
                   "name": os.getenv("MAIL_FROM_NAME", "Banque Alimentaire 22")},
        "to": [{"email": email}],
        "subject": "Connexion au publipostage BA22",
        "textContent": text,
        "htmlContent": html,
        "tags": ["publipostage-login"],
    }
    req = urllib.request.Request(
        BREVO_URL,
        data=json.dumps(payload).encode("utf-8"),
        headers={
            "api-key": os.environ["BREVO_API_KEY"],
            "Content-Type": "application/json",
            "Accept": "application/json",
        },
        method="POST",
    )
    try:
        with urllib.request.urlopen(req, timeout=15) as resp:
            if resp.status >= 300:
                raise RuntimeError(f"Brevo a répondu {resp.status}")
    except urllib.error.HTTPError as e:
        detail = e.read().decode("utf-8", errors="replace")
        raise RuntimeError(
            f"Brevo a répondu {e.code} : {detail} "
            f"(expéditeur utilisé : {payload['sender']['email']!r})"
        ) from None


# ---------------------------------------------------------------- connexion

@auth_bp.route("/login", methods=["GET", "POST"])
def login():
    if request.method == "GET":
        return render_template_string(LOGIN_HTML)

    email = request.form.get("email", "").strip().lower()
    db = _db()
    now = time.time()
    db.execute("DELETE FROM nonces WHERE created_at < ?", (now - 86400,))

    allowed = db.execute("SELECT 1 FROM users WHERE email = ?", (email,)).fetchone()
    recent = db.execute(
        "SELECT COUNT(*) FROM nonces WHERE email = ? AND created_at > ?",
        (email, now - 3600),
    ).fetchone()[0]

    if allowed and recent < MAX_LINKS_PER_HOUR:
        nonce = secrets.token_urlsafe(16)
        db.execute("INSERT INTO nonces(nonce, email, created_at) VALUES (?, ?, ?)",
                   (nonce, email, now))
        db.commit()
        token = _serializer().dumps({"email": email, "nonce": nonce})
        try:
            _send_link(email, url_for("auth.magic", token=token, _external=True))
        except Exception:
            current_app.logger.exception("Échec d'envoi du lien magique à %s", email)
            # un envoi raté ne doit pas compter dans la limite horaire
            db.execute("DELETE FROM nonces WHERE nonce = ?", (nonce,))
            db.commit()
    else:
        db.commit()

    # même réponse dans tous les cas : on ne révèle pas qui est autorisé
    return render_template_string(CHECK_MAIL_HTML, email=email)


@auth_bp.route("/auth/<token>")
def magic(token):
    try:
        data = _serializer().loads(token, max_age=LINK_MAX_AGE)
    except SignatureExpired:
        return render_template_string(ERROR_HTML, message="Ce lien a expiré."), 403
    except BadSignature:
        return render_template_string(ERROR_HTML, message="Ce lien n'est pas valide."), 403

    db = _db()
    cur = db.execute(
        "UPDATE nonces SET used_at = ? WHERE nonce = ? AND email = ? AND used_at IS NULL",
        (time.time(), data.get("nonce"), data.get("email")),
    )
    db.commit()
    if cur.rowcount != 1:
        return render_template_string(ERROR_HTML, message="Ce lien a déjà été utilisé."), 403

    if not db.execute("SELECT 1 FROM users WHERE email = ?", (data["email"],)).fetchone():
        abort(403)

    nxt = session.pop("next", None)
    session.clear()
    session.permanent = True
    session["user"] = data["email"]
    if not nxt or not nxt.startswith("/") or nxt.startswith("//"):
        nxt = "/"
    return redirect(nxt)


@auth_bp.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("auth.login"))


# ---------------------------------------------------------------- administration

@auth_bp.route("/admin/utilisateurs", methods=["GET", "POST"])
def admin_users():
    admin_required()
    db = _db()
    if request.method == "POST":
        email = request.form.get("email", "").strip().lower()
        is_admin = 1 if request.form.get("is_admin") else 0
        if not email.endswith(DOMAIN):
            flash(f"Seules les adresses {DOMAIN} sont acceptées.")
        else:
            db.execute(
                "INSERT INTO users(email, is_admin, added_at) VALUES (?, ?, ?) "
                "ON CONFLICT(email) DO UPDATE SET is_admin = excluded.is_admin",
                (email, is_admin, time.time()),
            )
            db.commit()
            flash(f"{email} ajouté.")
        return redirect(url_for("auth.admin_users"))

    users = db.execute("SELECT * FROM users ORDER BY email").fetchall()
    return render_template_string(ADMIN_HTML, users=users)


@auth_bp.post("/admin/utilisateurs/supprimer")
def admin_remove_user():
    admin_required()
    email = request.form.get("email", "").strip().lower()
    if email == g.user["email"]:
        flash("Vous ne pouvez pas vous retirer vous-même.")
    else:
        db = _db()
        db.execute("DELETE FROM users WHERE email = ?", (email,))
        db.execute("DELETE FROM nonces WHERE email = ?", (email,))
        db.commit()
        flash(f"{email} retiré.")
    return redirect(url_for("auth.admin_users"))


# ---------------------------------------------------------------- gabarits HTML
# Gabarits minimaux pour démarrer ; à déplacer dans templates/ et à habiller
# avec le style du reste de l'application.

_BASE_CSS = """
<style>
  body { font-family: system-ui, sans-serif; max-width: 34rem; margin: 3rem auto; padding: 0 1rem; }
  input[type=email] { width: 100%; padding: .6rem; font-size: 1rem; margin: .5rem 0 1rem; }
  button { padding: .6rem 1.2rem; font-size: 1rem; cursor: pointer; }
  table { width: 100%; border-collapse: collapse; } td, th { padding: .4rem; border-bottom: 1px solid #ddd; text-align: left; }
  .flash { background: #eef6ee; padding: .6rem; margin-bottom: 1rem; }
</style>
"""

LOGIN_HTML = _BASE_CSS + """
<h1>Publipostage BA22</h1>
<form method="post">
  <label for="email">Votre adresse @banquealimentaire.org</label>
  <input type="email" id="email" name="email" required autofocus>
  <button type="submit">Recevoir mon lien de connexion</button>
</form>
"""

CHECK_MAIL_HTML = _BASE_CSS + """
<h1>Vérifiez votre boîte mail</h1>
<p>Si l'adresse <strong>{{ email }}</strong> est autorisée, un lien de connexion
vient de lui être envoyé. Il est valable 15 minutes.</p>
<p><a href="{{ url_for('auth.login') }}">Revenir</a></p>
"""

ERROR_HTML = _BASE_CSS + """
<h1>Connexion impossible</h1>
<p>{{ message }}</p>
<p><a href="{{ url_for('auth.login') }}">Demander un nouveau lien</a></p>
"""

ADMIN_HTML = _BASE_CSS + """
<h1>Utilisateurs autorisés</h1>
{% for m in get_flashed_messages() %}<div class="flash">{{ m }}</div>{% endfor %}
<table>
  <tr><th>Adresse</th><th>Admin</th><th></th></tr>
  {% for u in users %}
  <tr>
    <td>{{ u.email }}</td>
    <td>{{ "oui" if u.is_admin else "" }}</td>
    <td>
      <form method="post" action="{{ url_for('auth.admin_remove_user') }}">
        <input type="hidden" name="email" value="{{ u.email }}">
        <button type="submit">Retirer</button>
      </form>
    </td>
  </tr>
  {% endfor %}
</table>
<h2>Ajouter</h2>
<form method="post">
  <input type="email" name="email" placeholder="fonction@banquealimentaire.org" required>
  <label><input type="checkbox" name="is_admin"> administrateur</label>
  <button type="submit">Ajouter</button>
</form>
<p><a href="/">Retour</a> · <a href="{{ url_for('auth.logout') }}">Se déconnecter</a></p>
"""
