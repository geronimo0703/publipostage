#!/usr/bin/env python3
# coding: utf-8
# ==============================================================================
# publipostage_stock_ed7.py
# Changelog :
#   ed4 : version initiale fonctionnelle
#   ed4-fix1 : correction NameError 'e' non défini lors de l'échec send_email()
#   ed4-fix2 : correction envoi multi-destinataires (séparateur ";" → liste)
#   ed5 : ajout du choix utilisateur pour personnaliser le corps des emails
#          avec les données individuelles du CSV (placeholders {{colonne}})
#   ed6 : choix du modèle email via une liste déroulante tkinter (jamais
#          finalisé)
#   ed7 : suppression totale de tkinter, options en ligne de commande.
#       : extraction de toute la logique dans run_publipostage(), une
#          fonction importable directement par app.py (plus de
#          subprocess + "python3", indispensable pour le binaire
#          PyInstaller où aucun interpréteur Python n'est garanti sur
#          la machine du volontaire). main()/CLI devient un client
#          fin de cette fonction, conservé pour les tests en dev.
#          Suppression des print() de debug qui affichaient le mot de
#          passe SMTP en clair dans les logs.
#       : Pour le fichier .csv, possibilité d'avoir un .xslx à la place
#         => tester le fichier avec pd.read_csv
#       : pour le message, accepter .html ou .docx
#       : template_path (.docx document) rendu optionnel : si absent,
#         aucun PDF n'est généré ni joint à l'email (mail seul).
#       : traitement du cas windows office365
#       : on ne garde que les 5 derniers logs
#       : dans le template .docx, on récupére le reply_to introduit dans le formulaire
# ==============================================================================

import argparse
import smtplib
import subprocess
import sys
from datetime import datetime
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from pathlib import Path

import pandas as pd
from docx import Document
from jinja2 import Template


# --------------------------------------------------------------------------
# Fonctions utilitaires (inchangées dans leur logique)
# --------------------------------------------------------------------------

def convert_docx_to_pdf(docx_path: Path, pdf_path: Path, log) -> None:
    # 1. Sur Windows, on tente Word via COM (Office 365 suffit, pas besoin de LibreOffice)
    if sys.platform == "win32":
        try:
            import win32com.client

            word = win32com.client.Dispatch("Word.Application")
            word.Visible = False
            doc = word.Documents.Open(str(docx_path.resolve()))
            # 17 = wdFormatPDF
            doc.SaveAs(str(pdf_path.resolve()), FileFormat=17)
            doc.Close()
            word.Quit()

            if pdf_path.exists():
                log(f"📄 Fichier converti (Word) : {pdf_path.name}")
                return
            else:
                log(f"❌ Conversion PDF (Word) : fichier attendu introuvable ({pdf_path})")
                return
        except ImportError:
            log("⚠️ pywin32 non installé, tentative avec LibreOffice…")
        except Exception as e:
            log(f"⚠️ Échec conversion via Word ({e}), tentative avec LibreOffice…")

    # 2. Fallback LibreOffice (Windows sans Word, ou Linux/Mac)
    try:
        subprocess.run(
            [
                "libreoffice", "--headless", "--convert-to", "pdf",
                "--outdir", str(pdf_path.parent), str(docx_path),
            ],
            check=True,
        )
        generated_pdf = pdf_path.parent / (docx_path.stem + ".pdf")
        if generated_pdf != pdf_path and generated_pdf.exists():
            generated_pdf.rename(pdf_path)

        if pdf_path.exists():
            log(f"📄 Fichier converti (LibreOffice) : {pdf_path.name}")
        else:
            log(f"❌ Conversion PDF : fichier attendu introuvable ({pdf_path})")
    except subprocess.CalledProcessError as e:
        log(f"❌ Échec de la conversion : {e}")
    except FileNotFoundError:
        log("❌ Ni Word (COM) ni LibreOffice ne sont disponibles sur ce poste.")



def lire_modele_email(chemin_fichier: Path, data=None, personnaliser=False, log=print):
    if not chemin_fichier.exists():
        log(f"⚠️ Le fichier {chemin_fichier} est introuvable.")
        return None

    suffix = chemin_fichier.suffix.lower()

    if suffix == ".docx":
        import mammoth
        with open(chemin_fichier, "rb") as f:
            result = mammoth.convert_to_html(f)
        contenu = result.value
        if result.messages:
            for msg in result.messages:
                log(f"⚠️ mammoth : {msg}")
    elif suffix == ".html":
        contenu = chemin_fichier.read_text(encoding="utf-8")
    else:
        log(f"⚠️ Format non supporté pour le modèle email : {suffix}")
        return None

    if personnaliser and data:
        contenu = Template(contenu).render(data=data, **data)

    return contenu


def normaliser_destinataires(raw_email: str):
    """Convertit 'a@x.com;b@y.com' (ou avec des virgules) en liste propre."""
    if not raw_email:
        return []
    return [a.strip() for a in raw_email.replace(";", ",").split(",") if a.strip()]


def send_email(to_email_raw, subject, pdf_path, modele_path, smtp_config, data=None, personnaliser=False, log=print):
    smtp_user = smtp_config.get("user")
    smtp_from = smtp_config.get("from", smtp_user)
    smtp_password = smtp_config.get("password")
    smtp_server = smtp_config.get("server", "smtp.gmail.com")
    smtp_port = smtp_config.get("port", 587)
    print(f"🔍 smtp_from dans send_email = '{smtp_from}'", flush=True)

    if not smtp_user or not smtp_password:
        msg_err = "SMTP_USER / SMTP_PASSWORD manquants (vérifie le fichier .env)"
        log(f"❌ {msg_err}")
        return False, msg_err

    if data is None:
        data = {}
    data["reply_to"] = smtp_config.get("reply_to", "")

    corps = lire_modele_email(modele_path, data=data, personnaliser=personnaliser, log=log)
    if corps is None:
        msg_err = f"Modèle introuvable : '{modele_path}'"
        log(f"❌ Impossible d'envoyer l'email : {msg_err}")
        return False, msg_err

    recipients = normaliser_destinataires(to_email_raw)
    if not recipients:
        return False, "Aucune adresse email valide"

    msg = MIMEMultipart()
    msg["From"] = smtp_from
    msg["To"] = ", ".join(recipients)
    msg["Subject"] = subject
    reply_to = smtp_config.get("reply_to")
    print(f"🔍 reply_to dans send_email = '{reply_to}'", flush=True)
    if reply_to:
        msg["Reply-To"] = reply_to
    msg.attach(MIMEText(corps, "html"))

    # Pièce jointe PDF : uniquement si un chemin valide est fourni
    if pdf_path and Path(pdf_path).exists():
        with open(pdf_path, "rb") as f:
            part = MIMEBase("application", "octet-stream")
            part.set_payload(f.read())
            encoders.encode_base64(part)
            part.add_header("Content-Disposition", f"attachment; filename={Path(pdf_path).name}")
            msg.attach(part)
        log(f"📎 Pièce jointe : {Path(pdf_path).name}")
    else:
        log("📧 Envoi sans pièce jointe")

    try:
        with smtplib.SMTP(smtp_server, smtp_port) as server:
            server.starttls()
            server.login(smtp_user, smtp_password)
            server.sendmail(smtp_user, recipients, msg.as_string())
        log(f"✉️ E-mail envoyé à {', '.join(recipients)}")
        return True, ""
    except Exception as e:
        log(f"❌ Échec de l'envoi à {', '.join(recipients)}: {e}")
        return False, str(e)


# --------------------------------------------------------------------------
# Fonction principale, importable directement par app.py
# --------------------------------------------------------------------------

def run_publipostage(
    csv_path: Path,
    template_path: Path = None,       # ← optionnel : None = pas de PDF joint
    pdf_dir: Path = None,
    doc_dir: Path = None,
    logs_dir: Path = None,
    send_emails: bool = False,
    personalize: bool = False,
    email_template_path: Path = None,
    smtp_user: str = None,
    smtp_password: str = None,
    smtp_server: str = "smtp.gmail.com",
    smtp_port: int = 587,
    smtp_from: str = None,
    reply_to: str = None,
    dry_run: bool = False,
):
    """
    Exécute le traitement complet : génère les fiches docx/pdf à partir du
    CSV et du template (si fourni), envoie les emails si demandé, met à
    jour le CSV avec les statuts.

    template_path est optionnel : s'il est None, aucun document PDF n'est
    généré ni joint à l'email (mail personnalisé seul).

    Ne lève pas d'exception pour les erreurs "métier" (ligne en échec,
    email manquant...) : celles-ci sont juste loguées et le traitement
    continue sur les lignes suivantes. Lève en revanche une exception si
    le CSV ou le template sont illisibles (erreur fatale, à remonter à
    l'appelant).

    Retourne un dict :
        {
            "success": bool,       # False si une erreur fatale a stoppé le traitement
            "logs": [str, ...],    # toutes les lignes de log produites
            "log_file": str,       # chemin du fichier log généré
            "previews": [...],     # aperçus dry_run
            "preview_pdf": str,    # chemin du PDF de prévisualisation (ou None)
        }
    """
    csv_path = Path(csv_path)
    template_path = Path(template_path) if template_path else None
    pdf_dir = Path(pdf_dir) if pdf_dir else None
    doc_dir = Path(doc_dir) if doc_dir else None
    logs_dir = Path(logs_dir)
    email_template_path = Path(email_template_path) if email_template_path else None
    previews = []
    preview_pdf_path = None

    for d in filter(None, (pdf_dir, doc_dir, logs_dir)):
        d.mkdir(parents=True, exist_ok=True)

    smtp_config = {
        "user": smtp_user,
        "password": smtp_password,
        "server": smtp_server,
        "port": smtp_port,
        "from": smtp_from or smtp_user,
        "reply_to": reply_to,
    }

    logs = []
    log_filename = logs_dir / f"log_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"
    # En dry_run (preview), on ne persiste rien sur disque : le retour en
    # mémoire (result["logs"]) suffit pour /preview. Ça évite d'accumuler
    # un log_*.txt à chaque clic sur "Aperçu".
    log_file_handle = open(log_filename, "a", encoding="utf-8") if not dry_run else None


    def log(msg):
        """Ajoute une ligne au log en mémoire et, hors dry_run, au fichier de log."""
        logs.append(msg)
        if log_file_handle:
            log_file_handle.write(msg + "\n")
            log_file_handle.flush()

    try:
        log(f"Template .docx utilisé : {template_path or '(aucun — mail sans pièce jointe)'}")
        log(f"Fichier .csv utilisé : {csv_path}")
        if template_path:
            log(f"Dossier de sortie .docx : {doc_dir}")
            log(f"Dossier de sortie .pdf  : {pdf_dir}")
        log(f"Envoi des emails : {send_emails}")
        log(f"Personnalisation : {personalize}")
        log(f"Modèle email : {email_template_path}\n")

        # Lecture CSV ou Excel selon l'extension
        suffix = Path(csv_path).suffix.lower()
        if suffix == '.xlsx':
            try:
                df = pd.read_excel(csv_path, sheet_name=0, engine='openpyxl')
                log("=== COLONNES DU FICHIER EXCEL (feuille 1) ===")
            except Exception:
                df = pd.read_csv(csv_path, encoding='utf-8-sig', sep=None, engine='python')
                log("⚠️  Fichier .xlsx détecté comme CSV — privilégiez un vrai export Excel.")
                log("=== COLONNES DU FICHIER (détecté comme CSV) ===")
        elif suffix == '.xls':
            try:
                df = pd.read_excel(csv_path, sheet_name=0, engine='xlrd')
                log("=== COLONNES DU FICHIER EXCEL ancien format (feuille 1) ===")
            except Exception:
                df = pd.read_csv(csv_path, encoding='utf-8-sig', sep=None, engine='python')
                log("⚠️  Format .xls détecté — privilégiez .xlsx pour éviter les erreurs de compatibilité.")
                log("=== COLONNES DU FICHIER EXCEL ancien format (feuille 1) ===")
        else:
            df = pd.read_csv(csv_path, encoding='utf-8-sig', sep=None, engine='python')
            log("=== COLONNES DU CSV ===")
        log(str(df.columns.tolist()))

        for col in ("Statut", "Date_envoi"):
            if col not in df.columns:
                df[col] = ""
            df[col] = df[col].astype("object")

        for index, row in df.iterrows():
            data = row.to_dict()
            data["DATE_DU_JOUR"] = datetime.now().strftime("%d/%m/%Y")
            data["reply_to"] = smtp_config.get("reply_to", "")  # on prend le paramètre passé  dans le formulaire
            log(f"🔍 [ligne {index}] reply_to après écrasement = '{data['reply_to']}' (smtp_config.reply_to = '{smtp_config.get('reply_to')}')")

            if "Nom_fichier" not in data or pd.isna(data.get("Nom_fichier")):
                log(f"⚠️ Ligne {index} ignorée : colonne 'Nom_fichier' manquante ou vide")
                df.at[index, "Statut"] = "Erreur: Nom_fichier manquant"
                continue

            # --- Mode dry_run (prévisualisation) ---
            if dry_run:
                data["reply_to"] = smtp_config.get("reply_to", "")
                email_preview = lire_modele_email(
                    email_template_path, data=data,
                    personnaliser=personalize, log=log
                ) if email_template_path else None
                previews.append({
                    "nom":    data.get("Nom_fichier", "?"),
                    "email":  data.get("Email", "—"),
                    "sujet":  data.get("Sujet", "—"),
                    "apercu": email_preview if email_preview else "—",
                })

                # Générer le PDF de prévisualisation uniquement si un template est fourni
                if preview_pdf_path is None and template_path:
                    doc = Document(template_path)
                    for paragraph in doc.paragraphs:
                        for key, value in data.items():
                            if pd.notna(value):
                                clean_value = str(value).strip().replace('\n', ' ').replace('\r', '')
                                placeholder = f"{{{{{key}}}}}"
                                if placeholder in paragraph.text:
                                    paragraph.text = paragraph.text.replace(placeholder, clean_value)
                    for table in doc.tables:
                        for trow in table.rows:
                            for cell in trow.cells:
                                for paragraph in cell.paragraphs:
                                    for key, value in data.items():
                                        placeholder = f"{{{{{key}}}}}"
                                        if placeholder in paragraph.text:
                                            paragraph.text = paragraph.text.replace(placeholder, str(value))
                    doc.add_paragraph(f"\nFait à Lannion, le {data['DATE_DU_JOUR']}")

                    preview_docx = doc_dir / f"_preview_{data['Nom_fichier']}.docx"
                    doc.save(preview_docx)
                    preview_pdf = pdf_dir / f"_preview_{data['Nom_fichier']}.pdf"
                    convert_docx_to_pdf(preview_docx, preview_pdf, log=log)
                    preview_docx.unlink(missing_ok=True)
                    if preview_pdf.exists():
                        preview_pdf_path = preview_pdf

                continue

            # --- Mode normal ---

            # Génération docx/PDF uniquement si un template est fourni
            pdf_to_attach = None
            if template_path:
                doc = Document(template_path)

                for paragraph in doc.paragraphs:
                    for key, value in data.items():
                        if pd.notna(value):
                            clean_value = str(value).strip().replace('\n', ' ').replace('\r', '')
                            placeholder = f"{{{{{key}}}}}"
                            if placeholder in paragraph.text:
                                paragraph.text = paragraph.text.replace(placeholder, clean_value)

                for table in doc.tables:
                    for trow in table.rows:
                        for cell in trow.cells:
                            for paragraph in cell.paragraphs:
                                for key, value in data.items():
                                    placeholder = f"{{{{{key}}}}}"
                                    if placeholder in paragraph.text:
                                        paragraph.text = paragraph.text.replace(placeholder, str(value))
                                    else:
                                        placeholder_with_spaces = f"{{{{ {key} }}}}"
                                        if placeholder_with_spaces in paragraph.text:
                                            paragraph.text = paragraph.text.replace(placeholder_with_spaces, str(value))

                doc.add_paragraph(f"\nFait à Lannion, le {data['DATE_DU_JOUR']}")

                output_filename = f"fiche_PS_remplie_{data['Nom_fichier']}.docx"
                output_path = doc_dir / output_filename
                doc.save(output_path)
                log(f"Document Word généré pour {data['Nom_fichier']} : {output_path}")

                pdf_path = pdf_dir / f"fiche_PS_remplie_{data['Nom_fichier']}.pdf"
                convert_docx_to_pdf(output_path, pdf_path, log=log)
                pdf_to_attach = pdf_path
            else:
                log(f"📧 Pas de document PDF pour {data.get('Nom_fichier', index)} (mode mail seul)")

            log(f"✅ Traité : {data['Nom_fichier']}")

            if send_emails:
                email_to_raw = data.get("Email", "")
                client = data.get("Client", "")
                if email_to_raw:
                    subject = data.get("Sujet", "")
                    success, error_msg = send_email(
                        email_to_raw, subject,
                        pdf_to_attach,           # None si pas de template docx
                        email_template_path,
                        smtp_config=smtp_config, data=data, personnaliser=personalize, log=log,
                    )
                    df.at[index, "Statut"] = "Envoyé" if success else f"Erreur: {error_msg}"
                    df.at[index, "Date_envoi"] = (
                        datetime.now().strftime("%d/%m/%Y %H:%M:%S") if success else ""
                    )
                    if success:
                        log(f"✅ Mail bien envoyé à {client}")
                else:
                    log(f"⚠️ Aucune adresse e-mail trouvée pour {client}")
                    df.at[index, "Statut"] = "Email manquant"
            else:
                df.at[index, "Statut"] = "Non envoyé (mode test)"
            log("=====================================================================")

        if not dry_run:
            statuts_path = logs_dir / f"statuts_{Path(csv_path).stem}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"
            df.to_csv(statuts_path, index=False, encoding='utf-8-sig')
            log(f"✅ Statuts sauvegardés dans : {statuts_path.name}")
            log("✨ Tous les documents ont été traités.")

        return {
            "success": True,
            "logs": logs,
            "log_file": str(log_filename) if not dry_run else None,
            "previews": previews,
            "preview_pdf": str(preview_pdf_path) if preview_pdf_path else None,
        }

    except Exception as e:
        log(f"❌ Erreur fatale : {e}")
        return {"success": False, "logs": logs, "log_file": str(log_filename)}

    finally:
        if log_file_handle:
            log_file_handle.close()


# --------------------------------------------------------------------------
# Client en ligne de commande, pour usage/tests en dev uniquement.
# app.py n'appelle plus ce script via subprocess, il importe directement
# run_publipostage() ci-dessus.
# --------------------------------------------------------------------------

def parse_args():
    parser = argparse.ArgumentParser(
        description="Génère les fiches PS (docx + pdf) et envoie les emails associés."
    )
    parser.add_argument("--csv", required=True)
    parser.add_argument("--template")   # optionnel : pas de PDF si absent
    parser.add_argument("--send-emails", action="store_true")
    parser.add_argument("--personalize", action="store_true")
    parser.add_argument("--email-template")
    args = parser.parse_args()
    if args.send_emails and not args.email_template:
        parser.error("--email-template est requis quand --send-emails est utilisé")
    return args


def main():
    import os
    import yaml
    from dotenv import load_dotenv

    args = parse_args()

    base_dir = Path(__file__).resolve().parent.parent
    with open(base_dir / "config" / "config.yaml", "r", encoding="utf-8") as f:
        config = yaml.safe_load(f)
    load_dotenv(base_dir / ".env")

    result = run_publipostage(
        csv_path=args.csv,
        template_path=args.template,   # peut être None
        pdf_dir=base_dir / config["paths"]["pdf_dir"],
        doc_dir=base_dir / config["paths"]["doc_dir"],
        logs_dir=base_dir / config["paths"]["logs_dir"],
        send_emails=args.send_emails,
        personalize=args.personalize,
        email_template_path=args.email_template,
        smtp_user=os.getenv("SMTP_USER"),
        smtp_password=os.getenv("SMTP_PASSWORD"),
        smtp_server=config.get("email", {}).get("smtp_server", "smtp.gmail.com"),
        smtp_port=config.get("email", {}).get("smtp_port", 587),
    )

    print("\n".join(result["logs"]))
    sys.exit(0 if result["success"] else 1)


if __name__ == "__main__":
    main()
