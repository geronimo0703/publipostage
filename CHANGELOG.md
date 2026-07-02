# Changelog — Publipostage Banque Alimentaire 22


## [v0.2.22] - 2026-07-02
### Ajouté
- purge des logs ; on ne conserve que les 5 derniers logs et .csv
- dans le template .docx, on récupére le reply_to passédans le formulaire
- traitement du cas ou LibreOffice n'est pas présent

## [v0.2.21] - 2026-06-30
### Ajouté 
- Ajout message pour réimporter un/des fichier(s) si modifier par le user

# .....

## v7 — 2026-06-29
### Ajouté
- Champ **Reply-To** dans le formulaire "Configurer la campagne" : l'adresse saisie est transmise comme header MIME et affichée dans le corps du mail via `{{reply_to}}`
- Injection de `reply_to` dans `data` en mode **dry_run** (preview) pour cohérence avec l'envoi réel
- Date de modification affichée dans les listes déroulantes templates (document et email)
- Bandeau d'avertissement dans `index.html` : rappel de réimporter les templates modifiés

## v6 — 2026-06-xx
### Modifié
- Suppression totale de tkinter, options en ligne de commande
- Extraction de la logique dans `run_publipostage()`, fonction importable par `app.py`
- Support `.xlsx` en plus de `.csv`
- Modèle email accepte `.html` ou `.docx` (mammoth)
- `template_path` rendu optionnel (mail sans pièce jointe)

## v5 — 2026-06-xx
### Ajouté
- Personnalisation du corps des emails avec les données CSV (`{{colonne}}`)

## v4 — 2026-06-xx
### Ajouté
- Version initiale fonctionnelle
- Correction NameError sur `send_email()`
- Correction envoi multi-destinataires (séparateur `;` → liste)
