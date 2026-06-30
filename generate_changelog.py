#!/usr/bin/env python3
"""
generate_changelog.py

Génère une section de CHANGELOG.md à partir des commits Git depuis le dernier tag
(ou depuis un commit/tag donné en argument).

Convention attendue dans les messages de commit (style Conventional Commits) :
    feat: ajout du champ reply_to
    fix: correction du bug preview en dry_run
    docs: mise à jour du README
    chore: nettoyage du requirements.txt

Si tes commits ne suivent pas cette convention, les commits sont simplement
listés dans une section "Autres".

Usage :
    python generate_changelog.py                  # depuis le dernier tag jusqu'à HEAD
    python generate_changelog.py v1.2.0            # depuis le tag v1.2.0 jusqu'à HEAD
    python generate_changelog.py v1.2.0 v1.3.0      # entre deux tags
    python generate_changelog.py --version 1.3.0 --write   # insère direct dans CHANGELOG.md
"""

import argparse
import subprocess
import sys
import re
from datetime import date
from pathlib import Path

# Mapping type de commit -> section du changelog (Keep a Changelog)
CATEGORIES = {
    "feat": "Ajouté",
    "add": "Ajouté",
    "fix": "Corrigé",
    "bug": "Corrigé",
    "change": "Modifié",
    "refactor": "Modifié",
    "perf": "Modifié",
    "remove": "Supprimé",
    "del": "Supprimé",
    "security": "Sécurité",
    "docs": "Documentation",
    "chore": "Divers",
}

CATEGORY_ORDER = ["Ajouté", "Modifié", "Corrigé", "Supprimé", "Sécurité", "Documentation", "Divers", "Autres"]

COMMIT_RE = re.compile(r"^(?P<type>\w+)(\(.+\))?:\s*(?P<msg>.+)$")


def run_git(args):
    result = subprocess.run(["git"] + args, capture_output=True, text=True)
    if result.returncode != 0:
        print(f"Erreur git: {result.stderr}", file=sys.stderr)
        sys.exit(1)
    return result.stdout.strip()


def get_last_tag():
    try:
        return run_git(["describe", "--tags", "--abbrev=0"])
    except SystemExit:
        return None


def get_commits(start, end="HEAD"):
    range_spec = f"{start}..{end}" if start else end
    log = run_git(["log", range_spec, "--pretty=format:%s"])
    return [line for line in log.splitlines() if line.strip()]


def categorize(commits):
    grouped = {cat: [] for cat in CATEGORY_ORDER}
    for commit in commits:
        match = COMMIT_RE.match(commit)
        if match:
            ctype = match.group("type").lower()
            msg = match.group("msg").strip()
            category = CATEGORIES.get(ctype, "Autres")
        else:
            msg = commit.strip()
            category = "Autres"
        grouped[category].append(msg)
    return grouped


def format_section(version, grouped, version_date=None):
    version_date = version_date or date.today().isoformat()
    lines = [f"## [{version}] - {version_date}"]
    has_content = False
    for category in CATEGORY_ORDER:
        items = grouped.get(category, [])
        if items:
            has_content = True
            lines.append(f"### {category}")
            for item in items:
                lines.append(f"- {item[0].upper() + item[1:]}")
            lines.append("")
    if not has_content:
        lines.append("_Aucun changement notable._")
        lines.append("")
    return "\n".join(lines)


def insert_into_changelog(section_text, changelog_path="CHANGELOG.md"):
    path = Path(changelog_path)
    if not path.exists():
        content = "# Changelog\n\n" + section_text
        path.write_text(content, encoding="utf-8")
        print(f"Fichier {changelog_path} créé.")
        return

    content = path.read_text(encoding="utf-8")
    # Insère juste après le titre "# Changelog" (et après une éventuelle section [Non publié])
    lines = content.splitlines(keepends=True)
    insert_idx = 0
    for i, line in enumerate(lines):
        if line.startswith("# "):
            insert_idx = i + 1
            break
    # Saute les lignes vides juste après le titre
    while insert_idx < len(lines) and lines[insert_idx].strip() == "":
        insert_idx += 1

    new_content = "".join(lines[:insert_idx]) + "\n" + section_text + "\n" + "".join(lines[insert_idx:])
    path.write_text(new_content, encoding="utf-8")
    print(f"Section insérée dans {changelog_path}.")


def main():
    parser = argparse.ArgumentParser(description="Génère une entrée CHANGELOG.md depuis les commits Git.")
    parser.add_argument("start", nargs="?", help="Tag/commit de départ (par défaut : dernier tag)")
    parser.add_argument("end", nargs="?", default="HEAD", help="Tag/commit de fin (par défaut : HEAD)")
    parser.add_argument("--version", default=None, help="Numéro de version à afficher (ex: 1.3.0)")
    parser.add_argument("--date", default=None, help="Date à afficher (par défaut : aujourd'hui)")
    parser.add_argument("--write", action="store_true", help="Insère directement le résultat dans CHANGELOG.md")
    parser.add_argument("--changelog", default="CHANGELOG.md", help="Chemin du fichier changelog")
    args = parser.parse_args()

    start = args.start or get_last_tag()
    commits = get_commits(start, args.end)

    if not commits:
        print("Aucun commit trouvé sur cette plage.", file=sys.stderr)
        sys.exit(0)

    grouped = categorize(commits)
    version_label = args.version or args.end
    section = format_section(version_label, grouped, args.date)

    print(section)

    if args.write:
        insert_into_changelog(section, args.changelog)


if __name__ == "__main__":
    main()
