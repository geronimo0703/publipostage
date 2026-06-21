#!/usr/bin/env bash
# Crée un raccourci Bureau pour Publipostage.
# À lancer une seule fois après extraction du zip, depuis le même dossier
# que l'exécutable publipostage-linux.
set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
EXE_PATH="$SCRIPT_DIR/publipostage-linux"
DESKTOP_DIR="$HOME/Desktop"
DESKTOP_FILE="$DESKTOP_DIR/Publipostage.desktop"

if [ ! -f "$EXE_PATH" ]; then
    echo "Erreur : exécutable introuvable à $EXE_PATH"
    echo "Assurez-vous que ce script reste dans le même dossier que publipostage-linux."
    exit 1
fi

chmod +x "$EXE_PATH"

mkdir -p "$DESKTOP_DIR"

cat > "$DESKTOP_FILE" <<EOF
[Desktop Entry]
Type=Application
Name=Publipostage
Comment=Outil de publipostage - Banque Alimentaire 22
Exec="$EXE_PATH"
Terminal=false
Categories=Utility;
EOF

chmod +x "$DESKTOP_FILE"

# Sur certains environnements (GNOME notamment), un fichier .desktop
# fraîchement créé est considéré "non fiable" tant qu'il n'est pas
# marqué comme exécutable de confiance. On tente de le marquer
# automatiquement ; si ça échoue, l'utilisateur devra faire
# clic droit > "Autoriser le lancement" une fois.
if command -v gio >/dev/null 2>&1; then
    gio set "$DESKTOP_FILE" metadata::trusted true 2>/dev/null || true
fi

echo "Raccourci créé sur le Bureau : $DESKTOP_FILE"
echo "Vous pouvez maintenant double-cliquer dessus pour lancer Publipostage."
