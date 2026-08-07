#!/usr/bin/env python3
"""
Générateur d'Application macOS (.app)
Crée une vraie application Mac double-cliquable dans le Finder sans ouvrir le terminal !
"""

import os
import sys
import shutil
import stat

APP_NAME = "Générateur Élèves.app"

def create_mac_app():
    script_dir = os.path.dirname(os.path.abspath(__file__))
    app_path = os.path.join(script_dir, APP_NAME)
    
    contents_dir = os.path.join(app_path, "Contents")
    macos_dir = os.path.join(contents_dir, "MacOS")
    resources_dir = os.path.join(contents_dir, "Resources")

    # Nettoyer l'ancienne app si elle existe
    if os.path.exists(app_path):
        shutil.rmtree(app_path)

    os.makedirs(macos_dir, exist_ok=True)
    os.makedirs(resources_dir, exist_ok=True)

    # 1. Créer le script d'exécution binaire (launcher)
    launcher_path = os.path.join(macos_dir, "launcher")
    launcher_content = f"""#!/bin/bash
# Script de lancement automatique macOS
DIR="$( cd "$( dirname "${{BASH_SOURCE[0]}}" )/../Resources" && pwd )"
cd "$DIR"

# Détecter le binaire python3
PYTHON_BIN="$(which python3)"
if [ -z "$PYTHON_BIN" ]; then
    PYTHON_BIN="/usr/bin/python3"
fi

# Exécuter l'application GUI sans terminal
"$PYTHON_BIN" app_gui.py > /dev/null 2>&1 &
"""
    with open(launcher_path, "w", encoding="utf-8") as f:
        f.write(launcher_content)

    # Rendre le launcher exécutable (chmod +x)
    st = os.stat(launcher_path)
    os.chmod(launcher_path, st.st_mode | stat.S_IXUSR | stat.S_IXGRP | stat.S_IXOTH)

    # 2. Créer le fichier Info.plist pour macOS
    plist_path = os.path.join(contents_dir, "Info.plist")
    plist_content = """<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0">
<dict>
    <key>CFBundleExecutable</key>
    <string>launcher</string>
    <key>CFBundleIconFile</key>
    <string>AppIcon</string>
    <key>CFBundleIdentifier</key>
    <string>com.examera.generateur-eleves</string>
    <key>CFBundleName</key>
    <string>Générateur Élèves</string>
    <key>CFBundlePackageType</key>
    <string>APPL</string>
    <key>CFBundleShortVersionString</key>
    <string>1.0</string>
    <key>LSMinimumSystemVersion</key>
    <string>10.13</string>
    <key>NSHighResolutionCapable</key>
    <true/>
</dict>
</plist>
"""
    with open(plist_path, "w", encoding="utf-8") as f:
        f.write(plist_content)

    # 3. Copier les fichiers sources du projet dans Resources
    files_to_copy = [
        "app_gui.py", "app_tui.py", "dupliquer.py", ".version",
        "template.xlsx", "liste_eleves_1M1.xlsx", "liste_eleves_1M2.xlsx"
    ]
    
    for fname in files_to_copy:
        src = os.path.join(script_dir, fname)
        if os.path.exists(src):
            shutil.copy2(src, os.path.join(resources_dir, fname))

    print("===============================================================")
    print(f"🎉 APPLICATION MAC CRÉÉE AVEC SUCCÈS : {APP_NAME}")
    print("===============================================================")
    print("Vous pouvez désormais double-cliquer dessus directement dans")
    print("le Finder pour lancer l'application sans ouvrir le terminal !")
    print(f"📍 Emplacement : {os.path.abspath(app_path)}")
    print("===============================================================")

if __name__ == "__main__":
    create_mac_app()
