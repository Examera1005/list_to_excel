#!/usr/bin/env python3
import os
import shutil

# Fichiers de configuration
TEMPLATE_FILE = "template.xlsx"
LISTE_FILE = "eleves.txt"

def dupliquer_templates():
    if not os.path.exists(TEMPLATE_FILE):
        print(f"❌ Erreur : Le fichier template '{TEMPLATE_FILE}' est introuvable.")
        return

    if not os.path.exists(LISTE_FILE):
        print(f"❌ Erreur : Le fichier liste '{LISTE_FILE}' est introuvable.")
        return

    with open(LISTE_FILE, "r", encoding="utf-8") as f:
        eleves = [line.strip() for line in f if line.strip()]

    print(f"🚀 Duplication en cours pour {len(eleves)} élève(s)...")
    for eleve in eleves:
        # Formater le nom du fichier (remplace les espaces par des underscores)
        nom_fichier = f"{eleve.replace(' ', '_')}.xlsx"
        shutil.copyfile(TEMPLATE_FILE, nom_fichier)
        print(f"  ✅ Créé : {nom_fichier}")

    print("🎉 Terminé avec succès !")

if __name__ == "__main__":
    dupliquer_templates()
