#!/usr/bin/env python3
"""
Générateur de Fichiers Élèves - Interface Terminal (TUI)
Compatible macOS / Linux / Windows
"""

import os
import sys
import shutil
import glob
import subprocess

def clear_screen():
    os.system('cls' if os.name == 'nt' else 'clear')

def get_native_file_picker(title="Choisir un fichier", file_types=""):
    """Tente d'ouvrir la fenêtre native macOS pour choisir un fichier."""
    if sys.platform == "darwin":
        try:
            cmd = f'osascript -e "POSIX path of (choose file with prompt \\"{title}\\")"'
            res = subprocess.check_output(cmd, shell=True).decode('utf-8').strip()
            if res:
                return res
        except Exception:
            pass
    return None

def get_native_folder_picker(title="Choisir un dossier"):
    """Tente d'ouvrir la fenêtre native macOS pour choisir un dossier."""
    if sys.platform == "darwin":
        try:
            cmd = f'osascript -e "POSIX path of (choose folder with prompt \\"{title}\\")"'
            res = subprocess.check_output(cmd, shell=True).decode('utf-8').strip()
            if res:
                return res
        except Exception:
            pass
    return None

def main():
    clear_screen()
    print("===============================================================")
    print("  🎓 GÉNÉRATEUR DE DOCUMENTS ÉLÈVES (Template Excel)")
    print("===============================================================\n")

    # -------------------------------------------------------------------------
    # 1. SÉLECTION DU FICHIER TEMPLATE (.xlsx)
    # -------------------------------------------------------------------------
    print("📌 1. FICHIER TEMPLATE EXCEL")
    templates_locaux = glob.glob("*.xlsx")
    template_path = ""

    if templates_locaux:
        print("Fichiers Excel trouvés dans le dossier actuel :")
        for idx, f in enumerate(templates_locaux, 1):
            print(f"  [{idx}] {f}")
        print(f"  [M] Ouvrir le sélecteur de fichier macOS")
        print(f"  [S] Saisir le chemin manuellement")
        
        choix = input("\n👉 Choix (défaut=1) : ").strip()
        if choix.upper() == 'M':
            template_path = get_native_file_picker("Sélectionnez votre fichier template Excel")
        elif choix.upper() == 'S':
            template_path = input("Chemin du fichier template : ").strip()
        else:
            try:
                idx_sel = int(choix) - 1 if choix else 0
                template_path = templates_locaux[idx_sel]
            except (ValueError, IndexError):
                template_path = templates_locaux[0]
    else:
        print("Aucun fichier .xlsx dans le dossier courant.")
        if sys.platform == "darwin":
            print("Ouverture du sélecteur macOS...")
            template_path = get_native_file_picker("Sélectionnez votre fichier template Excel")
        
        if not template_path:
            template_path = input("👉 Entrez le chemin du fichier template.xlsx : ").strip()

    if not template_path or not os.path.exists(template_path):
        print(f"\n❌ Fichier template introuvable : '{template_path}'")
        sys.exit(1)

    print(f"✅ Template sélectionné : {template_path}\n")

    # -------------------------------------------------------------------------
    # 2. SÉLECTION DE LA LISTE DES ÉLÈVES (.txt)
    # -------------------------------------------------------------------------
    print("---------------------------------------------------------------")
    print("📌 2. LISTE DES ÉLÈVES")
    txt_locaux = glob.glob("*.txt")
    liste_path = ""

    if txt_locaux:
        print("Fichiers texte trouvés dans le dossier actuel :")
        for idx, f in enumerate(txt_locaux, 1):
            print(f"  [{idx}] {f}")
        print(f"  [M] Ouvrir le sélecteur macOS")
        print(f"  [S] Saisir le chemin manuellement")
        
        choix = input("\n👉 Choix (défaut=1) : ").strip()
        if choix.upper() == 'M':
            liste_path = get_native_file_picker("Sélectionnez le fichier liste d'élèves")
        elif choix.upper() == 'S':
            liste_path = input("Chemin du fichier liste : ").strip()
        else:
            try:
                idx_sel = int(choix) - 1 if choix else 0
                liste_path = txt_locaux[idx_sel]
            except (ValueError, IndexError):
                liste_path = txt_locaux[0]
    else:
        if sys.platform == "darwin":
            liste_path = get_native_file_picker("Sélectionnez le fichier liste d'élèves")
        if not liste_path:
            liste_path = input("👉 Entrez le chemin du fichier eleves.txt : ").strip()

    if not liste_path or not os.path.exists(liste_path):
        print(f"\n❌ Fichier de liste introuvable : '{liste_path}'")
        sys.exit(1)

    # Lecture des élèves
    with open(liste_path, "r", encoding="utf-8") as f:
        eleves = [line.strip() for line in f if line.strip()]

    print(f"✅ Liste chargée : {len(eleves)} élève(s) trouvé(s)\n")

    # -------------------------------------------------------------------------
    # 3. DOSSIER CIBLE DE DESTINATION
    # -------------------------------------------------------------------------
    print("---------------------------------------------------------------")
    print("📌 3. DOSSIER DE DESTINATION")
    print("Où souhaitez-vous enregistrer les fichiers générés ?")
    print("  [1] Créer un dossier dans le dossier courant (ex: './fichiers_eleves')")
    print("  [M] Parcourir et choisir un dossier avec le sélecteur macOS")
    
    choix_dossier = input("\n👉 Choix (défaut=1) : ").strip()
    dossier_cible = ""

    if choix_dossier.upper() == 'M':
        dossier_cible = get_native_folder_picker("Sélectionnez le dossier de destination")
    
    if not dossier_cible:
        nom_defaut = "fichiers_eleves"
        saisie = input(f"Nom du dossier à créer/utiliser (défaut='{nom_defaut}') : ").strip()
        dossier_cible = saisie if saisie else nom_defaut

    os.makedirs(dossier_cible, exist_ok=True)
    print(f"✅ Dossier cible prêt : {os.path.abspath(dossier_cible)}\n")

    # -------------------------------------------------------------------------
    # 4. SUFFIXE / PREFIXE ET FORMATAGE DE NOM
    # -------------------------------------------------------------------------
    print("---------------------------------------------------------------")
    print("📌 4. PERSONNALISATION DU NOM DE FICHIER")
    print("Exemple de nom par défaut : Marie_Curie.xlsx")
    ajout = input("👉 Souhaitez-vous ajouter du texte (ex: _CM2, _Trimestre1, _2026) ? (Laissez vide pour aucun) : ").strip()

    position = "suffixe"
    if ajout:
        print("\nPosition du texte ajouté :")
        print(f"  [1] À la fin (Suffixe)   -> Marie_Curie{ajout}.xlsx")
        print(f"  [2] Au début (Préfixe)  -> {ajout}_Marie_Curie.xlsx")
        pos_choix = input("👉 Choix (défaut=1) : ").strip()
        if pos_choix == "2":
            position = "prefixe"

    # -------------------------------------------------------------------------
    # 5. RECAPITULATIF ET CONFIRMATION
    # -------------------------------------------------------------------------
    print("\n===============================================================")
    print("📋 RÉCAPITULATIF DE L'AUTOMATISATION")
    print("===============================================================")
    print(f" 📄 Template       : {template_path}")
    print(f" 👥 Liste élèves   : {liste_path} ({len(eleves)} élèves)")
    print(f" 📁 Dossier cible  : {os.path.abspath(dossier_cible)}")
    
    print("\n🔍 Aperçu des premiers fichiers qui vont être créés :")
    exemples = eleves[:3]
    for e in exemples:
        clean_name = e.replace(' ', '_')
        if not ajout:
            nom_f = f"{clean_name}.xlsx"
        elif position == "prefixe":
            nom_f = f"{ajout}_{clean_name}.xlsx"
        else:
            nom_f = f"{clean_name}_{ajout}.xlsx"
        print(f"   • {nom_f}")
    if len(eleves) > 3:
        print(f"   ... et {len(eleves) - 3} autre(s)")

    print("===============================================================")
    confirm = input("\n🚀 Lancer la génération ? [O/n] : ").strip().lower()

    if confirm in ['', 'o', 'oui', 'y', 'yes']:
        print("\n⏳ Traitement en cours...")
        count = 0
        for eleve in eleves:
            clean_name = eleve.replace(' ', '_')
            if not ajout:
                nom_f = f"{clean_name}.xlsx"
            elif position == "prefixe":
                nom_f = f"{ajout}_{clean_name}.xlsx"
            else:
                nom_f = f"{clean_name}_{ajout}.xlsx"

            destination = os.path.join(dossier_cible, nom_f)
            shutil.copyfile(template_path, destination)
            count += 1
            print(f"  [OK {count}/{len(eleves)}] {nom_f}")

        print(f"\n🎉 SUCCÈS ! {count} fichier(s) généré(s) dans :\n   {os.path.abspath(dossier_cible)}")
    else:
        print("\n❌ Opération annulée.")

if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n\nInterrompu par l'utilisateur.")
        sys.exit(0)
