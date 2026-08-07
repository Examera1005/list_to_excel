#!/usr/bin/env python3
"""
Générateur de Fichiers Élèves - Interface Terminal (TUI)
Compatible macOS / Linux / Windows
"""

import os
import sys
import glob
import subprocess
import openpyxl

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

def parse_student_list(liste_path):
    """Extrait les élèves depuis un fichier Excel (.xlsx) ou Texte (.txt)."""
    eleves = []
    
    if liste_path.endswith(".xlsx"):
        wb = openpyxl.load_workbook(liste_path, data_only=True)
        ws = wb.active
        
        header_row = None
        for r in range(1, 30):
            vals = [str(ws.cell(row=r, column=c).value or '').strip() for c in range(1, 15)]
            if 'Nom' in vals and ('Prénom' in vals or 'Prenom' in vals):
                header_row = r
                break
                
        if not header_row:
            print("❌ En-têtes 'Nom' et 'Prénom' introuvables dans le fichier Excel.")
            return []
            
        nom_c = next(c for c in range(1, 15) if str(ws.cell(row=header_row, column=c).value).strip().lower() == 'nom')
        prenom_c = next(c for c in range(1, 15) if str(ws.cell(row=header_row, column=c).value).strip().lower() in ['prénom', 'prenom'])
        group_c = next((c for c in range(1, 15) if str(ws.cell(row=header_row, column=c).value).strip().lower() in ['groupe', 'classe']), None)
        
        for r in range(header_row + 1, ws.max_row + 1):
            nom = ws.cell(row=r, column=nom_c).value
            prenom = ws.cell(row=r, column=prenom_c).value
            grp = ws.cell(row=r, column=group_c).value if group_c else ''
            
            if nom and prenom:
                eleves.append({
                    "nom": str(nom).strip(),
                    "prenom": str(prenom).strip(),
                    "classe": str(grp).strip() if grp else ""
                })
    else:
        with open(liste_path, "r", encoding="utf-8") as f:
            for line in f:
                line = line.strip()
                if line:
                    parts = line.split()
                    if len(parts) >= 2:
                        prenom = parts[0]
                        nom = " ".join(parts[1:])
                    else:
                        prenom = line
                        nom = ""
                    eleves.append({"nom": nom, "prenom": prenom, "classe": ""})
                    
    return eleves

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
        print(f"  [M] Ouvrir le sélecteur macOS")
        print(f"  [S] Saisir le chemin manuellement")
        
        # Prédélection du template
        def_idx = 1
        for idx, f in enumerate(templates_locaux, 1):
            if "copie" in f.lower() or "template" in f.lower():
                def_idx = idx
                break

        choix = input(f"\n👉 Choix (défaut={def_idx}) : ").strip()
        if choix.upper() == 'M':
            template_path = get_native_file_picker("Sélectionnez votre fichier template Excel")
        elif choix.upper() == 'S':
            template_path = input("Chemin du fichier template : ").strip()
        else:
            try:
                idx_sel = int(choix) - 1 if choix else (def_idx - 1)
                template_path = templates_locaux[idx_sel]
            except (ValueError, IndexError):
                template_path = templates_locaux[def_idx - 1]
    else:
        print("Aucun fichier .xlsx dans le dossier courant.")
        if sys.platform == "darwin":
            template_path = get_native_file_picker("Sélectionnez votre fichier template Excel")
        if not template_path:
            template_path = input("👉 Entrez le chemin du fichier template.xlsx : ").strip()

    if not template_path or not os.path.exists(template_path):
        print(f"\n❌ Fichier template introuvable : '{template_path}'")
        sys.exit(1)

    print(f"✅ Template sélectionné : {template_path}\n")

    # -------------------------------------------------------------------------
    # 2. SÉLECTION DE LA LISTE DES ÉLÈVES (.xlsx ou .txt)
    # -------------------------------------------------------------------------
    print("---------------------------------------------------------------")
    print("📌 2. LISTE DES ÉLÈVES (.xlsx ou .txt)")
    fichiers_liste = glob.glob("*.xlsx") + glob.glob("*.txt")
    liste_path = ""

    if fichiers_liste:
        print("Fichiers trouvés dans le dossier actuel :")
        for idx, f in enumerate(fichiers_liste, 1):
            print(f"  [{idx}] {f}")
        print(f"  [M] Ouvrir le sélecteur macOS")
        print(f"  [S] Saisir le chemin manuellement")
        
        # Prédirection de la liste
        def_idx_l = 1
        for idx, f in enumerate(fichiers_liste, 1):
            if "liste" in f.lower() or f == "eleves.txt":
                def_idx_l = idx
                break

        choix = input(f"\n👉 Choix (défaut={def_idx_l}) : ").strip()
        if choix.upper() == 'M':
            liste_path = get_native_file_picker("Sélectionnez le fichier liste d'élèves")
        elif choix.upper() == 'S':
            liste_path = input("Chemin du fichier liste : ").strip()
        else:
            try:
                idx_sel = int(choix) - 1 if choix else (def_idx_l - 1)
                liste_path = fichiers_liste[idx_sel]
            except (ValueError, IndexError):
                liste_path = fichiers_liste[def_idx_l - 1]
    else:
        if sys.platform == "darwin":
            liste_path = get_native_file_picker("Sélectionnez le fichier liste d'élèves")
        if not liste_path:
            liste_path = input("👉 Entrez le chemin du fichier liste : ").strip()

    if not liste_path or not os.path.exists(liste_path):
        print(f"\n❌ Fichier de liste introuvable : '{liste_path}'")
        sys.exit(1)

    eleves = parse_student_list(liste_path)
    if not eleves:
        print(f"\n❌ Impossible d'extraire la liste d'élèves depuis '{liste_path}'")
        sys.exit(1)

    print(f"✅ Liste chargée : {len(eleves)} élève(s) trouvé(s)\n")

    # -------------------------------------------------------------------------
    # 3. DOSSIER CIBLE DE DESTINATION
    # -------------------------------------------------------------------------
    print("---------------------------------------------------------------")
    print("📌 3. DOSSIER DE DESTINATION")
    print("  [1] Créer un dossier dans le dossier courant ('./fichiers_eleves')")
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
    # 4. OPTIONS DE FORMATAGE DES NOMS ET CELLULE C3
    # -------------------------------------------------------------------------
    print("---------------------------------------------------------------")
    print("📌 4. FORMATAGE DU NOM DE FICHIER ET CELLULE C3")
    print("Format du nom de fichier :")
    print("  [1] Avec tiret bas '_'  -> 1M4_Abadin_Robert.xlsx (Recommandé)")
    print("  [2] Avec espaces ' '   -> 1M4 Abadin Robert.xlsx")
    fmt_choix = input("👉 Choix (défaut=1) : ").strip()
    separateur = " " if fmt_choix == "2" else "_"

    print("\nFormat de la cellule C3 :")
    print("  [1] Prénom Nom (ex: Robert Abadin)")
    print("  [2] Nom Prénom (ex: Abadin Robert)")
    c3_choix = input("👉 Choix (défaut=1) : ").strip()
    c3_format = "nom_prenom" if c3_choix == "2" else "prenom_nom"

    # -------------------------------------------------------------------------
    # 5. RÉCAPITULATIF ET CONFIRMATION
    # -------------------------------------------------------------------------
    print("\n===============================================================")
    print("📋 RÉCAPITULATIF DE L'AUTOMATISATION")
    print("===============================================================")
    print(f" 📄 Template       : {template_path}")
    print(f" 👥 Liste élèves   : {liste_path} ({len(eleves)} élèves)")
    print(f" 📁 Dossier cible  : {os.path.abspath(dossier_cible)}")
    print(f" ✏️ Cellule C3      : {'Nom Prénom' if c3_format == 'nom_prenom' else 'Prénom Nom'}")
    
    print("\n🔍 Aperçu des premiers fichiers qui vont être créés :")
    exemples = eleves[:3]
    for e in exemples:
        nom = e["nom"]
        prenom = e["prenom"]
        classe = e["classe"]
        parts = [p for p in [classe, nom, prenom] if p]
        if separateur == "_":
            nom_f = "_".join(parts).replace(" ", "_") + ".xlsx"
        else:
            nom_f = " ".join(parts) + ".xlsx"
        c3_val = f"{nom} {prenom}" if c3_format == "nom_prenom" else f"{prenom} {nom}"
        print(f"   • Nom de fichier : {nom_f} | C3 = '{c3_val.strip()}'")
    if len(eleves) > 3:
        print(f"   ... et {len(eleves) - 3} autre(s)")

    print("===============================================================")
    confirm = input("\n🚀 Lancer la génération ? [O/n] : ").strip().lower()

    if confirm in ['', 'o', 'oui', 'y', 'yes']:
        print("\n⏳ Traitement en cours...")
        count = 0
        for e in eleves:
            nom = e["nom"]
            prenom = e["prenom"]
            classe = e["classe"]
            parts = [p for p in [classe, nom, prenom] if p]
            
            if separateur == "_":
                nom_f = "_".join(parts).replace(" ", "_") + ".xlsx"
            else:
                nom_f = " ".join(parts) + ".xlsx"
                
            c3_val = f"{nom} {prenom}" if c3_format == "nom_prenom" else f"{prenom} {nom}"

            destination = os.path.join(dossier_cible, nom_f)
            
            wb = openpyxl.load_workbook(template_path)
            ws = wb.active
            ws['C3'] = c3_val.strip()
            wb.save(destination)
            
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
