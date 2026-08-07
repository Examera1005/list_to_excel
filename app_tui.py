#!/usr/bin/env python3
"""
Générateur de Fichiers Élèves - Interface Terminal (TUI)
Compatible macOS / Linux / Windows

Fonctionnalités avancées :
- Sélecteur interactif fluide (Flèches HAUT/BAS + Touche ESPACE pour cocher/décocher)
- Traitement PAR LOTS (Multi-classes) en 1 seul clic
- Gestion robuste des erreurs (fichiers texte nommés .xlsx créés avec nano)
- Contrôle strict des saisies utilisateur (boucle d'erreur en cas d'entrée invalide)
- Organisation automatique des sous-dossiers par classe
"""

import os
import sys
import glob
import subprocess

def ensure_openpyxl():
    """Vérifie si openpyxl est installé et propose de l'installer si nécessaire."""
    try:
        import openpyxl
        return openpyxl
    except ImportError:
        print("===============================================================")
        print(" ⚠️ DÉPENDANCE MANQUANTE : openpyxl")
        print("===============================================================")
        print("La bibliothèque 'openpyxl' est nécessaire pour lire et modifier")
        print("les fichiers Excel (.xlsx).\n")
        
        choix = input("👉 Souhaitez-vous l'installer automatiquement maintenant ? [O/n] : ").strip().lower()
        if choix in ['', 'o', 'oui', 'y', 'yes']:
            print("\n⏳ Installation de 'openpyxl' via pip...")
            try:
                subprocess.check_call([sys.executable, "-m", "pip", "install", "openpyxl"])
            except Exception:
                try:
                    subprocess.check_call([sys.executable, "-m", "pip", "install", "openpyxl", "--break-system-packages"])
                except Exception as err:
                    print(f"\n❌ Échec de l'installation automatique : {err}")
                    print("💡 Veuillez installer la bibliothèque manuellement avec :")
                    print("   pip3 install openpyxl")
                    sys.exit(1)
                    
            print("✅ 'openpyxl' a été installé avec succès !\n")
            import openpyxl
            return openpyxl
        else:
            print("\n❌ Impossible de continuer sans 'openpyxl'.")
            sys.exit(1)

openpyxl = ensure_openpyxl()

def clear_screen():
    os.system('cls' if os.name == 'nt' else 'clear')

def get_valid_input(prompt_text, valid_options, default=None):
    """Demande une saisie et boucle tant que la réponse n'est pas valide."""
    valid_upper = [str(o).upper() for o in valid_options]
    while True:
        saisie = input(prompt_text).strip()
        if not saisie and default is not None:
            return str(default).upper()
        if saisie.upper() in valid_upper:
            return saisie.upper()
        print(f"❌ Saisie invalide ('{saisie}'). Options autorisées : {', '.join(map(str, valid_options))}" + 
              (f" [Appuyez sur Entrée pour '{default}']" if default is not None else ""))

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

def read_single_key():
    """Lit une seule touche du clavier (Flèches, Espace, Entrée, etc.)."""
    if os.name == 'nt':
        import msvcrt
        ch = msvcrt.getch()
        if ch in (b'\x00', b'\xe0'):
            code = msvcrt.getch()
            if code == b'H': return 'UP'
            if code == b'P': return 'DOWN'
        elif ch in (b'\r', b'\n'): return 'ENTER'
        elif ch == b' ': return 'SPACE'
        return ch.decode('utf-8', errors='ignore').upper()
    else:
        import tty, termios
        fd = sys.stdin.fileno()
        old_settings = termios.tcgetattr(fd)
        try:
            tty.setraw(fd)
            ch = sys.stdin.read(1)
            if ch == '\x1b':
                ch2 = sys.stdin.read(1)
                ch3 = sys.stdin.read(1)
                if ch3 == 'A': return 'UP'
                if ch3 == 'B': return 'DOWN'
            elif ch in ('\r', '\n'): return 'ENTER'
            elif ch == ' ': return 'SPACE'
            return ch.upper()
        finally:
            termios.tcsetattr(fd, termios.TCSANOW, old_settings)

def interactive_checkbox_selector(items, title="SÉLECTION DES CLASSES"):
    """
    Sélecteur interactif avec Flèches HAUT/BAS, ESPACE pour cocher/décocher, ENTRÉE pour valider.
    """
    if not items:
        return []

    cursor = 0
    checked = [True] * len(items)

    while True:
        clear_screen()
        print("===============================================================")
        print(f" 👥 {title}")
        print("===============================================================")
        print(" 💡 Navigation :  Flèche ⬆️  / ⬇️  |  ESPACE = Cocher / Décocher")
        print("                [A] = Tout cocher | [N] = Tout décocher | ENTRÉE = Valider\n")
        
        for idx, item in enumerate(items):
            is_cur = "👉 " if idx == cursor else "   "
            is_chk = "[✓]" if checked[idx] else "[ ]"
            print(f"{is_cur}{is_chk} {item}")
            
        nb_checked = sum(checked)
        print("---------------------------------------------------------------")
        print(f"📊 {nb_checked} classe(s) sélectionnée(s) sur {len(items)}")
        
        key = read_single_key()
        if key == 'UP':
            cursor = (cursor - 1) % len(items)
        elif key == 'DOWN':
            cursor = (cursor + 1) % len(items)
        elif key == 'SPACE':
            checked[cursor] = not checked[cursor]
        elif key == 'A':
            checked = [True] * len(items)
        elif key == 'N':
            checked = [False] * len(items)
        elif key == 'ENTER':
            selected = [items[i] for i in range(len(items)) if checked[i]]
            if not selected:
                print("\n❌ Veuillez cocher au moins un fichier avec la touche ESPACE !")
                import time; time.sleep(1.2)
                continue
            return selected

def parse_student_list(liste_path):
    """
    Extrait les élèves depuis un fichier Excel (.xlsx) ou Texte (.txt).
    Gère automatiquement le fallback vers le texte si un fichier créé avec nano a une extension .xlsx.
    """
    eleves = []
    
    # 1. Tenter la lecture comme fichier Excel (.xlsx)
    if liste_path.endswith(".xlsx"):
        try:
            wb = openpyxl.load_workbook(liste_path, data_only=True)
            ws = wb.active
            
            header_row = None
            for r in range(1, 30):
                vals = [str(ws.cell(row=r, column=c).value or '').strip() for c in range(1, 15)]
                if 'Nom' in vals and ('Prénom' in vals or 'Prenom' in vals):
                    header_row = r
                    break
                    
            if header_row:
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
                return eleves
        except Exception:
            # Si le fichier n'est pas un binaire Excel valide (ex: fichier texte nommé .xlsx créé avec nano)
            pass

    # 2. Fallback / Lecture comme fichier Texte brut (.txt ou fichier texte .xlsx)
    try:
        with open(liste_path, "r", encoding="utf-8", errors="ignore") as f:
            for line in f:
                line = line.strip()
                if line and not line.startswith("#"):
                    parts = line.split()
                    if len(parts) >= 2:
                        prenom = parts[0]
                        nom = " ".join(parts[1:])
                    else:
                        prenom = line
                        nom = ""
                    eleves.append({"nom": nom, "prenom": prenom, "classe": ""})
    except Exception as err:
        print(f"⚠️ Impossible de lire le fichier '{liste_path}' : {err}")
        
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

    # Filtrer les vrais templates Excel
    vrais_templates = []
    for f in templates_locaux:
        try:
            wb_t = openpyxl.load_workbook(f, read_only=True)
            wb_t.close()
            vrais_templates.append(f)
        except Exception:
            pass

    if vrais_templates:
        print("Fichiers Excel modèles trouvés dans le dossier actuel :")
        valid_opts = []
        def_idx = 1
        for idx, f in enumerate(vrais_templates, 1):
            print(f"  [{idx}] {f}")
            valid_opts.append(str(idx))
            if "copie" in f.lower() or "template" in f.lower():
                def_idx = idx

        if sys.platform == "darwin":
            print("  [M] Ouvrir le sélecteur macOS")
            valid_opts.append("M")
        print("  [S] Saisir le chemin manuellement")
        valid_opts.append("S")
        
        choix = get_valid_input(f"\n👉 Choix (défaut={def_idx}) : ", valid_opts, default=def_idx)
        
        if choix == 'M':
            template_path = get_native_file_picker("Sélectionnez votre fichier template Excel")
        elif choix == 'S':
            while True:
                template_path = input("Chemin du fichier template : ").strip()
                if template_path and os.path.exists(template_path):
                    break
                print(f"❌ Fichier introuvable : '{template_path}'. Réessayez.")
        else:
            template_path = vrais_templates[int(choix) - 1]
    else:
        print("Aucun fichier modèle .xlsx valide trouvé.")
        if sys.platform == "darwin":
            template_path = get_native_file_picker("Sélectionnez votre fichier template Excel")
        while not template_path or not os.path.exists(template_path):
            template_path = input("👉 Entrez le chemin du fichier template.xlsx : ").strip()
            if not os.path.exists(template_path):
                print(f"❌ Fichier introuvable : '{template_path}'")

    print(f"✅ Template sélectionné : {template_path}\n")

    # -------------------------------------------------------------------------
    # 2. SÉLECTION DES LISTES DE CLASSES (.xlsx ou .txt)
    # -------------------------------------------------------------------------
    print("---------------------------------------------------------------")
    print("📌 2. SÉLECTION DE LA OU DES LISTES D'ÉLÈVES")
    fichiers_liste = [f for f in glob.glob("*.xlsx") + glob.glob("*.txt") if os.path.abspath(f) != os.path.abspath(template_path)]
    listes_selectionnees = []

    if fichiers_liste:
        print("  [1] Choisir UN SEUL fichier de liste")
        print(f"  [2] TRAITEMENT PAR LOTS (Multi-classes - Sélecteur Flèches + Espace)")
        if sys.platform == "darwin":
            print("  [M] Sélecteur macOS")
        print("  [S] Saisir le chemin manuellement")
        
        opts = ["1", "2", "S"]
        if sys.platform == "darwin":
            opts.append("M")

        mode_l = get_valid_input("\n👉 Mode de sélection (défaut=1) : ", opts, default="1")
        
        if mode_l == "2":
            listes_selectionnees = interactive_checkbox_selector(fichiers_liste, title="SÉLECTION MULTIPLE DES CLASSES (Flèches + Espace)")
        elif mode_l == "M":
            f_mac = get_native_file_picker("Sélectionnez le fichier liste d'élèves")
            if f_mac:
                listes_selectionnees = [f_mac]
        elif mode_l == "S":
            while True:
                f_s = input("Chemin du fichier liste : ").strip()
                if f_s and os.path.exists(f_s):
                    listes_selectionnees = [f_s]
                    break
                print(f"❌ Fichier introuvable : '{f_s}'. Réessayez.")
        else:
            print("\nFichiers trouvés :")
            v_opts = []
            def_l = 1
            for idx, f in enumerate(fichiers_liste, 1):
                print(f"  [{idx}] {f}")
                v_opts.append(str(idx))
                if "liste" in f.lower() or f == "eleves.txt":
                    def_l = idx
            ch_l = get_valid_input(f"\n👉 Choix du fichier (défaut={def_l}) : ", v_opts, default=def_l)
            listes_selectionnees = [fichiers_liste[int(ch_l) - 1]]
    else:
        print("Aucune liste d'élèves trouvée dans le dossier.")
        if sys.platform == "darwin":
            f_mac = get_native_file_picker("Sélectionnez le fichier liste d'élèves")
            if f_mac:
                listes_selectionnees = [f_mac]
        while not listes_selectionnees:
            f_s = input("👉 Entrez le chemin du fichier liste : ").strip()
            if os.path.exists(f_s):
                listes_selectionnees = [f_s]

    print(f"\n✅ {len(listes_selectionnees)} fichier(s) de classe retenu(s) pour le traitement :\n")
    for f in listes_selectionnees:
        print(f"   • {f}")

    # -------------------------------------------------------------------------
    # 3. ORGANISATION DES DOSSIERS DE DESTINATION
    # -------------------------------------------------------------------------
    print("\n---------------------------------------------------------------")
    print("📌 3. DOSSIER PARENT DE DESTINATION")
    print("Où souhaitez-vous enregistrer les fichiers générés ?")
    print("  [1] Dossier './fichiers_eleves' (Créera automatiquement un sous-dossier par classe si multi-classes)")
    print("  [2] Dossier personnalisé")
    if sys.platform == "darwin":
        print("  [M] Parcourir avec le sélecteur macOS")

    opts_d = ["1", "2"]
    if sys.platform == "darwin":
        opts_d.append("M")

    choix_d = get_valid_input("\n👉 Choix (défaut=1) : ", opts_d, default="1")
    dossier_parent = "fichiers_eleves"

    if choix_d == "2":
        s = input("Entrez le nom/chemin du dossier parent : ").strip()
        if s:
            dossier_parent = s
    elif choix_d == "M":
        mac_d = get_native_folder_picker("Sélectionnez le dossier parent de destination")
        if mac_d:
            dossier_parent = mac_d

    os.makedirs(dossier_parent, exist_ok=True)
    print(f"✅ Dossier parent prêt : {os.path.abspath(dossier_parent)}\n")

    # -------------------------------------------------------------------------
    # 4. OPTIONS DE FORMATAGE DES NOMS ET CELLULE C3
    # -------------------------------------------------------------------------
    print("---------------------------------------------------------------")
    print("📌 4. FORMATAGE DU NOM DE FICHIER ET CELLULE C3")
    print("Format du nom de fichier :")
    print("  [1] Avec tiret bas '_'  -> 1M1_Nom_Prenom.xlsx (Recommandé)")
    print("  [2] Avec espaces ' '   -> 1M1 Nom Prenom.xlsx")
    fmt_choix = get_valid_input("👉 Choix (défaut=1) : ", ["1", "2"], default="1")
    separateur = " " if fmt_choix == "2" else "_"

    print("\nFormat de la cellule C3 :")
    print("  [1] Prénom Nom (ex: Alice Dupont)")
    print("  [2] Nom Prénom (ex: Dupont Alice)")
    c3_choix = get_valid_input("👉 Choix (défaut=1) : ", ["1", "2"], default="1")
    c3_format = "nom_prenom" if c3_choix == "2" else "prenom_nom"

    # -------------------------------------------------------------------------
    # 5. RÉCAPITULATIF ET CONFIRMATION
    # -------------------------------------------------------------------------
    print("\n===============================================================")
    print("📋 RÉCAPITULATIF DE L'AUTOMATISATION")
    print("===============================================================")
    print(f" 📄 Template       : {template_path}")
    print(f" 👥 Listes classes : {len(listes_selectionnees)} fichier(s) sélectionné(s)")
    for f in listes_selectionnees:
        print(f"     • {f}")
    print(f" 📁 Dossier Parent : {os.path.abspath(dossier_parent)}")
    print(f" ✏️ Cellule C3      : {'Nom Prénom' if c3_format == 'nom_prenom' else 'Prénom Nom'}")
    print("===============================================================")
    
    confirm = get_valid_input("\n🚀 Lancer la génération pour TOUTES les classes ? [O/N] : ", ["O", "N", "OUI", "NON", "Y", "YES"], default="O")

    if confirm in ['O', 'OUI', 'Y', 'YES']:
        print("\n⏳ Traitement en cours...")
        total_eleves = 0
        total_fichiers = 0

        for liste_file in listes_selectionnees:
            eleves = parse_student_list(liste_file)
            if not eleves:
                print(f"\n⚠️ Aucun élève trouvé dans '{liste_file}'. Fichier ignoré.")
                continue

            nom_classe = eleves[0]["classe"] if eleves[0]["classe"] else os.path.splitext(os.path.basename(liste_file))[0]
            
            if len(listes_selectionnees) > 1:
                target_dir = os.path.join(dossier_parent, nom_classe)
            else:
                target_dir = dossier_parent

            os.makedirs(target_dir, exist_ok=True)
            print(f"\n📂 Classe '{nom_classe}' ({len(eleves)} élèves) -> Dossier: {os.path.abspath(target_dir)}")

            count = 0
            for e in eleves:
                nom = e["nom"]
                prenom = e["prenom"]
                classe = e["classe"] if e["classe"] else nom_classe
                parts = [p for p in [classe, nom, prenom] if p]
                
                nom_f = ("_".join(parts).replace(" ", "_") if separateur == "_" else " ".join(parts)) + ".xlsx"
                c3_val = f"{nom} {prenom}" if c3_format == "nom_prenom" else f"{prenom} {nom}"

                destination = os.path.join(target_dir, nom_f)
                
                wb = openpyxl.load_workbook(template_path)
                ws = wb.active
                ws['C3'] = c3_val.strip()
                wb.save(destination)
                
                count += 1
                total_eleves += 1
                total_fichiers += 1
                print(f"  [OK {count}/{len(eleves)}] {nom_f}")

        print(f"\n🎉 SUCCÈS TOTAL ! {total_fichiers} fichier(s) généré(s) pour {len(listes_selectionnees)} classe(s) !")
        print(f"📁 Localisation : {os.path.abspath(dossier_parent)}")
    else:
        print("\n❌ Opération annulée.")

if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n\nInterrompu par l'utilisateur.")
        sys.exit(0)
