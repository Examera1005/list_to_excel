#!/usr/bin/env python3
"""
Script de duplication automatique de templates Excel par élève.
- Extrait la liste des élèves (fichier .xlsx ou .txt)
- Remplit la cellule C3 du template Excel avec "Prénom Nom"
- Génère les fichiers au format "[Classe]_[Nom]_[Prénom].xlsx"
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

def find_files():
    """Détecte automatiquement le fichier template et le fichier de liste."""
    all_xlsx = glob.glob("*.xlsx")
    
    template_file = None
    liste_file = None
    
    for f in all_xlsx:
        f_lower = f.lower()
        if "liste" in f_lower:
            liste_file = f
        elif "copie" in f_lower or "template" in f_lower:
            template_file = f
            
    if not template_file:
        for f in all_xlsx:
            if f != liste_file:
                template_file = f
                break
                
    if not liste_file and os.path.exists("eleves.txt"):
        liste_file = "eleves.txt"
        
    return template_file, liste_file

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
            print("❌ Impossible de trouver les en-têtes 'Nom' et 'Prénom' dans la feuille.")
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
    template_file, liste_file = find_files()
    
    if len(sys.argv) > 1:
        template_file = sys.argv[1]
    if len(sys.argv) > 2:
        liste_file = sys.argv[2]
        
    print("===============================================================")
    print(" 🎓 DÉMARRAGE DU SCRIPT DE DUPLICATION EXCEL")
    print("===============================================================")
    print(f" 📄 Template       : {template_file}")
    print(f" 👥 Liste d'élèves : {liste_file}")
    print("---------------------------------------------------------------")
    
    if not template_file or not os.path.exists(template_file):
        print("❌ Erreur : Fichier template introuvable.")
        sys.exit(1)
        
    if not liste_file or not os.path.exists(liste_file):
        print("❌ Erreur : Fichier liste d'élèves introuvable.")
        sys.exit(1)
        
    eleves = parse_student_list(liste_file)
    print(f"✅ {len(eleves)} élève(s) extrait(s) de la liste.\n")
    
    output_dir = "fichiers_eleves"
    os.makedirs(output_dir, exist_ok=True)
    
    count = 0
    for e in eleves:
        nom = e["nom"]
        prenom = e["prenom"]
        classe = e["classe"]
        
        parts = [p for p in [classe, nom, prenom] if p]
        nom_fichier = "_".join(parts).replace(" ", "_") + ".xlsx"
        
        wb = openpyxl.load_workbook(template_file)
        ws = wb.active
        
        ws['C3'] = f"{prenom} {nom}".strip()
        
        dest_path = os.path.join(output_dir, nom_fichier)
        wb.save(dest_path)
        count += 1
        print(f"  ✅ [{count}/{len(eleves)}] Généré : {nom_fichier} (Cellule C3 = '{ws['C3'].value}')")
        
    print("\n===============================================================")
    print(f"🎉 SUCCÈS ! {count} fichier(s) généré(s) dans le dossier :")
    print(f"   📁 {os.path.abspath(output_dir)}")
    print("===============================================================")

if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n\nInterrompu par l'utilisateur.")
        sys.exit(0)
