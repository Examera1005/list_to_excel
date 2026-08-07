#!/usr/bin/env python3
"""
🎓 Générateur de Fichiers Élèves - Interface Graphique (GUI)
Compatible macOS et Linux (PySide6 / Qt6)
Interface moderne, légère (~30Mo RAM) et réactive.
"""

import os
import sys
import glob
import subprocess

def ensure_pyside6():
    """Vérifie l'installation de PySide6 et propose l'installation automatique si besoin."""
    try:
        from PySide6 import QtWidgets, QtCore, QtGui
        return QtWidgets, QtCore, QtGui
    except ImportError:
        print("===============================================================")
        print(" ⚠️ DÉPENDANCE MANQUANTE : PySide6")
        print("===============================================================")
        print("L'interface graphique nécessite la bibliothèque 'PySide6'.\n")
        
        choix = input("👉 Souhaitez-vous l'installer automatiquement maintenant ? [O/n] : ").strip().lower()
        if choix in ['', 'o', 'oui', 'y', 'yes']:
            print("\n⏳ Installation de 'PySide6' via pip...")
            try:
                subprocess.check_call([sys.executable, "-m", "pip", "install", "PySide6"])
            except Exception:
                try:
                    subprocess.check_call([sys.executable, "-m", "pip", "install", "PySide6", "--break-system-packages"])
                except Exception as err:
                    print(f"\n❌ Échec de l'installation automatique : {err}")
                    print("💡 Veuillez l'installer manuellement avec :")
                    print("   pip3 install PySide6")
                    sys.exit(1)
                    
            print("✅ 'PySide6' a été installé avec succès !\n")
            from PySide6 import QtWidgets, QtCore, QtGui
            return QtWidgets, QtCore, QtGui
        else:
            print("\n❌ Impossible d'ouvrir la GUI sans PySide6.")
            sys.exit(1)

QtWidgets, QtCore, QtGui = ensure_pyside6()
import openpyxl

def parse_student_list(liste_path):
    """Extrait les élèves depuis un fichier Excel (.xlsx) ou Texte (.txt)."""
    eleves = []
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
                        eleves.append({"nom": str(nom).strip(), "prenom": str(prenom).strip(), "classe": str(grp).strip() if grp else ""})
                return eleves
        except Exception:
            pass

    try:
        with open(liste_path, "r", encoding="utf-8", errors="ignore") as f:
            for line in f:
                line = line.strip()
                if line and not line.startswith('#'):
                    parts = line.split()
                    if len(parts) >= 2:
                        prenom = parts[0]
                        nom = " ".join(parts[1:])
                    else:
                        prenom = line
                        nom = ""
                    eleves.append({"nom": nom, "prenom": prenom, "classe": ""})
    except Exception:
        pass
    return eleves

class WorkerThread(QtCore.QThread):
    """Thread d'arrière-plan pour la génération sans bloquer l'interface."""
    progress_signal = QtCore.Signal(int, int, str)  # (current, total, message)
    finished_signal = QtCore.Signal(int, int, str)  # (total_files, total_classes, parent_dir)
    error_signal = QtCore.Signal(str)

    def __init__(self, template_path, listes_files, parent_dir, separateur, c3_format):
        super().__init__()
        self.template_path = template_path
        self.listes_files = listes_files
        self.parent_dir = parent_dir
        self.separateur = separateur
        self.c3_format = c3_format

    def run(self):
        try:
            total_fichiers = 0
            # Compter d'abord le nombre total d'élèves
            all_data = []
            for l_file in self.listes_files:
                eleves = parse_student_list(l_file)
                if eleves:
                    all_data.append((l_file, eleves))

            total_items = sum(len(e) for _, e in all_data)
            current_item = 0

            for l_file, eleves in all_data:
                nom_classe = eleves[0]["classe"] if eleves[0]["classe"] else os.path.splitext(os.path.basename(l_file))[0]
                target_dir = os.path.join(self.parent_dir, nom_classe) if len(all_data) > 1 else self.parent_dir
                os.makedirs(target_dir, exist_ok=True)

                for e in eleves:
                    nom = e["nom"]
                    prenom = e["prenom"]
                    classe = e["classe"] if e["classe"] else nom_classe
                    parts = [p for p in [classe, nom, prenom] if p]
                    
                    nom_f = ("_".join(parts).replace(" ", "_") if self.separateur == "_" else " ".join(parts)) + ".xlsx"
                    c3_val = f"{nom} {prenom}" if self.c3_format == "nom_prenom" else f"{prenom} {nom}"

                    destination = os.path.join(target_dir, nom_f)
                    
                    wb = openpyxl.load_workbook(self.template_path)
                    ws = wb.active
                    ws['C3'] = c3_val.strip()
                    wb.save(destination)

                    current_item += 1
                    total_fichiers += 1
                    msg = f"✅ [{current_item}/{total_items}] [{nom_classe}] Généré : {nom_f} (C3 = '{c3_val.strip()}')"
                    self.progress_signal.emit(current_item, total_items, msg)

            self.finished_signal.emit(total_fichiers, len(all_data), os.path.abspath(self.parent_dir))
        except Exception as err:
            self.error_signal.emit(str(err))

class MainWindow(QtWidgets.QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("🎓 Générateur de Documents Élèves")
        self.setMinimumSize(850, 700)
        self.resize(900, 750)

        self.template_path = ""
        self.listes_files = []
        self.parent_dir = os.path.abspath("fichiers_eleves")

        self.init_ui()
        self.auto_detect_files()

    def init_ui(self):
        # Feuille de style CSS élégante et sombre (Dark Theme) avec effets hover
        self.setStyleSheet("""
            QMainWindow {
                background-color: #0F172A;
            }
            QWidget {
                color: #F8FAFC;
                font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif;
                font-size: 13px;
            }
            QGroupBox {
                background-color: #1E293B;
                border: 1px solid #334155;
                border-radius: 12px;
                margin-top: 10px;
                padding: 15px;
                font-weight: bold;
                font-size: 14px;
                color: #38BDF8;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 15px;
                padding: 0 5px;
            }
            QPushButton {
                background-color: #3B82F6;
                color: #FFFFFF;
                border: none;
                border-radius: 8px;
                padding: 9px 16px;
                font-weight: 600;
            }
            QPushButton:hover {
                background-color: #2563EB;
            }
            QPushButton:pressed {
                background-color: #1D4ED8;
            }
            QPushButton#btnLaunch {
                background-color: #10B981;
                font-size: 15px;
                padding: 12px 24px;
                font-weight: bold;
            }
            QPushButton#btnLaunch:hover {
                background-color: #059669;
            }
            QPushButton#btnLaunch:pressed {
                background-color: #047857;
            }
            QLineEdit {
                background-color: #0F172A;
                border: 1px solid #334155;
                border-radius: 8px;
                padding: 8px 12px;
                color: #F1F5F9;
            }
            QListWidget {
                background-color: #0F172A;
                border: 1px solid #334155;
                border-radius: 8px;
                padding: 6px;
                color: #F8FAFC;
            }
            QListWidget::item {
                padding: 6px;
                border-radius: 4px;
            }
            QListWidget::item:hover {
                background-color: #334155;
            }
            QListWidget::item:selected {
                background-color: #2563EB;
                color: #FFFFFF;
            }
            QRadioButton, QCheckBox {
                color: #CBD5E1;
                spacing: 8px;
            }
            QRadioButton::indicator, QCheckBox::indicator {
                width: 18px;
                height: 18px;
            }
            QProgressBar {
                background-color: #0F172A;
                border: 1px solid #334155;
                border-radius: 8px;
                text-align: center;
                color: #F8FAFC;
                font-weight: bold;
            }
            QProgressBar::chunk {
                background-color: #10B981;
                border-radius: 7px;
            }
            QTextEdit {
                background-color: #0F172A;
                border: 1px solid #334155;
                border-radius: 8px;
                color: #34D399;
                font-family: monospace;
                font-size: 12px;
                padding: 8px;
            }
        """)

        central_widget = QtWidgets.QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QtWidgets.QVBoxLayout(central_widget)
        main_layout.setContentsMargins(20, 20, 20, 20)
        main_layout.setSpacing(15)

        # En-tête
        header_layout = QtWidgets.QHBoxLayout()
        icon_label = QtWidgets.QLabel("🎓")
        icon_label.setFont(QtGui.QFont("Segoe UI Emoji", 26))
        header_text_layout = QtWidgets.QVBoxLayout()
        title_label = QtWidgets.QLabel("Générateur de Documents Élèves")
        title_label.setStyleSheet("font-size: 20px; font-weight: bold; color: #F8FAFC;")
        subtitle_label = QtWidgets.QLabel("Duplication automatique de templates Excel par classe")
        subtitle_label.setStyleSheet("font-size: 12px; color: #94A3B8;")
        header_text_layout.addWidget(title_label)
        header_text_layout.addWidget(subtitle_label)
        header_layout.addWidget(icon_label)
        header_layout.addLayout(header_text_layout)
        header_layout.addStretch()
        main_layout.addLayout(header_layout)

        # --- GROUPE 1: Template Excel ---
        group_template = QtWidgets.QGroupBox("1. Fichier Modèle Template Excel (.xlsx)")
        gt_layout = QtWidgets.QHBoxLayout(group_template)
        self.txt_template = QtWidgets.QLineEdit()
        self.txt_template.setReadOnly(True)
        self.txt_template.setPlaceholderText("Aucun template sélectionné...")
        btn_browse_tpl = QtWidgets.QPushButton("📁 Parcourir...")
        btn_browse_tpl.setCursor(QtCore.Qt.PointingHandCursor)
        btn_browse_tpl.clicked.connect(self.browse_template)
        gt_layout.addWidget(self.txt_template)
        gt_layout.addWidget(btn_browse_tpl)
        main_layout.addWidget(group_template)

        # --- GROUPE 2: Listes des élèves ---
        group_listes = QtWidgets.QGroupBox("2. Listes de Classes (.xlsx ou .txt)")
        gl_layout = QtWidgets.QVBoxLayout(group_listes)
        
        gl_top_layout = QtWidgets.QHBoxLayout()
        self.lbl_listes_count = QtWidgets.QLabel("0 classe(s) sélectionnée(s)")
        self.lbl_listes_count.setStyleSheet("color: #94A3B8; font-weight: normal;")
        btn_browse_listes = QtWidgets.QPushButton("👥 Parcourir les listes (Sélection multiple Finder/Explorateur)...")
        btn_browse_listes.setCursor(QtCore.Qt.PointingHandCursor)
        btn_browse_listes.clicked.connect(self.browse_listes)
        btn_clear_listes = QtWidgets.QPushButton("🗑️ Vider")
        btn_clear_listes.setStyleSheet("background-color: #EF4444;")
        btn_clear_listes.setCursor(QtCore.Qt.PointingHandCursor)
        btn_clear_listes.clicked.connect(self.clear_listes)
        
        gl_top_layout.addWidget(self.lbl_listes_count)
        gl_top_layout.addStretch()
        gl_top_layout.addWidget(btn_browse_listes)
        gl_top_layout.addWidget(btn_clear_listes)
        
        self.lst_listes = QtWidgets.QListWidget()
        self.lst_listes.setMaximumHeight(110)
        gl_layout.addLayout(gl_top_layout)
        gl_layout.addWidget(self.lst_listes)
        main_layout.addWidget(group_listes)

        # --- GROUPE 3: Dossier de destination ---
        group_dest = QtWidgets.QGroupBox("3. Dossier de Destination Parent")
        gd_layout = QtWidgets.QHBoxLayout(group_dest)
        self.txt_dest = QtWidgets.QLineEdit(self.parent_dir)
        btn_browse_dest = QtWidgets.QPushButton("📂 Choisir le dossier...")
        btn_browse_dest.setCursor(QtCore.Qt.PointingHandCursor)
        btn_browse_dest.clicked.connect(self.browse_dest)
        gd_layout.addWidget(self.txt_dest)
        gd_layout.addWidget(btn_browse_dest)
        main_layout.addWidget(group_dest)

        # --- GROUPE 4: Options de formatage ---
        group_opts = QtWidgets.QGroupBox("4. Options de Nommage & Cellule C3")
        go_layout = QtWidgets.QHBoxLayout(group_opts)
        
        # Séparateur de fichier
        sep_layout = QtWidgets.QVBoxLayout()
        sep_title = QtWidgets.QLabel("Format du nom de fichier :")
        sep_title.setStyleSheet("color: #E2E8F0; font-weight: bold;")
        self.rad_sep_underscore = QtWidgets.QRadioButton("Avec tirets bas '_' (1M1_Nom_Prenom.xlsx)")
        self.rad_sep_underscore.setChecked(True)
        self.rad_sep_space = QtWidgets.QRadioButton("Avec espaces ' ' (1M1 Nom Prenom.xlsx)")
        sep_layout.addWidget(sep_title)
        sep_layout.addWidget(self.rad_sep_underscore)
        sep_layout.addWidget(self.rad_sep_space)
        
        # Format C3
        c3_layout = QtWidgets.QVBoxLayout()
        c3_title = QtWidgets.QLabel("Format de la cellule C3 :")
        c3_title.setStyleSheet("color: #E2E8F0; font-weight: bold;")
        self.rad_c3_prenom_nom = QtWidgets.QRadioButton("Prénom Nom (ex: Alice Dupont)")
        self.rad_c3_prenom_nom.setChecked(True)
        self.rad_c3_nom_prenom = QtWidgets.QRadioButton("Nom Prénom (ex: Dupont Alice)")
        c3_layout.addWidget(c3_title)
        c3_layout.addWidget(self.rad_c3_prenom_nom)
        c3_layout.addWidget(self.rad_c3_nom_prenom)

        go_layout.addLayout(sep_layout)
        go_layout.addSpacing(30)
        go_layout.addLayout(c3_layout)
        main_layout.addWidget(group_opts)

        # --- BOUTON DE LANCEMENT & PROGRESSION ---
        launch_layout = QtWidgets.QHBoxLayout()
        self.btn_launch = QtWidgets.QPushButton("🚀 LANCER LA GÉNÉRATION PAR LOTS")
        self.btn_launch.setObjectName("btnLaunch")
        self.btn_launch.setCursor(QtCore.Qt.PointingHandCursor)
        self.btn_launch.clicked.connect(self.start_generation)
        launch_layout.addWidget(self.btn_launch)
        main_layout.addLayout(launch_layout)

        self.progress_bar = QtWidgets.QProgressBar()
        self.progress_bar.setValue(0)
        self.progress_bar.setVisible(False)
        main_layout.addWidget(self.progress_bar)

        # Journal d'exécution (Log Console)
        self.txt_log = QtWidgets.QTextEdit()
        self.txt_log.setReadOnly(True)
        self.txt_log.setMaximumHeight(130)
        self.txt_log.setPlaceholderText("Le journal de génération s'affichera ici...")
        main_layout.addWidget(self.txt_log)

    def auto_detect_files(self):
        """Détecte automatiquement les templates et listes présents dans le dossier courant."""
        all_xlsx = glob.glob("*.xlsx")
        vrais_templates = []
        for f in all_xlsx:
            try:
                wb = openpyxl.load_workbook(f, read_only=True)
                wb.close()
                vrais_templates.append(os.path.abspath(f))
            except Exception:
                pass

        if vrais_templates:
            self.template_path = vrais_templates[0]
            self.txt_template.setText(self.template_path)

        listes = [os.path.abspath(f) for f in glob.glob("*.xlsx") + glob.glob("*.txt") if os.path.abspath(f) != self.template_path]
        if listes:
            self.listes_files = listes
            self.update_listes_widget()

    def browse_template(self):
        filePath, _ = QtWidgets.QFileDialog.getOpenFileName(
            self, "Sélectionnez le fichier template Excel", "", "Fichiers Excel (*.xlsx)"
        )
        if filePath:
            self.template_path = filePath
            self.txt_template.setText(self.template_path)

    def browse_listes(self):
        files, _ = QtWidgets.QFileDialog.getOpenFileNames(
            self, "Sélectionnez une ou plusieurs listes d'élèves", "", "Fichiers Élèves (*.xlsx *.txt)"
        )
        if files:
            for f in files:
                abs_f = os.path.abspath(f)
                if abs_f not in self.listes_files:
                    self.listes_files.append(abs_f)
            self.update_listes_widget()

    def clear_listes(self):
        self.listes_files.clear()
        self.update_listes_widget()

    def update_listes_widget(self):
        self.lst_listes.clear()
        for f in self.listes_files:
            self.lst_listes.addItem(f"📄 {f}")
        self.lbl_listes_count.setText(f"{len(self.listes_files)} classe(s) sélectionnée(s)")

    def browse_dest(self):
        dirPath = QtWidgets.QFileDialog.getExistingDirectory(self, "Sélectionnez le dossier de destination parent")
        if dirPath:
            self.parent_dir = dirPath
            self.txt_dest.setText(self.parent_dir)

    def start_generation(self):
        if not self.template_path or not os.path.exists(self.template_path):
            QtWidgets.QMessageBox.warning(self, "Template manquant", "Veuillez sélectionner un fichier template Excel validé.")
            return
        if not self.listes_files:
            QtWidgets.QMessageBox.warning(self, "Listes manquantes", "Veuillez sélectionner au moins un fichier de liste d'élèves.")
            return

        parent_d = self.txt_dest.text().strip()
        if not parent_d:
            parent_d = os.path.abspath("fichiers_eleves")

        separateur = "_" if self.rad_sep_underscore.isChecked() else " "
        c3_format = "nom_prenom" if self.rad_c3_nom_prenom.isChecked() else "prenom_nom"

        self.btn_launch.setEnabled(False)
        self.progress_bar.setVisible(True)
        self.progress_bar.setValue(0)
        self.txt_log.clear()
        self.txt_log.append("🚀 Démarrage de la génération par lots...\n")

        # Lancer le Thread d'arrière-plan
        self.worker = WorkerThread(self.template_path, self.listes_files, parent_d, separateur, c3_format)
        self.worker.progress_signal.connect(self.on_progress)
        self.worker.finished_signal.connect(self.on_finished)
        self.worker.error_signal.connect(self.on_error)
        self.worker.start()

    def on_progress(self, current, total, message):
        percent = int((current / total) * 100) if total > 0 else 0
        self.progress_bar.setValue(percent)
        self.txt_log.append(message)

    def on_finished(self, total_files, total_classes, parent_dir):
        self.progress_bar.setValue(100)
        self.btn_launch.setEnabled(True)
        self.txt_log.append(f"\n🎉 SUCCÈS TOTAL ! {total_files} fichier(s) généré(s) pour {total_classes} classe(s) !")
        self.txt_log.append(f"📁 Localisation : {parent_dir}")
        
        QtWidgets.QMessageBox.information(
            self, "Génération terminée !", 
            f"🎉 Opération terminée avec succès !\n\n• {total_files} fichier(s) d'évaluation créés\n• {total_classes} classe(s) traitée(s)\n\n📁 Dossier : {parent_dir}"
        )

    def on_error(self, err_msg):
        self.btn_launch.setEnabled(True)
        self.progress_bar.setVisible(False)
        QtWidgets.QMessageBox.critical(self, "Erreur", f"Une erreur s'est produite lors de la génération :\n{err_msg}")

def main():
    app = QtWidgets.QApplication(sys.argv)
    app.setApplicationName("Générateur de Documents Élèves")
    window = MainWindow()
    window.show()
    sys.exit(app.exec())

if __name__ == "__main__":
    main()
