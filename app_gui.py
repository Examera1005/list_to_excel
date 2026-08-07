#!/usr/bin/env python3
"""
🎓 Générateur de Fichiers Élèves - Interface Graphique (GUI)
Compatible macOS et Linux.

Double moteur GUI :
- PySide6 (Qt6) par défaut : Look sombre ultra-moderne, arrondis, effets au survol.
- Tkinter (Fallback natif) : Ultra-léger (~15Mo RAM), zéro dépendances si PySide6 est absent.
"""

import os
import sys
import glob
import subprocess
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


# =============================================================================
# MOTEUR 1 : PySide6 (Qt6) - Interface Moderne avec Effets Hover
# =============================================================================
def run_pyside6_app():
    from PySide6 import QtWidgets, QtCore, QtGui

    class WorkerThread(QtCore.QThread):
        progress_signal = QtCore.Signal(int, int, str)
        finished_signal = QtCore.Signal(int, int, str)
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
            self.setWindowTitle("🎓 Générateur de Documents Élèves (PySide6 Mode)")
            self.setMinimumSize(850, 700)
            self.resize(900, 750)

            self.template_path = ""
            self.listes_files = []
            self.parent_dir = os.path.abspath("fichiers_eleves")

            self.init_ui()
            self.auto_detect_files()

        def init_ui(self):
            self.setStyleSheet("""
                QMainWindow { background-color: #0F172A; }
                QWidget { color: #F8FAFC; font-family: system-ui, sans-serif; font-size: 13px; }
                QGroupBox {
                    background-color: #1E293B; border: 1px solid #334155; border-radius: 12px;
                    margin-top: 10px; padding: 15px; font-weight: bold; font-size: 14px; color: #38BDF8;
                }
                QPushButton {
                    background-color: #3B82F6; color: #FFFFFF; border: none; border-radius: 8px;
                    padding: 9px 16px; font-weight: 600;
                }
                QPushButton:hover { background-color: #2563EB; }
                QPushButton:pressed { background-color: #1D4ED8; }
                QPushButton#btnLaunch { background-color: #10B981; font-size: 15px; padding: 12px 24px; font-weight: bold; }
                QPushButton#btnLaunch:hover { background-color: #059669; }
                QLineEdit, QListWidget, QTextEdit {
                    background-color: #0F172A; border: 1px solid #334155; border-radius: 8px; padding: 8px; color: #F8FAFC;
                }
                QProgressBar {
                    background-color: #0F172A; border: 1px solid #334155; border-radius: 8px; text-align: center; color: #F8FAFC; font-weight: bold;
                }
                QProgressBar::chunk { background-color: #10B981; border-radius: 7px; }
            """)

            central_widget = QtWidgets.QWidget()
            self.setCentralWidget(central_widget)
            main_layout = QtWidgets.QVBoxLayout(central_widget)
            main_layout.setContentsMargins(20, 20, 20, 20)
            main_layout.setSpacing(15)

            # Header
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

            # 1. Template
            group_template = QtWidgets.QGroupBox("1. Fichier Modèle Template Excel (.xlsx)")
            gt_layout = QtWidgets.QHBoxLayout(group_template)
            self.txt_template = QtWidgets.QLineEdit()
            self.txt_template.setReadOnly(True)
            btn_browse_tpl = QtWidgets.QPushButton("📁 Parcourir...")
            btn_browse_tpl.setCursor(QtCore.Qt.PointingHandCursor)
            btn_browse_tpl.clicked.connect(self.browse_template)
            gt_layout.addWidget(self.txt_template)
            gt_layout.addWidget(btn_browse_tpl)
            main_layout.addWidget(group_template)

            # 2. Listes
            group_listes = QtWidgets.QGroupBox("2. Listes de Classes (.xlsx ou .txt)")
            gl_layout = QtWidgets.QVBoxLayout(group_listes)
            gl_top = QtWidgets.QHBoxLayout()
            self.lbl_listes_count = QtWidgets.QLabel("0 classe(s) sélectionnée(s)")
            btn_browse_listes = QtWidgets.QPushButton("👥 Parcourir les listes (Finder/Explorateur)...")
            btn_browse_listes.setCursor(QtCore.Qt.PointingHandCursor)
            btn_browse_listes.clicked.connect(self.browse_listes)
            btn_clear_listes = QtWidgets.QPushButton("🗑️ Vider")
            btn_clear_listes.setStyleSheet("background-color: #EF4444;")
            btn_clear_listes.setCursor(QtCore.Qt.PointingHandCursor)
            btn_clear_listes.clicked.connect(self.clear_listes)
            gl_top.addWidget(self.lbl_listes_count)
            gl_top.addStretch()
            gl_top.addWidget(btn_browse_listes)
            gl_top.addWidget(btn_clear_listes)
            self.lst_listes = QtWidgets.QListWidget()
            self.lst_listes.setMaximumHeight(110)
            gl_layout.addLayout(gl_top)
            gl_layout.addWidget(self.lst_listes)
            main_layout.addWidget(group_listes)

            # 3. Destination
            group_dest = QtWidgets.QGroupBox("3. Dossier de Destination Parent")
            gd_layout = QtWidgets.QHBoxLayout(group_dest)
            self.txt_dest = QtWidgets.QLineEdit(self.parent_dir)
            btn_browse_dest = QtWidgets.QPushButton("📂 Choisir...")
            btn_browse_dest.setCursor(QtCore.Qt.PointingHandCursor)
            btn_browse_dest.clicked.connect(self.browse_dest)
            gd_layout.addWidget(self.txt_dest)
            gd_layout.addWidget(btn_browse_dest)
            main_layout.addWidget(group_dest)

            # 4. Options
            group_opts = QtWidgets.QGroupBox("4. Options de Nommage & Cellule C3")
            go_layout = QtWidgets.QHBoxLayout(group_opts)
            sep_layout = QtWidgets.QVBoxLayout()
            sep_layout.addWidget(QtWidgets.QLabel("Format du nom de fichier :"))
            self.rad_sep_underscore = QtWidgets.QRadioButton("Avec tirets bas '_' (1M1_Nom_Prenom.xlsx)")
            self.rad_sep_underscore.setChecked(True)
            self.rad_sep_space = QtWidgets.QRadioButton("Avec espaces ' ' (1M1 Nom Prenom.xlsx)")
            sep_layout.addWidget(self.rad_sep_underscore)
            sep_layout.addWidget(self.rad_sep_space)
            
            c3_layout = QtWidgets.QVBoxLayout()
            c3_layout.addWidget(QtWidgets.QLabel("Format de la cellule C3 :"))
            self.rad_c3_prenom_nom = QtWidgets.QRadioButton("Prénom Nom (ex: Alice Dupont)")
            self.rad_c3_prenom_nom.setChecked(True)
            self.rad_c3_nom_prenom = QtWidgets.QRadioButton("Nom Prénom (ex: Dupont Alice)")
            c3_layout.addWidget(self.rad_c3_prenom_nom)
            c3_layout.addWidget(self.rad_c3_nom_prenom)

            go_layout.addLayout(sep_layout)
            go_layout.addSpacing(30)
            go_layout.addLayout(c3_layout)
            main_layout.addWidget(group_opts)

            # Launch & Progress
            self.btn_launch = QtWidgets.QPushButton("🚀 LANCER LA GÉNÉRATION PAR LOTS")
            self.btn_launch.setObjectName("btnLaunch")
            self.btn_launch.setCursor(QtCore.Qt.PointingHandCursor)
            self.btn_launch.clicked.connect(self.start_generation)
            main_layout.addWidget(self.btn_launch)

            self.progress_bar = QtWidgets.QProgressBar()
            self.progress_bar.setVisible(False)
            main_layout.addWidget(self.progress_bar)

            self.txt_log = QtWidgets.QTextEdit()
            self.txt_log.setReadOnly(True)
            self.txt_log.setMaximumHeight(130)
            main_layout.addWidget(self.txt_log)

        def auto_detect_files(self):
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
            filePath, _ = QtWidgets.QFileDialog.getOpenFileName(self, "Sélectionnez le template Excel", "", "Fichiers Excel (*.xlsx)")
            if filePath:
                self.template_path = filePath
                self.txt_template.setText(self.template_path)

        def browse_listes(self):
            files, _ = QtWidgets.QFileDialog.getOpenFileNames(self, "Sélectionnez les listes d'élèves", "", "Listes d'élèves (*.xlsx *.txt)")
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
            dirPath = QtWidgets.QFileDialog.getExistingDirectory(self, "Sélectionnez le dossier parent")
            if dirPath:
                self.parent_dir = dirPath
                self.txt_dest.setText(self.parent_dir)

        def start_generation(self):
            if not self.template_path or not os.path.exists(self.template_path):
                QtWidgets.QMessageBox.warning(self, "Template manquant", "Veuillez sélectionner un fichier template Excel.")
                return
            if not self.listes_files:
                QtWidgets.QMessageBox.warning(self, "Listes manquantes", "Veuillez sélectionner au moins une liste d'élèves.")
                return
            parent_d = self.txt_dest.text().strip() or os.path.abspath("fichiers_eleves")
            separateur = "_" if self.rad_sep_underscore.isChecked() else " "
            c3_format = "nom_prenom" if self.rad_c3_nom_prenom.isChecked() else "prenom_nom"

            self.btn_launch.setEnabled(False)
            self.progress_bar.setVisible(True)
            self.progress_bar.setValue(0)
            self.txt_log.clear()
            self.txt_log.append("🚀 Démarrage de la génération par lots...\n")

            self.worker = WorkerThread(self.template_path, self.listes_files, parent_d, separateur, c3_format)
            self.worker.progress_signal.connect(self.on_progress)
            self.worker.finished_signal.connect(self.on_finished)
            self.worker.error_signal.connect(self.on_error)
            self.worker.start()

        def on_progress(self, current, total, message):
            self.progress_bar.setValue(int((current / total) * 100) if total > 0 else 0)
            self.txt_log.append(message)

        def on_finished(self, total_files, total_classes, parent_dir):
            self.progress_bar.setValue(100)
            self.btn_launch.setEnabled(True)
            self.txt_log.append(f"\n🎉 SUCCÈS TOTAL ! {total_files} fichier(s) généré(s) pour {total_classes} classe(s) !")
            QtWidgets.QMessageBox.information(self, "Succès !", f"🎉 {total_files} fichier(s) généré(s) pour {total_classes} classe(s) !\n\n📁 Dossier : {parent_dir}")

        def on_error(self, err_msg):
            self.btn_launch.setEnabled(True)
            self.progress_bar.setVisible(False)
            QtWidgets.QMessageBox.critical(self, "Erreur", f"Erreur :\n{err_msg}")

    app = QtWidgets.QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())


# =============================================================================
# MOTEUR 2 : Tkinter (Fallback Natif) - Hyper Léger (~15Mo RAM) avec Hover
# =============================================================================
def run_tkinter_app():
    import tkinter as tk
    from tkinter import filedialog, messagebox, ttk
    import threading

    class TkApp(tk.Tk):
        def __init__(self):
            super().__init__()
            self.title("🎓 Générateur de Documents Élèves (Tkinter Natif)")
            self.geometry("850x700")
            self.configure(bg="#0F172A")

            self.template_path = ""
            self.listes_files = []
            self.parent_dir = os.path.abspath("fichiers_eleves")

            self.init_ui()
            self.auto_detect_files()

        def make_hover_button(self, parent, text, command, bg="#3B82F6", hover_bg="#2563EB", fg="#FFFFFF", **kwargs):
            btn = tk.Button(parent, text=text, command=command, bg=bg, fg=fg, activebackground=hover_bg,
                            activeforeground="#FFFFFF", bd=0, relief="flat", font=("Segoe UI", 10, "bold"), cursor="hand2", **kwargs)
            btn.bind("<Enter>", lambda e: btn.config(bg=hover_bg))
            btn.bind("<Leave>", lambda e: btn.config(bg=bg))
            return btn

        def init_ui(self):
            main_frame = tk.Frame(self, bg="#0F172A", padx=20, pady=20)
            main_frame.pack(fill="both", expand=True)

            # Header
            header_frame = tk.Frame(main_frame, bg="#0F172A")
            header_frame.pack(fill="x", pady=(0, 15))
            lbl_title = tk.Label(header_frame, text="🎓 Générateur de Documents Élèves", font=("Segoe UI", 18, "bold"), bg="#0F172A", fg="#F8FAFC")
            lbl_title.pack(anchor="w")
            lbl_sub = tk.Label(header_frame, text="Mode Natif Tkinter (~15Mo RAM) | Duplication par classe", font=("Segoe UI", 10), bg="#0F172A", fg="#94A3B8")
            lbl_sub.pack(anchor="w")

            # 1. Template Group
            f1 = tk.LabelFrame(main_frame, text=" 1. Fichier Modèle Template Excel (.xlsx) ", font=("Segoe UI", 11, "bold"), bg="#1E293B", fg="#38BDF8", padx=15, pady=12, bd=1, relief="solid")
            f1.pack(fill="x", pady=6)
            self.entry_tpl = tk.Entry(f1, bg="#0F172A", fg="#F8FAFC", insertbackground="white", bd=1, relief="solid", font=("Segoe UI", 10))
            self.entry_tpl.pack(side="left", fill="x", expand=True, padx=(0, 10), ipady=4)
            btn_tpl = self.make_hover_button(f1, "📁 Parcourir...", self.browse_template, padx=12, pady=4)
            btn_tpl.pack(side="right")

            # 2. Listes Group
            f2 = tk.LabelFrame(main_frame, text=" 2. Listes de Classes (.xlsx ou .txt) ", font=("Segoe UI", 11, "bold"), bg="#1E293B", fg="#38BDF8", padx=15, pady=12, bd=1, relief="solid")
            f2.pack(fill="x", pady=6)
            top_l = tk.Frame(f2, bg="#1E293B")
            top_l.pack(fill="x", pady=(0, 6))
            self.lbl_count = tk.Label(top_l, text="0 classe(s) sélectionnée(s)", font=("Segoe UI", 10), bg="#1E293B", fg="#94A3B8")
            self.lbl_count.pack(side="left")
            btn_clear = self.make_hover_button(top_l, "🗑️ Vider", self.clear_listes, bg="#EF4444", hover_bg="#DC2626", padx=8, pady=3)
            btn_clear.pack(side="right", padx=(6, 0))
            btn_listes = self.make_hover_button(top_l, "👥 Parcourir les listes (Finder/Explorateur)...", self.browse_listes, padx=12, pady=3)
            btn_listes.pack(side="right")
            
            self.listbox = tk.Listbox(f2, bg="#0F172A", fg="#F8FAFC", selectbackground="#2563EB", height=4, bd=1, relief="solid", font=("Segoe UI", 10))
            self.listbox.pack(fill="x")

            # 3. Destination Group
            f3 = tk.LabelFrame(main_frame, text=" 3. Dossier de Destination Parent ", font=("Segoe UI", 11, "bold"), bg="#1E293B", fg="#38BDF8", padx=15, pady=12, bd=1, relief="solid")
            f3.pack(fill="x", pady=6)
            self.entry_dest = tk.Entry(f3, bg="#0F172A", fg="#F8FAFC", insertbackground="white", bd=1, relief="solid", font=("Segoe UI", 10))
            self.entry_dest.insert(0, self.parent_dir)
            self.entry_dest.pack(side="left", fill="x", expand=True, padx=(0, 10), ipady=4)
            btn_dest = self.make_hover_button(f3, "📂 Choisir...", self.browse_dest, padx=12, pady=4)
            btn_dest.pack(side="right")

            # 4. Options Group
            f4 = tk.LabelFrame(main_frame, text=" 4. Options de Nommage & Cellule C3 ", font=("Segoe UI", 11, "bold"), bg="#1E293B", fg="#38BDF8", padx=15, pady=12, bd=1, relief="solid")
            f4.pack(fill="x", pady=6)
            
            self.var_sep = tk.StringVar(value="_")
            lbl_sep = tk.Label(f4, text="Nom de fichier :", bg="#1E293B", fg="#E2E8F0", font=("Segoe UI", 10, "bold"))
            lbl_sep.grid(row=0, column=0, sticky="w", padx=5)
            r1 = tk.Radiobutton(f4, text="Avec tiret '_' (1M1_Nom_Prenom.xlsx)", variable=self.var_sep, value="_", bg="#1E293B", fg="#CBD5E1", selectcolor="#0F172A", activebackground="#1E293B")
            r1.grid(row=1, column=0, sticky="w", padx=5)
            r2 = tk.Radiobutton(f4, text="Avec espace ' ' (1M1 Nom Prenom.xlsx)", variable=self.var_sep, value=" ", bg="#1E293B", fg="#CBD5E1", selectcolor="#0F172A", activebackground="#1E293B")
            r2.grid(row=2, column=0, sticky="w", padx=5)

            self.var_c3 = tk.StringVar(value="prenom_nom")
            lbl_c3 = tk.Label(f4, text="Format cellule C3 :", bg="#1E293B", fg="#E2E8F0", font=("Segoe UI", 10, "bold"))
            lbl_c3.grid(row=0, column=1, sticky="w", padx=(30, 5))
            r3 = tk.Radiobutton(f4, text="Prénom Nom (ex: Alice Dupont)", variable=self.var_c3, value="prenom_nom", bg="#1E293B", fg="#CBD5E1", selectcolor="#0F172A", activebackground="#1E293B")
            r3.grid(row=1, column=1, sticky="w", padx=(30, 5))
            r4 = tk.Radiobutton(f4, text="Nom Prénom (ex: Dupont Alice)", variable=self.var_c3, value="nom_prenom", bg="#1E293B", fg="#CBD5E1", selectcolor="#0F172A", activebackground="#1E293B")
            r4.grid(row=2, column=1, sticky="w", padx=(30, 5))

            # Launch Button
            self.btn_launch = self.make_hover_button(main_frame, "🚀 LANCER LA GÉNÉRATION PAR LOTS", self.start_generation, bg="#10B981", hover_bg="#059669", pady=8)
            self.btn_launch.pack(fill="x", pady=10)

            # Console Log
            self.txt_log = tk.Text(main_frame, bg="#0F172A", fg="#34D399", height=6, font=("Consolas", 9), bd=1, relief="solid")
            self.txt_log.pack(fill="both", expand=True)

        def auto_detect_files(self):
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
                self.entry_tpl.delete(0, tk.END)
                self.entry_tpl.insert(0, self.template_path)
            listes = [os.path.abspath(f) for f in glob.glob("*.xlsx") + glob.glob("*.txt") if os.path.abspath(f) != self.template_path]
            if listes:
                self.listes_files = listes
                self.update_listbox()

        def browse_template(self):
            path = filedialog.askopenfilename(title="Sélectionner le template Excel", filetypes=[("Excel Files", "*.xlsx")])
            if path:
                self.template_path = path
                self.entry_tpl.delete(0, tk.END)
                self.entry_tpl.insert(0, self.template_path)

        def browse_listes(self):
            files = filedialog.askopenfilenames(title="Sélectionner les listes d'élèves", filetypes=[("Listes Élèves", "*.xlsx *.txt")])
            if files:
                for f in files:
                    abs_f = os.path.abspath(f)
                    if abs_f not in self.listes_files:
                        self.listes_files.append(abs_f)
                self.update_listbox()

        def clear_listes(self):
            self.listes_files.clear()
            self.update_listbox()

        def update_listbox(self):
            self.listbox.delete(0, tk.END)
            for f in self.listes_files:
                self.listbox.insert(tk.END, f"📄 {f}")
            self.lbl_count.config(text=f"{len(self.listes_files)} classe(s) sélectionnée(s)")

        def browse_dest(self):
            path = filedialog.askdirectory(title="Sélectionner le dossier parent")
            if path:
                self.parent_dir = path
                self.entry_dest.delete(0, tk.END)
                self.entry_dest.insert(0, self.parent_dir)

        def log_msg(self, msg):
            self.txt_log.insert(tk.END, msg + "\n")
            self.txt_log.see(tk.END)

        def start_generation(self):
            if not self.template_path or not os.path.exists(self.template_path):
                messagebox.showwarning("Avertissement", "Veuillez sélectionner un fichier template Excel.")
                return
            if not self.listes_files:
                messagebox.showwarning("Avertissement", "Veuillez sélectionner au moins un fichier de liste d'élèves.")
                return

            self.btn_launch.config(state="disabled")
            self.txt_log.delete("1.0", tk.END)
            self.log_msg("🚀 Démarrage de la génération par lots...")

            def worker_task():
                try:
                    parent_d = self.entry_dest.get().strip() or os.path.abspath("fichiers_eleves")
                    separateur = self.var_sep.get()
                    c3_format = self.var_c3.get()

                    total_fichiers = 0
                    all_data = []
                    for l_file in self.listes_files:
                        eleves = parse_student_list(l_file)
                        if eleves:
                            all_data.append((l_file, eleves))

                    total_items = sum(len(e) for _, e in all_data)
                    current_item = 0

                    for l_file, eleves in all_data:
                        nom_classe = eleves[0]["classe"] if eleves[0]["classe"] else os.path.splitext(os.path.basename(l_file))[0]
                        target_dir = os.path.join(parent_d, nom_classe) if len(all_data) > 1 else parent_d
                        os.makedirs(target_dir, exist_ok=True)

                        for e in eleves:
                            nom = e["nom"]
                            prenom = e["prenom"]
                            classe = e["classe"] if e["classe"] else nom_classe
                            parts = [p for p in [classe, nom, prenom] if p]
                            
                            nom_f = ("_".join(parts).replace(" ", "_") if separateur == "_" else " ".join(parts)) + ".xlsx"
                            c3_val = f"{nom} {prenom}" if c3_format == "nom_prenom" else f"{prenom} {nom}"

                            destination = os.path.join(target_dir, nom_f)
                            
                            wb = openpyxl.load_workbook(self.template_path)
                            ws = wb.active
                            ws['C3'] = c3_val.strip()
                            wb.save(destination)

                            current_item += 1
                            total_fichiers += 1
                            m = f"✅ [{current_item}/{total_items}] [{nom_classe}] Généré : {nom_f} (C3 = '{c3_val.strip()}')"
                            self.after(0, self.log_msg, m)

                    self.after(0, self.on_finish, total_fichiers, len(all_data), os.path.abspath(parent_d))
                except Exception as err:
                    self.after(0, lambda: messagebox.showerror("Erreur", str(err)))
                    self.after(0, lambda: self.btn_launch.config(state="normal"))

            threading.Thread(target=worker_task, daemon=True).start()

        def on_finish(self, total_files, total_classes, parent_dir):
            self.btn_launch.config(state="normal")
            self.log_msg(f"\n🎉 SUCCÈS TOTAL ! {total_files} fichier(s) généré(s) pour {total_classes} classe(s) !")
            messagebox.showinfo("Succès !", f"🎉 {total_files} fichier(s) généré(s) pour {total_classes} classe(s) !\n\n📁 Dossier : {parent_dir}")

    app = TkApp()
    app.mainloop()


# =============================================================================
# MAIN : DÉTECTION DU MOTEUR DISPONIBLE
# =============================================================================
def main():
    # 1. Tenter de lancer PySide6 (Look Moderne Qt6)
    try:
        import PySide6
        run_pyside6_app()
        return
    except ImportError:
        pass

    # 2. Tenter de lancer Tkinter (Natif ultra-léger)
    try:
        import tkinter
        run_tkinter_app()
        return
    except ImportError:
        pass

    # 3. Si aucun moteur n'est présent, installer PySide6
    print("===============================================================")
    print(" ⚠️ DÉPENDANCE GUI MANQUANTE")
    print("===============================================================")
    print("L'interface graphique nécessite 'PySide6' ou 'tkinter'.\n")
    
    choix = input("👉 Souhaitez-vous installer PySide6 automatiquement ? [O/n] : ").strip().lower()
    if choix in ['', 'o', 'oui', 'y', 'yes']:
        try:
            subprocess.check_call([sys.executable, "-m", "pip", "install", "PySide6"])
        except Exception:
            subprocess.check_call([sys.executable, "-m", "pip", "install", "PySide6", "--break-system-packages"])
        run_pyside6_app()
    else:
        sys.exit(1)

if __name__ == "__main__":
    main()
