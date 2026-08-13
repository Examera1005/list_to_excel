[🇫🇷 Français](README.md) | 🇬🇧 **English**

# 🎓 Student File Generator (Excel Template)

An automatic and elegant Python tool to generate individual Excel evaluation files for each student from your class lists (`.xlsx` or `.txt`) and a template file (`.xlsx`).

Compatible with **macOS** and **Linux**.

---

## 📁 Project Structure

* **`app_gui.py`**: Modern Zinc & Violet graphical user interface (PySide6 with Tkinter fallback).
* **`build_mac_app.py`**: macOS native application builder (`.app`), double-clickable in Finder.
* **`app_tui.py`**: Interactive terminal interface (Arrow keys + Space selector).
* **`dupliquer.py`**: Standalone batch processing Python script.
* **`template.xlsx`**: Sample Excel template file.
* **`liste_eleves_1M1.xlsx`**: Sample student list for class **1M1**.
* **`liste_eleves_1M2.xlsx`**: Sample student list for class **1M2**.

---

## ✨ Key Features

* 🚀 **Multi-class Batch Mode**: Process multiple classes in a single run with automatic creation of subfolders per class (e.g., `./fichiers_eleves/1M1/`, `./fichiers_eleves/1M2/`).
* 🍏 **Native Mac Application (`.app`)**: Generates `Générateur Élèves.app`, launchable with a double-click in Finder without opening a terminal.
* 🔍 **FZF-style Search & Filter (`🔍 Filter...`)**: Instantly filter your class lists by typing their name.
* ⌨️ **Zsh / Fish Autocompletion (`Tab`)**: Automatic completion of file and directory paths in input fields.
* 🎯 **Targeted Deletion (`Remove selected`)**: Remove only selected items without clearing everything.
* 🎨 **Minimalist Zinc & Violet Design**: Clean dark theme, high legibility, strong contrast, responsive buttons, and hover effects.
* 🛡️ **Fault Tolerance & Auto-installation**: Gracefully handles plain-text files named `.xlsx` created via `nano`, and auto-installs dependencies (`openpyxl`, `PySide6`).

---

## 🚀 User Guide

### 1. Native macOS Application (Recommended on Mac)

To generate the double-clickable Mac app bundle in Finder:

```bash
python3 build_mac_app.py
```

You will get the file **`Générateur Élèves.app`**. Double-click it in Finder to launch the application!

---

### 2. Graphical User Interface (GUI)

To run the graphical interface directly:

```bash
python3 app_gui.py
```

#### Detached Launch from Terminal:
If you want to close the terminal while keeping the app running:

* On Mac: `open "Générateur Élèves.app"`
* On Linux / Mac: `nohup python3 app_gui.py > /dev/null 2>&1 &`

---

### 3. Interactive Terminal Interface (TUI)

To use the interactive terminal menu:

```bash
python3 app_tui.py
```

---

### 4. Automatic Mode (CLI Script)

To run the generation directly from the command line:

```bash
python3 dupliquer.py
```
