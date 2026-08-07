# 🎓 Générateur de Fichiers Élèves (Template Excel)

Un outil en Python pour générer automatiquement les fichiers d'évaluation Excel individuels par élève à partir d'une liste de classe (`.xlsx` ou `.txt`) et d'un template modèle (`.xlsx`).

Compatible avec **macOS**, **Linux** et **Windows**.

---

## 📁 Fichiers du projet

* **`dupliquer.py`** : Script Python principal de duplication automatique.
* **`app_tui.py`** : Interface interactive (TUI) guidée pas-à-pas dans le terminal.
* **`template.xlsx`** : Exemple de fichier modèle Excel.
* **`liste_eleves.xlsx`** : Exemple de fichier de liste d'élèves (avec colonnes *Nom*, *Prénom*, *Groupe/Classe*).

---

## ✨ Fonctionnalités

* 📂 **Lecture des listes d'élèves** : Détection automatique des colonnes `Nom`, `Prénom` et `Groupe`/`Classe` depuis un fichier `.xlsx` (ou lecture ligne par ligne d'un fichier `.txt`).
* 🏷️ **Nommage des fichiers** : Format `[Classe]_[Nom]_[Prénom].xlsx` (ex: `1M1_Dupont_Alice.xlsx`).
* ✏️ **Mise à jour de la cellule C3** : Écrit automatiquement le nom complet de l'élève (`Prénom Nom`) dans la cellule **C3** du fichier généré.
* 🎨 **Conservation de la mise en page** : Préserve toutes les formules, calculs, couleurs et bordures du fichier template.

---

## 🚀 Utilisation

### 1. Prérequis

Assurez-vous d'installer la bibliothèque `openpyxl` :

```bash
pip install openpyxl
```

---

### 2. Mode automatique (Script CLI)

Placez votre fichier template (ex: `template.xlsx`) et votre fichier de liste (ex: `liste_eleves.xlsx`) dans le dossier du projet, puis exécutez :

```bash
python3 dupliquer.py
```

Les fichiers individuels seront générés automatiquement dans le dossier **`fichiers_eleves/`**.

---

### 3. Mode interactif (Menu guidé)

Pour personnaliser le formatage (avec tiret `_` ou avec espace ` `), modifier l'ordre dans la cellule C3, ou choisir manuellement les fichiers :

```bash
python3 app_tui.py
```
