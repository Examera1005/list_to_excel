# 🎓 Générateur de Fichiers Élèves (Template Excel)

Un outil en Python pour générer automatiquement les fichiers d'évaluation Excel individuels par élève à partir d'une liste de classe (`.xlsx` ou `.txt`) et d'un template modèle (`.xlsx`).

Compatible avec **macOS**, **Linux** et **Windows**.

---

## 📁 Fichiers du projet

* **`dupliquer.py`** : Script Python de duplication automatique rapide.
* **`app_tui.py`** : Interface interactive (TUI) guidée avec gestionnaire multi-dossiers et validation des saisies.
* **`template.xlsx`** : Exemple de fichier modèle Excel.
* **`liste_eleves.xlsx`** : Exemple de fichier de liste d'élèves (*Nom*, *Prénom*, *Groupe/Classe*).

---

## ✨ Fonctionnalités avancées

* 📂 **Lecture des listes d'élèves** : Détection automatique des colonnes `Nom`, `Prénom` et `Groupe`/`Classe` depuis un fichier `.xlsx` ou `.txt`.
* 🏷️ **Nommage des fichiers** : Format `[Classe]_[Nom]_[Prénom].xlsx` (ex: `1M1_Dupont_Alice.xlsx`).
* ✏️ **Mise à jour de la cellule C3** : Remplit automatiquement le nom complet de l'élève (`Prénom Nom`) dans la cellule **C3**.
* 🛡️ **Validation stricte des saisies** : En cas de réponse invalide dans le TUI (ex: `lol`), le programme indique l'erreur et redemande la saisie au lieu de choisir par défaut.
* 📁 **Navigateur et Sélection MULTIPLE de dossiers** : Choisissez un ou **plusieurs dossiers/sous-dossiers de destination** en même temps via un navigateur interactif avec cases à cocher `[✓]`.
* 📦 **Installation automatique des dépendances** : Détecte et propose d'installer `openpyxl` si la bibliothèque est manquante.

---

## 🚀 Utilisation

### 1. Prérequis

Assurez-vous d'installer la bibliothèque `openpyxl` (le script vous proposera de l'installer automatiquement si nécessaire) :

```bash
pip install openpyxl
```

---

### 2. Mode interactif avancé (Recommandé)

Permet de naviguer dans les sous-dossiers, de cocher plusieurs répertoires cibles et de personnaliser les options :

```bash
python3 app_tui.py
```

---

### 3. Mode automatique rapide (CLI)

```bash
python3 dupliquer.py
```
