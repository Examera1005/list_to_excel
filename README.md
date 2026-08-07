# 🎓 Générateur de Fichiers Élèves (Template Excel)

Un outil automatique pour générer les fichiers d'évaluation Excel individuels par élève à partir d'une liste de classe (`.xlsx` ou `.txt`) et d'un template (`.xlsx`).

Compatible avec **macOS**, **Linux** et **Windows**.

---

## 🛠️ Fonctionnalités

* 📂 **Lecture automatique des listes d'élèves** : Supporte les fichiers Excel d'export (extraits avec *Groupe / Classe*, *Nom*, *Prénom*) ou les simples fichiers texte `.txt`.
* 🏷️ **Nommage personnalisé des fichiers** : Structure au format `[Classe]_[Nom]_[Prénom].xlsx` (ex: `1M4_Arcan_Danny.xlsx`).
* ✏️ **Mise à jour automatique de la cellule C3** : Remplit la cellule **C3** du template Excel avec le nom complet de l'élève (`Prénom Nom`).
* 🎨 **Conservation intégrale du style** : Préserve les formules, les couleurs, les bordures et les graphiques du fichier template.

---

## 🚀 Utilisation

### Option 1 : Lancement rapide (Automatique)

Pour lancer la duplication automatique basée sur le template et la liste présents dans le dossier :

```bash
python3 dupliquer.py
```

Les 20 fichiers d'élèves seront automatiquement créés dans le dossier `fichiers_eleves/`.

---

### Option 2 : Interface Interactive (TUI)

Pour choisir le template, la liste, le dossier de destination, ainsi que le formatage du nom (avec tiret `_` ou avec espace ` `) et l'ordre dans la cellule C3 :

```bash
python3 app_tui.py
```

---

## 📦 Dépendances Python

Ce projet utilise `openpyxl` pour la manipulation des fichiers `.xlsx` :

```bash
pip install openpyxl
```
