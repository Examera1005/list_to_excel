🇫🇷 **Français** | [🇬🇧 English](README_EN.md)

# 🎓 Générateur de Fichiers Élèves (Template Excel)

Un outil automatique et élégant en Python pour générer les fichiers d'évaluation Excel individuels par élève à partir de vos listes de classes (`.xlsx` ou `.txt`) et d'un fichier modèle (`.xlsx`).

Compatible avec **macOS** et **Linux**.

---

## 📁 Structure du projet

* **`app_gui.py`** : Interface graphique moderne Zinc & Violet (PySide6 avec fallback Tkinter).
* **`build_mac_app.py`** : Générateur d’application native macOS (`.app`) double-cliquable dans le Finder.
* **`app_tui.py`** : Interface interactive dans le terminal (Sélecteur à flèches + Espace).
* **`dupliquer.py`** : Script Python de traitement par lots autonome.
* **`template.xlsx`** : Exemple de fichier modèle Excel.
* **`liste_eleves_1M1.xlsx`** : Exemple de liste d'élèves pour la classe **1M1**.
* **`liste_eleves_1M2.xlsx`** : Exemple de liste d'élèves pour la classe **1M2**.

---

## ✨ Fonctionnalités clés

* 🚀 **Traitement par lots Multi-classes (Batch Mode)** : Traitez plusieurs classes en une seule exécution avec création automatique de sous-dossiers par classe (ex: `./fichiers_eleves/1M1/`, `./fichiers_eleves/1M2/`).
* 🍏 **Application Mac Natif (`.app`)** : Génère `Générateur Élèves.app` utilisable au double-clic dans le Finder sans ouvrir le terminal.
* 🔍 **Recherche et filtre type FZF (`🔍 Filtrer...`)** : Filtrez instantanément vos listes de classes en tapant leur nom.
* ⌨️ **Autocomplétion Zsh / Fish (`Tab`)** : Complétion automatique des chemins de fichiers et dossiers dans les champs de saisie.
* 🎯 **Suppression ciblée (`Supprimer sélection`)** : Retirez uniquement les éléments sélectionnés sans tout effacer.
* 🎨 **Design Zinc & Violet Minimaliste** : Style sombre épuré haute lisibilité, fort contraste, boutons réactifs et effets au survol (Hover).
* 🛡️ **Tolérance aux pannes & Auto-installation** : Gestion des fichiers texte nommés `.xlsx` créés via `nano` et auto-installation des dépendances (`openpyxl`, `PySide6`).

---

## 🚀 Guide d'utilisation

### 1. Application macOS Natif (Recommandé sur Mac)

Pour créer l'application Mac double-cliquable dans le Finder :

```bash
python3 build_mac_app.py
```

Vous obtiendrez le fichier **`Générateur Élèves.app`**. Double-cliquez dessus dans le Finder pour lancer l'application !

---

### 2. Interface Graphique (GUI)

Pour exécuter l'interface graphique directement :

```bash
python3 app_gui.py
```

#### Lancement détaché du terminal :
Si vous voulez fermer le terminal tout en gardant l'application ouverte :

* Sur Mac : `open "Générateur Élèves.app"`
* Sur Linux / Mac : `nohup python3 app_gui.py > /dev/null 2>&1 &`

---

### 3. Interface Terminal Interactive (TUI)

Pour utiliser le menu interactif dans le terminal :

```bash
python3 app_tui.py
```

---

### 4. Mode automatique (Script CLI)

Pour exécuter la génération en ligne de commande :

```bash
python3 dupliquer.py
```
