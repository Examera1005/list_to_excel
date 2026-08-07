# 🎓 Générateur de Fichiers Élèves (Template Excel)

Un outil en Python pour générer automatiquement les fichiers d'évaluation Excel individuels par élève à partir de vos listes de classes (`.xlsx` ou `.txt`) et d'un template modèle (`.xlsx`).

Compatible avec **macOS** et **Linux**.

---

## ✨ Nouveautés & Fonctionnalités clés

* 🎯 **Suppression ciblée (`Supprimer la sélection`)** : Supprimez uniquement les classes sélectionnées dans la liste sans devoir tout effacer.
* 🔍 **Recherche dynamique FZF (`🔍 Filtrer...`)** : Filtrez instantanément vos listes de classes en tapant leur nom comme dans `fzf`.
* ⌨️ **Autocomplétion de chemins Zsh/Fish (`Tab`)** : Autocomplétion automatique des fichiers et dossiers lors de la saisie au clavier (`QCompleter` / `readline`).
* 🍏 **Application macOS Natif (`Générateur Élèves.app`)** : Double-cliquable dans le Finder sans ouvrir le terminal (`python3 build_mac_app.py`).
* 🎨 **Design Zinc & Violet Minimaliste** : Haute lisibilité, fort contraste et palette sombre épurée.

---

## 🎨 Options d'utilisation

### 1. Interface Graphique (GUI Moderne) - *Recommandé*

```bash
python3 app_gui.py
```

### 2. Application Mac dédiée (.app)

```bash
python3 build_mac_app.py
```
*(Crée `Générateur Élèves.app` directement double-cliquable dans le Finder).*

---

### 3. Interface Terminal Interactive (TUI)

```bash
python3 app_tui.py
```

---

### 4. Script CLI Rapide

```bash
python3 dupliquer.py
```
