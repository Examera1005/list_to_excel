# 🎓 Générateur de Fichiers Élèves (Template Excel)

Un outil élégant et léger en Python pour générer automatiquement les fichiers d'évaluation Excel individuels par élève à partir de vos listes de classes (`.xlsx` ou `.txt`) et d'un template modèle (`.xlsx`).

Compatible avec **macOS** et **Linux**.

---

## 🎨 Options d'utilisation

### 1. Interface Graphique (GUI Moderne) - *Recommandé*

Une interface fenêtre sombre, élégante et très légère (~30 Mo RAM) :

```bash
python3 app_gui.py
```

**Fonctionnalités de l'interface GUI :**
* 📁 **Sélecteur natif Finder (macOS) / Explorateur (Linux)** : Ouvrez la fenêtre système de votre système d'exploitation pour sélectionner directement une ou **plusieurs listes d'élèves en même temps** (Cmd + Clic / Maj + Clic / Glisser-déposer).
* ✨ **Design sombre & effets Hover** : Boutons réactifs et modernes.
* ⚡ **Traitement asynchrone** : Barre de progression et journal en direct sans aucun ralentissement.
* 🛠️ **Installation automatique** : Détecte et propose d'installer `PySide6` si la dépendance est manquante.

---

### 2. Interface Terminal Interactive (TUI)

Un menu interactif guidé dans le terminal avec sélecteur à flèches `⬆️` / `⬇️` et touche `ESPACE` :

```bash
python3 app_tui.py
```

---

### 3. Script CLI Rapide

Pour exécuter la génération automatique directement depuis la ligne de commande :

```bash
python3 dupliquer.py
```

---

## 📁 Structure du projet

* **`app_gui.py`** : Interface graphique fenêtre (GUI PySide6 / Qt).
* **`app_tui.py`** : Interface interactive dans le terminal (TUI).
* **`dupliquer.py`** : Script Python de traitement autonome.
* **`template.xlsx`** : Exemple de fichier modèle Excel.
* **`liste_eleves_1M1.xlsx`** : Exemple de liste d'élèves pour la classe **1M1**.
* **`liste_eleves_1M2.xlsx`** : Exemple de liste d'élèves pour la classe **1M2**.
