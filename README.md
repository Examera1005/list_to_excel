# 🎓 Générateur de Fichiers Élèves (Template Excel)

Un outil en Python pour générer automatiquement les fichiers d'évaluation Excel individuels par élève à partir de vos listes de classes (`.xlsx` ou `.txt`) et d'un template modèle (`.xlsx`).

Compatible avec **macOS**, **Linux** et **Windows**.

---

## 📁 Fichiers du projet

* **`dupliquer.py`** : Script Python de duplication automatique rapide avec support multi-classes.
* **`app_tui.py`** : Interface interactive (TUI) guidée (Sélecteur Flèches + Espace, mode multi-classes).
* **`template.xlsx`** : Exemple de fichier modèle Excel.
* **`liste_eleves_1M1.xlsx`** : Exemple de liste d'élèves pour la classe **1M1**.
* **`liste_eleves_1M2.xlsx`** : Exemple de liste d'élèves pour la classe **1M2**.

---

## ✨ Fonctionnalités clés

* 🚀 **Traitement par lots Multi-classes (Batch Mode)** : Traitez plusieurs classes d'un coup en **1 seule exécution** sans copier le script ni le relancer !
* 📂 **Organisation automatique en sous-dossiers** : Le script crée automatiquement un sous-dossier propre par classe (ex: `./fichiers_eleves/1M1/`, `./fichiers_eleves/1M2/`...).
* 🏷️ **Nommage personnalisé** : Format `[Classe]_[Nom]_[Prénom].xlsx` (ex: `1M1_Dupont_Alice.xlsx`).
* ✏️ **Cellule C3 automatique** : Écrit le nom complet (`Prénom Nom`) dans la cellule **C3** de chaque fichier.
* 🎮 **Sélecteur interactif Flèches + Espace** : Cochez/déconnectez les listes de classe facilement avec les flèches du clavier et la barre d'espace.

---

## 🚀 Utilisation

### Mode 1 : Interface interactive (Recommandé)

```bash
python3 app_tui.py
```

---

### Mode 2 : Lancement automatique en ligne de commande

```bash
python3 dupliquer.py
```
